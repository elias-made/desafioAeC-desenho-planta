"""
Space Planner — AeC
====================
Fluxo:
  1. Lê premissas.txt
  2. Lê planta.xlsx (aba JPIII)
  3. Gera mapa 2D + resumo de zonas para o LLM
  4. PlannerAgent orquestra:
       - Decide o que liberar
       - Para cada novo espaço: chama BlockDesignerAgent via tool criar_espaco
  5. Código executa as mudanças na planilha clonada
  6. Salva Excel + relatório TXT em ./propostas/
"""

import asyncio
import json
import os
import re
from copy import copy

import openpyxl
from dotenv import load_dotenv
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from Agents import PlannerDeps, planner_agent, BlocoLayout
from BlockMapper import scan_plant, describe_for_llm
from mapa_reduzido import scan as scan_map, render, find_region

load_dotenv()

# ── Configuração ─────────────────────────────────────────────────────────
SHEET_NAME   = 'JPIII'
CORRIDOR_GAP = 3
MAP_PADDING  = 3
MAP_TOP_N    = 2

FORBIDDEN_PATTERNS = {
    'SALA 1', 'SALA1', 'SALA 2', 'SALA2',
    'SALA 3', 'SALA3', 'SALA 4', 'SALA4',
    'COWORKING', 'SALA CLIENTE',
}

_PALETTE = [
    ('1E8449', '145A32'),
    ('CB4335', '922B21'),
    ('7D3C98', '6C3483'),
    ('1A5276', '154360'),
    ('B7770D', '9A7D0A'),
    ('117A65', '0E6655'),
]

def _fill(hex_color: str) -> PatternFill:
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type='solid')

FILL_LIBERADO = _fill('BDC3C7')
FILL_CATRACA  = _fill('F39C12')
FONT_WHITE    = Font(color='FFFFFF', bold=True, size=8)
FONT_SMALL    = Font(size=8)


# ── Carregamento ──────────────────────────────────────────────────────────

def load_plant(path='planta.xlsx', sheet=SHEET_NAME):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    print(f"  Planta: '{ws.title}' — {ws.max_row}L × {ws.max_column}C")
    return wb, ws


def clone_ws(ws):
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            src = ws.cell(r, c)
            tgt = new_ws.cell(r, c)
            tgt.value = src.value
            if src.has_style:
                tgt.font      = copy(src.font)
                tgt.fill      = copy(src.fill)
                tgt.alignment = copy(src.alignment)
                tgt.border    = copy(src.border)
    for col_l, dim in ws.column_dimensions.items():
        new_ws.column_dimensions[col_l].width = dim.width
    for row_i, dim in ws.row_dimensions.items():
        new_ws.row_dimensions[row_i].height = dim.height
    if ws.freeze_panes:
        new_ws.freeze_panes = ws.freeze_panes
    return new_wb, new_ws


# ── Contexto para o LLM ───────────────────────────────────────────────────

def build_plant_info(plant_data: dict) -> str:
    client_cells = plant_data['client_cells']
    lines = ["CÉLULAS POR VALOR NA PLANTA:"]
    for val in sorted(client_cells, key=lambda v: len(client_cells[v]), reverse=True)[:20]:
        lines.append(f"  '{val}': {len(client_cells[val])} PAs")
    lines += ["", "ZONAS DOS CLIENTES MAIS POPULOSOS:"]
    for val in sorted(client_cells, key=lambda v: len(client_cells[v]), reverse=True)[:3]:
        lines.append(describe_for_llm(val, client_cells[val], CORRIDOR_GAP))
    lines += ["", f"ÁREAS PROIBIDAS: {len(plant_data['forbidden'])} células (SALA 1-4, COWORKING)"]
    return "\n".join(lines)


def build_mapa_2d(ws) -> str:
    cells, client_syms = scan_map(ws)
    skip   = {'CT', 'CW', 'SA', '##', '  '}
    ranked = sorted(
        [(sym, lst) for sym, lst in client_syms.items()
         if sym not in skip and len(lst) > 5],
        key=lambda x: len(x[1]), reverse=True,
    )
    targets  = [sym for sym, _ in ranked[:MAP_TOP_N]]
    sections = [
        "LEGENDA: [XX]=célula  [  ]=vazio  [CT]=catraca  [SA]=sala",
        "PAREDES: thick vermelho = parede de ambiente",
        "",
    ]
    for target in targets:
        cell_list = client_syms.get(target, [])
        if not cell_list:
            continue
        r_min, r_max, c_min, c_max = find_region(
            cell_list, MAP_PADDING, ws.max_row, ws.max_column
        )
        if c_max - c_min > 40:
            from BlockMapper import flood_fill, group_zones
            zones   = group_zones(flood_fill(set(cell_list)), gap=3)
            biggest = zones[0] if zones else cell_list
            r_min, r_max, c_min, c_max = find_region(
                biggest, MAP_PADDING, ws.max_row, ws.max_column
            )
        grid_lines, v_walls, h_walls, ct_desc = render(cells, r_min, r_max, c_min, c_max)
        sections.append(
            f"── CLIENTE '{target}' ({len(cell_list)} PAs) "
            f"— {get_column_letter(c_min)}{r_min}:{get_column_letter(c_max)}{r_max} ──"
        )
        sections.extend(grid_lines)
        sections += ["PAREDES VERTICAIS:"] + (v_walls or ["  (nenhuma)"])
        sections += ["PAREDES HORIZONTAIS:"] + (h_walls or ["  (nenhuma)"])
        sections += ["CATRACAS:"] + (ct_desc or ["  (nenhuma)"])
        sections.append("")
    return "\n".join(sections)


# ── Canvas ────────────────────────────────────────────────────────────────

def build_canvas_map(premissas: str) -> dict:
    """
    Lê as premissas, detecta pedidos de criação de espaço e pré-aloca
    canvas (linha/coluna de início) para cada um.

    Retorna canvas_map: {nome: {row_start, col_start}}
    Os nomes são ESP-A, ESP-B, ... na ordem em que aparecem.
    """
    ROW_BASE  = 125
    COL_START = 2
    col_cursor = COL_START
    canvas_map = {}

    for line in premissas.splitlines():
        line = line.strip()
        m_pas  = re.search(r'(\d+)\s*PAs?', line, re.IGNORECASE)
        if not m_pas or not re.search(r'criar|novo espaço', line, re.IGNORECASE):
            continue
        idx  = len(canvas_map) + 1
        nome = f'ESP-{chr(64 + idx)}'   # ESP-A, ESP-B, ...
        n_pas = int(m_pas.group(1))

        # Estimativa de largura para espaçar os blocos
        n_ilhas   = -(-n_pas // 32)
        est_width = n_ilhas * 3 + 2    # ilhas + sala

        canvas_map[nome] = {
            'row_start': ROW_BASE,
            'col_start': col_cursor,
        }
        col_cursor += est_width + 4    # margem entre blocos

    return canvas_map


# ── Paredes ────────────────────────────────────────────────────────────────

def draw_walls_around_block(ws, cells: list):
    """
    Aplica bordas thick vermelhas no perímetro externo contínuo do bounding box.
    Itera sobre TODAS as células do bounding box (incluindo corredores).
    """
    if not cells:
        return
    wall  = Side(border_style='thick', color='FFFF0000')
    rows  = [r for r, c in cells]
    cols  = [c for r, c in cells]
    r_min, r_max = min(rows), max(rows)
    c_min, c_max = min(cols), max(cols)

    for c in range(c_min, c_max + 1):
        cell = ws.cell(r_min, c)
        b = copy(cell.border)
        cell.border = Border(top=wall, bottom=b.bottom, left=b.left, right=b.right)

    for c in range(c_min, c_max + 1):
        cell = ws.cell(r_max, c)
        b = copy(cell.border)
        cell.border = Border(top=b.top, bottom=wall, left=b.left, right=b.right)

    for r in range(r_min, r_max + 1):
        cell = ws.cell(r, c_min)
        b = copy(cell.border)
        cell.border = Border(top=b.top, bottom=b.bottom, left=wall, right=b.right)

    for r in range(r_min, r_max + 1):
        cell = ws.cell(r, c_max)
        b = copy(cell.border)
        cell.border = Border(top=b.top, bottom=b.bottom, left=b.left, right=wall)


# ── Validação e correção do layout ───────────────────────────────────────

def validate_layout(layout: BlocoLayout, n_pas: int, n_sala: int) -> BlocoLayout:
    """
    Verifica e corrige o layout retornado pela LLM:
      1. Remove ilhas PA que sobrepõem a SALA ou CATRACA
      2. Ajusta a contagem de PAs se necessário (remove células em excesso da última ilha PA)
      3. Loga warnings para diagnóstico
    """
    from copy import deepcopy

    # Constrói mapa de células ocupadas por tipo
    occupied: dict[tuple, str] = {}
    ilhas_ok = []

    # Processa SALA e CATRACA primeiro (têm prioridade)
    for ilha in layout.ilhas:
        if ilha.tipo in ('SALA', 'CATRACA'):
            for dr in range(ilha.altura):
                for dc in range(ilha.largura):
                    key = (ilha.row_offset + dr, ilha.col_offset + dc)
                    occupied[key] = ilha.tipo
            ilhas_ok.append(ilha)

    # Processa PAs verificando sobreposição
    pa_total = 0
    for ilha in layout.ilhas:
        if ilha.tipo != 'PA':
            continue
        conflito = False
        for dr in range(ilha.altura):
            for dc in range(ilha.largura):
                key = (ilha.row_offset + dr, ilha.col_offset + dc)
                if key in occupied:
                    print(f"  [Layout] AVISO: ilha PA ({ilha.row_offset},{ilha.col_offset} "
                          f"{ilha.altura}×{ilha.largura}) sobrepõe {occupied[key]} em {key} — removida")
                    conflito = True
                    break
            if conflito:
                break
        if not conflito:
            for dr in range(ilha.altura):
                for dc in range(ilha.largura):
                    occupied[(ilha.row_offset + dr, ilha.col_offset + dc)] = 'PA'
            pa_total += ilha.altura * ilha.largura
            ilhas_ok.append(ilha)

    # Ajuste fino: se PAs em excesso, reduz a última ilha PA
    if pa_total > n_pas:
        excesso = pa_total - n_pas
        for i in range(len(ilhas_ok) - 1, -1, -1):
            if ilhas_ok[i].tipo == 'PA':
                ilha = ilhas_ok[i]
                capacidade = ilha.altura * ilha.largura
                if capacidade <= excesso:
                    excesso -= capacidade
                    ilhas_ok.pop(i)
                else:
                    # Reduz a altura da ilha
                    linhas_remover = excesso // ilha.largura
                    resto = excesso % ilha.largura
                    nova_altura = ilha.altura - linhas_remover - (1 if resto else 0)
                    if nova_altura <= 0:
                        ilhas_ok.pop(i)
                    else:
                        ilhas_ok[i] = ilha.model_copy(update={'altura': nova_altura})
                    excesso = 0
                if excesso == 0:
                    break
    elif pa_total < n_pas:
        print(f"  [Layout] AVISO: layout gerou {pa_total} PAs mas precisava de {n_pas} — "
              f"faltam {n_pas - pa_total} PAs (a LLM deve ter cometido erro de contagem)")

    return layout.model_copy(update={'ilhas': ilhas_ok})

def apply_layout(ws, nome: str, layout: BlocoLayout, canvas: dict,
                 fill_pa: PatternFill, fill_sala: PatternFill) -> tuple[list, list, list]:
    """
    Expande as ilhas retangulares do BlocoLayout para coordenadas absolutas
    e pinta as células na planilha.

    Retorna (pa_cells, sala_cells, catraca_cells).
    """
    row_origin = canvas['row_start']
    col_origin = canvas['col_start']

    pa_cells:      list = []
    sala_cells:    list = []
    catraca_cells: list = []

    catraca_border = Border(
        left=Side(border_style='medium', color='FFFFFF00'),
        right=Side(border_style='medium', color='FFFFFF00'),
        top=Side(border_style='medium', color='FFFFFF00'),
        bottom=Side(border_style='medium', color='FFFFFF00'),
    )

    for ilha in layout.ilhas:
        for dr in range(ilha.altura):
            for dc in range(ilha.largura):
                r = row_origin + ilha.row_offset + dr
                c = col_origin + ilha.col_offset + dc
                if r < 1 or c < 1:
                    continue
                cell = ws.cell(r, c)

                if ilha.tipo == 'PA':
                    cell.value = nome
                    cell.fill  = fill_pa
                    cell.font  = FONT_WHITE
                    pa_cells.append((r, c))

                elif ilha.tipo == 'SALA':
                    cell.value = f'SALA-{nome}'
                    cell.fill  = fill_sala
                    cell.font  = FONT_WHITE
                    sala_cells.append((r, c))

                elif ilha.tipo == 'CATRACA':
                    cell.value  = f'CATRACA-{nome}'
                    cell.fill   = FILL_CATRACA
                    cell.font   = FONT_WHITE
                    cell.border = catraca_border
                    catraca_cells.append((r, c))

    return pa_cells, sala_cells, catraca_cells


# ── Execução da proposta ──────────────────────────────────────────────────

def execute(ws, proposta, plant_data: dict, canvas_map: dict) -> dict:
    """
    1. Libera células dos clientes indicados
    2. Para cada ação 'alocar': aplica o layout retornado pelo BlockDesignerAgent
    3. Desenha paredes ao redor de cada bloco (bounding box PA+SALA)
    """
    log = {'liberadas': {}, 'alocadas': {}}

    # ── LIBERAR ──────────────────────────────────────────────────────────
    for acao in proposta.acoes:
        if acao.tipo != 'liberar':
            continue
        alvo  = acao.cliente_a_liberar or acao.setor
        count = 0
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                if count >= acao.quantidade:
                    break
                cell = ws.cell(r, c)
                v    = cell.value
                if isinstance(v, float) and v == int(v):
                    v = str(int(v))
                elif v is not None:
                    v = str(v).strip()
                if v == str(alvo).strip():
                    log['liberadas'][(r, c)] = v
                    cell.value = '0'
                    cell.fill  = FILL_LIBERADO
                    cell.font  = FONT_SMALL
                    count += 1
            if count >= acao.quantidade:
                break
        print(f"  Liberadas: {count}/{acao.quantidade} células ('{alvo}')")

    # ── ALOCAR ────────────────────────────────────────────────────────────
    # Os layouts foram salvos na tool criar_espaco dentro do canvas_map.
    # Percorre os canvas que tiveram layout gerado, na ordem de criação.
    color_idx = 0
    for canvas in canvas_map.values():
        layout = canvas.get('_layout')
        if layout is None:
            continue

        # Nome que a LLM usou ao chamar a tool
        nome = canvas.get('_nome', layout.nome)

        pa_hex, sala_hex = _PALETTE[color_idx % len(_PALETTE)]
        color_idx += 1

        # Valida e corrige sobreposições antes de aplicar
        acao_pa  = sum(i.altura * i.largura for i in layout.ilhas if i.tipo == 'PA')
        acao_sala = sum(i.altura * i.largura for i in layout.ilhas if i.tipo == 'SALA')
        layout = validate_layout(layout, acao_pa, acao_sala)

        pa_cells, sala_cells, catraca_cells = apply_layout(
            ws, nome, layout, canvas, _fill(pa_hex), _fill(sala_hex)
        )

        # Paredes: bounding box das PAs + SALA (a catraca fica fora)
        all_inner = pa_cells + sala_cells
        if all_inner:
            rs = [r for r, c in all_inner]
            cs = [c for r, c in all_inner]
            bbox = [(r, c)
                    for r in range(min(rs), max(rs) + 1)
                    for c in range(min(cs), max(cs) + 1)]
            draw_walls_around_block(ws, bbox)

        log['alocadas'][nome] = {
            'pa': pa_cells, 'sala': sala_cells, 'catraca': catraca_cells
        }
        print(f"  {nome}: {len(pa_cells)} PAs + {len(sala_cells)} sala"
              f" + {len(catraca_cells)} catraca(s)")

    return log


# ── Relatório TXT ──────────────────────────────────────────────────────────

def write_report(ws_orig, ws_new, proposta, log, path: str) -> int:
    changes = []
    for r in range(1, ws_orig.max_row + 1):
        for c in range(1, ws_orig.max_column + 1):
            old = str(ws_orig.cell(r, c).value or '').strip()
            new = str(ws_new.cell(r, c).value or '').strip()
            if old != new:
                changes.append((r, c, get_column_letter(c), old, new))

    def section(title, items):
        out.append(f"\n{'─'*60}\n{title} ({len(items)})\n{'─'*60}")
        if not items:
            out.append("  (nenhuma)")
            return
        for r, c, col_l, old, new in sorted(items, key=lambda x: (x[0], x[1])):
            out.append(f"  {col_l}{r:<4}  {old:>8} → {new}")

    liberadas = [(r, c, cl, o, n) for r, c, cl, o, n in changes if n == '0']
    catracas  = [(r, c, cl, o, n) for r, c, cl, o, n in changes if 'CATRACA' in n]

    out = [
        '='*60, 'RELATÓRIO DE MUDANÇAS — SPACE PLAN', '='*60,
        f"Proposta : {proposta.proposta} — {proposta.nome}",
        f"Descrição: {proposta.descricao}",
        f"Custo    : {proposta.custo_obras.upper()}",
        f"Catracas : {proposta.catracas_novas} nova(s)",
        "",
        f"Total de células alteradas: {len(changes)}",
        f"  Liberadas : {len(liberadas)}",
    ]
    for nome, info in log['alocadas'].items():
        out.append(f"  {nome:<20}: {len(info['pa'])} PAs"
                   f" + {len(info['sala'])} sala"
                   f" + {len(info.get('catraca', []))} catraca(s)")
    out.append(f"  Catracas  : {len(catracas)}")

    section("POSIÇÕES LIBERADAS", liberadas)
    for nome, info in log['alocadas'].items():
        section(f"PAs {nome}",
                [(r, c, get_column_letter(c), '', nome) for r, c in info['pa']])
        if info['sala']:
            section(f"SALA {nome}",
                    [(r, c, get_column_letter(c), '', f'SALA-{nome}') for r, c in info['sala']])
        if info.get('catraca'):
            section(f"CATRACAS {nome}",
                    [(r, c, get_column_letter(c), '', f'CATRACA-{nome}') for r, c in info['catraca']])
    section("TODAS AS CATRACAS", catracas)
    out += ['', '='*60, 'FIM', '='*60]

    with open(path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(out))
    return len(changes)


# ── Main ───────────────────────────────────────────────────────────────────

async def main():
    with open('premissas.txt', encoding='utf-8') as f:
        premissas = f.read().strip()
    print("\n== PREMISSAS ==")
    print(premissas)

    print("\nCarregando planta...")
    wb, ws = load_plant()

    print("Analisando zonas...")
    plant_data = scan_plant(ws, FORBIDDEN_PATTERNS)
    plant_info = build_plant_info(plant_data)

    print("Gerando mapa 2D...")
    mapa_2d = build_mapa_2d(ws)
    print(f"  Mapa gerado ({len(mapa_2d.splitlines())} linhas)")

    # Canvas pré-calculado a partir das premissas
    # O PlannerAgent vai chamar criar_espaco(nome, n_pas, n_sala)
    # e a tool usa canvas_map para saber onde posicionar cada bloco
    canvas_map = build_canvas_map(premissas)
    print(f"  Canvas pré-calculados: {list(canvas_map.keys())}")

    print("\nConsultando PlannerAgent...")
    deps = PlannerDeps(
        plant_info=plant_info,
        mapa_2d=mapa_2d,
        premissas=premissas,
        canvas_map=canvas_map,
    )
    result = await planner_agent.run("Gere a proposta de space planning.", deps=deps)
    proposta = result.output
    print(f"  → Proposta: {proposta.nome}")

    os.makedirs('propostas', exist_ok=True)
    new_wb, new_ws = clone_ws(ws)
    new_ws.title = re.sub(r'[:\\/?*\[\]]', '', proposta.nome)[:31]
    log = execute(new_ws, proposta, plant_data, canvas_map)

    safe      = proposta.nome.lower().replace(' ', '_')[:40]
    xlsx_path = f"propostas/proposta_{safe}.xlsx"
    txt_path  = f"propostas/proposta_{safe}_mudancas.txt"
    json_path = "propostas/proposta_llm.json"

    new_wb.save(xlsx_path)
    n = write_report(ws, new_ws, proposta, log, txt_path)
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(proposta.model_dump(), f, ensure_ascii=False, indent=2)

    print(f"\n✓ {xlsx_path}")
    print(f"✓ {txt_path} ({n} células alteradas)")
    print(f"✓ {json_path}")


if __name__ == '__main__':
    asyncio.run(main())

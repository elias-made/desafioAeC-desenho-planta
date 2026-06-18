import asyncio
import json
import os
import re
from copy import copy
from typing import List, Set, Tuple
import shutil

import openpyxl
from dotenv import load_dotenv
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from ScannerPremissas import scan_orange_context, build_context_string_for_llm

from Agents import PlannerDeps, orquestrador, posicionador
from BlockMapper import scan_plant, describe_for_llm

load_dotenv()

SHEET_NAME, CORRIDOR_GAP = 'JPIII', 3
FORBIDDEN_PATTERNS = {'SALA 1', 'SALA 2', 'SALA 3', 'SALA 4', 'COWORKING', 'SALA CLIENTE', 'SALA1', 'SALA2', 'SALA3', 'SALA4'}

FILL_LIBERADO = PatternFill(start_color='BDC3C7', end_color='BDC3C7', fill_type='solid')
FILL_NEW_CLIENT = PatternFill(start_color='34495E', end_color='34495E', fill_type='solid')
FILL_SALA = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')

FONT_WHITE = Font(color='FFFFFF', bold=True, size=8)
FONT_SMALL = Font(size=8)
FONT_SALA = Font(color='0000FF', bold=True, size=8)

THIN_BLUE = Side(border_style="thin", color="0000FF")
BORDER_SALA = Border(top=THIN_BLUE, left=THIN_BLUE, right=THIN_BLUE, bottom=THIN_BLUE)

_bench_partners_cache = {}

def get_bench_partner(c, ws):
    global _bench_partners_cache
    if not _bench_partners_cache:
        pa_cols = set()
        for r in range(1, min(200, ws.max_row + 1)):
            for col in range(1, ws.max_column + 1):
                val = ws.cell(r, col).value
                if val is not None:
                    v_str = str(val).strip().upper()
                    if v_str not in ('0', 'VAZIO', 'CT', 'SA', 'SALA', 'CW', '##', ''):
                        pa_cols.add(col)
        
        sorted_cols = sorted(list(pa_cols))
        i = 0
        while i < len(sorted_cols) - 1:
            c1, c2 = sorted_cols[i], sorted_cols[i+1]
            if c2 == c1 + 1:
                _bench_partners_cache[c1] = c2
                _bench_partners_cache[c2] = c1
                i += 2
            else:
                i += 1
    return _bench_partners_cache.get(c)

def load_plant(path='planta.xlsx', sheet=SHEET_NAME):
    wb = openpyxl.load_workbook(path, data_only=True)
    return wb, wb[sheet]

def clone_ws(ws):
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            src, tgt = ws.cell(r, c), new_ws.cell(r, c)
            tgt.value = src.value
            if src.has_style:
                tgt.font, tgt.fill, tgt.alignment, tgt.border = copy(src.font), copy(src.fill), copy(src.alignment), copy(src.border)
    for col_l, dim in ws.column_dimensions.items(): new_ws.column_dimensions[col_l].width = dim.width
    for row_i, dim in ws.row_dimensions.items(): new_ws.row_dimensions[row_i].height = dim.height
    return new_wb, new_ws

def build_plant_info(plant_data):
    client_cells = plant_data['client_cells']
    lines = ["CÉLULAS POR VALOR NA PLANTA (EXATAMENTE como aparecem nas células):"]
    for v in sorted(client_cells, key=lambda v: len(client_cells[v]), reverse=True)[:20]:
        lines.append(f"  '{v}': {len(client_cells[v])} PAs")
    return "\n".join(lines)

def build_blocos_info(plant_data, ws_max_row, ws_max_col, ws=None):
    if ws is None:
        return "Nenhum bloco laranja mapeado."
        
    IGNORED_TAGS = {
        'CATRACA', 'ESCANINHOS', 'SALA', 'PUXADINHO', 'COWORKING', 
        'SALA1', 'SALA2', 'SALA3', 'SALA4', 'SALA CLIENTE'
    }
        
    macro_blocks = scan_orange_context('planta.xlsx', SHEET_NAME)
    client_cells = plant_data['client_cells']
    
    lines = []
    for idx, block in enumerate(macro_blocks, start=1):
        b_id = f"vazio-{idx}"
        r_min, r_max, c_min, c_max = block['bounding_box']
        
        lines.append(f"--------------------Bloco {idx} ({b_id})--------------------")
        lines.append(f"Localização: colunas {get_column_letter(c_min)}-{get_column_letter(c_max)}, linhas {r_min}-{r_max}")
        
        if block['texts']:
            lines.append("📌 Anotações na borda laranjas:")
            for txt in block['texts']:
                lines.append(f"  - {txt}")
        lines.append("")
        
        lines.append("Células disponíveis:")
        
        clients_in_block = {}
        empty_count = 0
        total_cells_in_block = 0
        
        for cliente, cells in client_cells.items():
            cliente_norm = str(cliente).strip()
            
            if cliente_norm.upper() in IGNORED_TAGS:
                continue
                
            count = sum(1 for (r, c) in cells if r_min <= r <= r_max and c_min <= c <= c_max)
            if count > 0:
                total_cells_in_block += count
                if cliente_norm.upper() in ('0', 'VAZIO'):
                    empty_count += count
                    
                clients_in_block[cliente_norm] = count
                
        for cl, qtd in sorted(clients_in_block.items(), key=lambda x: x[0]):
            if cl.upper() in ('0', 'VAZIO'):
                lines.append(f"Célula '{cl}': quantidade: {qtd};")
            else:
                lines.append(f"Cliente '{cl}': quantidade: {qtd};")
                
        lines.append("")
        lines.append(f"Células totais ocupadas e desocupadas: {total_cells_in_block}")
        lines.append(f"Células sem clientes: {empty_count};")
        lines.append("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n")
        
    return "\n".join(lines)

def execute_alocacao(ws, proposta, plant_data, allowed_cells: Set[Tuple[int, int]]) -> tuple:
    from BlockMapper import flood_fill, group_zones
    log = {'realocadas': {}, 'liberadas': {}, 'avisos': []}
    
    cell_values = {}
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is not None:
                cell_values[(r, c)] = str(int(v) if isinstance(v, float) and v == int(v) else v).strip()
            else:
                cell_values[(r, c)] = ""

    # 1. Mapeia fontes e preenchimentos originais de todos os clientes ativos da planilha
    client_fills = {}
    client_fonts = {}
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            v = cell.value
            if v is not None:
                v_str = str(int(v) if isinstance(v, float) and v == int(v) else v).strip().upper()
                if v_str not in ('0', 'VAZIO', 'CT', 'SA', 'SALA', 'CW', '##', ''):
                    if cell.fill and cell.fill.patternType == 'solid':
                        client_fills[v_str] = copy(cell.fill)
                    if cell.font:
                        client_fonts[v_str] = copy(cell.font)

    # Cores exclusivas para novos clientes que não existem originalmente na planilha
    default_new_fills = {
        'NOVO A': PatternFill(start_color='34495E', end_color='34495E', fill_type='solid'), # Azul escuro
        'NOVO B': PatternFill(start_color='9B59B6', end_color='9B59B6', fill_type='solid'), # Roxo Amethyst
    }
    default_new_fonts = {
        'NOVO A': Font(color='FFFFFF', bold=True, size=8),
        'NOVO B': Font(color='FFFFFF', bold=True, size=8),
    }

    # Carrega as coordenadas das caixas laranjas dinamicamente
    macro_blocks = scan_orange_context('planta.xlsx', SHEET_NAME)

    freed_by_client = {}
    non_client_values = {'0', 'VAZIO', 'CT', 'SA', 'SALA', 'CW', '##', ''}
    active_clients_cache = {v.upper() for v in cell_values.values() if v.upper() not in non_client_values}
    
    # Separa as ações em duas fases: Primeiro libera tudo, depois posiciona sobre os espaços livres
    acoes_liberar = [a for a in proposta.acoes if a.tipo.lower().strip() == 'liberar']
    acoes_realocar = [a for a in proposta.acoes if a.tipo.lower().strip() in ('realocar', 'alocar')]
    
    reduced_clients = {str(a.cliente_a_liberar or a.cliente).strip().upper() for a in acoes_liberar}

    def is_safe_cell(r, c, target):
        if (r, c) not in allowed_cells: return False
        t_up = target.upper()
        current_val = cell_values.get((r, c), "").strip().upper()
        
        partner_col = get_bench_partner(c, ws)
        if partner_col:
            partner_val = cell_values.get((r, partner_col), "").strip().upper()
            if partner_val and partner_val not in ('0', 'VAZIO', '', t_up) and partner_val not in reduced_clients:
                if partner_val in active_clients_cache: return False
        
        return current_val in non_client_values or current_val == t_up or current_val in reduced_clients

    def fill_bfs(disp, qtd):
        if not disp or qtd <= 0: return []
        alocadas = []
        for bloco in sorted(flood_fill(set(disp)), key=len, reverse=True):
            for atual in sorted(bloco, key=lambda coord: (coord[1], coord[0])):
                if len(alocadas) < qtd: alocadas.append(atual)
        if len(alocadas) < qtd: alocadas.extend([c for c in sorted(disp, key=lambda x: (x[1], x[0])) if c not in alocadas][:qtd - len(alocadas)])
        return alocadas

    # === FASE 1: LIBERAR (Sempre executada primeiro para criar as mesas '0') ===
    for acao in acoes_liberar:
        alvo = str(acao.cliente_a_liberar or acao.cliente).strip()
        alvo_norm = alvo.upper()
        
        target_block = acao.novo_cliente.strip()
        matching_block = next((b for b in macro_blocks if f"vazio-{macro_blocks.index(b)+1}" == target_block), None)
        
        if matching_block:
            r_min, r_max, c_min, c_max = matching_block['bounding_box']
            cells_to_free = []
            for r in range(r_min, r_max + 1):
                for c in range(c_min, c_max + 1):
                    if cell_values.get((r, c), "").upper() == alvo_norm:
                        cells_to_free.append((r, c))
            target_zone = cells_to_free
        else:
            cells_to_free = [k for k, v in cell_values.items() if v.upper() == alvo_norm]
            zones = group_zones(flood_fill(set(cells_to_free)), gap=CORRIDOR_GAP)
            target_zone = zones[0] if zones else cells_to_free
            
        # Alinha a variável de ordenação fora dos blocos condicionais
        cells_to_free_sorted = sorted(target_zone, key=lambda coord: (coord[1], coord[0]))
            
        for r, c in cells_to_free_sorted[:acao.quantidade]:
            ws.cell(r, c).value, cell_values[(r, c)] = '0', '0'
            ws.cell(r, c).fill, ws.cell(r, c).font = FILL_LIBERADO, FONT_SMALL
            freed_by_client.setdefault(alvo, []).append((r, c))
            log['liberadas'][(r, c)] = alvo

    # === FASE 2: REALOCAR (Apenas posiciona onde já tem espaço livre '0', 'vazio' ou 'VAZIO') ===
    for acao in acoes_realocar:
        pool = []
        target_block = acao.novo_cliente.strip()
        
        # DEFINE CLI_CLEAN NO INÍCIO DO LOOP (Corrige o NameError)
        cli_clean = re.sub(r'-(complemento|parte|excedente|residuo)', '', str(acao.cliente), flags=re.IGNORECASE).strip()
        m_novo = re.search(r'novo[-\s]*(?:cliente[-\s]*)?([a-zA-Z0-9]+)', cli_clean, flags=re.IGNORECASE)
        cli_clean = f"Novo {m_novo.group(1).upper()}" if m_novo else cli_clean[:15]
        
        matching_block = next((b for b in macro_blocks if f"vazio-{macro_blocks.index(b)+1}" == target_block), None)
        
        if matching_block:
            r_min, r_max, c_min, c_max = matching_block['bounding_box']
            
            # Coleta estritamente células que estão desocupadas ('0', 'VAZIO', 'vazio')
            for r in range(r_min, r_max + 1):
                for c in range(c_min, c_max + 1):
                    val = cell_values.get((r, c), "").upper()
                    if val in ('0', 'VAZIO', '') and is_safe_cell(r, c, cli_clean):
                        pool.append((r, c))
                        
        elif "liberad" in target_block.lower():
            client_origem = target_block.replace("-liberado", "").replace("-liberado".upper(), "").strip().upper()
            matching_key = next((key for key in freed_by_client.keys() if key.upper() == client_origem), None)
            if matching_key:
                pool = [(r, c) for r, c in freed_by_client[matching_key] if cell_values.get((r, c), "").upper() in ('0', 'VAZIO') and is_safe_cell(r, c, cli_clean)]
        
        dests = fill_bfs(pool, acao.quantidade)
        
        if len(dests) < acao.quantidade:
            geral = [k for k, v in cell_values.items() if v.upper() in ('0', 'VAZIO') and is_safe_cell(k[0], k[1], cli_clean) and k not in dests]
            dests.extend(fill_bfs(geral, acao.quantidade - len(dests)))

        for dr, dc in dests:
            ws.cell(dr, dc).value, cell_values[(dr, dc)] = cli_clean, cli_clean
            
            # Recupera de forma 100% dinâmica a cor de fundo e a fonte originais mapeadas do cliente
            fill_to_apply = client_fills.get(cli_clean.upper()) or default_new_fills.get(cli_clean.upper(), FILL_NEW_CLIENT)
            font_to_apply = FONT_WHITE if cli_clean.upper() in ['NOVO A', 'NOVO B'] else copy(ws.cell(dr, dc).font)
            
            ws.cell(dr, dc).fill = fill_to_apply
            if font_to_apply:
                ws.cell(dr, dc).font = font_to_apply
                
            active_clients_cache.add(cli_clean.upper())
            log['realocadas'].setdefault(f"{cli_clean} → {target_block}", []).append((dr, dc))

    return log, cell_values

def write_report(ws_orig, ws_new, proposta, log, path: str):
    changes = [(r, c, get_column_letter(c), str(ws_orig.cell(r, c).value or ''), str(ws_new.cell(r, c).value or ''))
               for r in range(1, ws_orig.max_row + 1) for c in range(1, ws_orig.max_column + 1) if str(ws_orig.cell(r, c).value or '') != str(ws_new.cell(r, c).value or '')]
    with open(path, 'w', encoding='utf-8') as f:
        f.write(f"PROPOSTA: {proposta.nome}\nCUSTO: {proposta.custo_obras}\n\nALTERACOES: {len(changes)}")
        if changes:
            f.write("\n\nDETALHES DAS MUDANCAS:\n")
            for r, c, col, old, new in sorted(changes, key=lambda x: (x[0], x[1])):
                f.write(f"  {col}{r}: {old} -> {new}\n")
        if log.get('avisos'):
            f.write("\n\nAVISOS:\n")
            for aviso in log['avisos']:
                f.write(f"  - {aviso}\n")

def formatar_input_deps(deps: PlannerDeps, prompt_adicional: str = "") -> str:
    return f"""================================================================================
INPUT DE AUDITORIA - PARÂMETROS RECEBIDOS PELO AGENTE
================================================================================

[PROMPT ADICIONAL ENVIADO]
{prompt_adicional}

[PREMISSAS DO TXT]
{deps.premissas}

[ESTRATÉGIA MACRO DEFINIDA]
{deps.estrategia_macro or '(Nenhuma)'}

[BLOCOS INFO]
{deps.blocos_info}

[PLANT INFO]
{deps.plant_info}
"""

def salvar_auditoria(nome_agente: str, input_data: str, output_data: str, safe_name: str):
    dir_auditoria = f"propostas/auditoria_{safe_name}"
    os.makedirs(dir_auditoria, exist_ok=True)
    with open(f"{dir_auditoria}/{nome_agente}_input.txt", "w", encoding="utf-8") as f:
        f.write(input_data)
    with open(f"{dir_auditoria}/{nome_agente}_output.txt", "w", encoding="utf-8") as f:
        f.write(output_data)

async def main():
    with open('premissas.txt', encoding='utf-8') as f: 
        premissas_txt = f.read().strip()

    dados_laranjas = scan_orange_context('planta.xlsx', SHEET_NAME)
    premissas_visuais = build_context_string_for_llm(dados_laranjas)
    premissas_completas = f"{premissas_txt}\n\n{premissas_visuais}"

    wb, ws = load_plant()
    plant_data = scan_plant(ws, FORBIDDEN_PATTERNS)
    
    deps = PlannerDeps(
        plant_info=build_plant_info(plant_data), 
        mapa_2d="[Mapa Omitido]", 
        premissas=premissas_completas,
        blocos_info=build_blocos_info(plant_data, ws.max_row, ws.max_column, ws)
    )
    
    print("1. Orquestrador definindo Estratégia...")
    estrategia = await orquestrador.run("Defina a estratégia macro.", deps=deps)
    deps.estrategia_macro = estrategia.output.model_dump_json()
    
    cliente_a_reduzir = estrategia.output.nome_exato_cliente_principal
    print(f"\n🔍 ALVO DE REDUÇÃO CORRETO IDENTIFICADO PELA IA: '{cliente_a_reduzir}'")
    
    allowed_cells = set(plant_data['client_cells'].get(cliente_a_reduzir, []))
    for k in ('VAZIO', 'vazio', '0'):
        allowed_cells.update(plant_data['client_cells'].get(k, []))
        
    print("2. Posicionador gerando Ações...")
    prompt_pos = "Gere a proposta detalhada."
    proposta = await posicionador.run(prompt_pos, deps=deps)
    
    print("3. Executando Alocação Bruta...")
    os.makedirs('propostas', exist_ok=True)
    
    safe = re.sub(r'[^a-zA-Z0-9_]', '', proposta.output.nome.replace(' ', '_'))[:40].lower()
    dest_file = f"propostas/proposta_{safe}.xlsx"
    
    shutil.copy("planta.xlsx", dest_file)
    
    new_wb = openpyxl.load_workbook(dest_file)
    new_ws = new_wb[SHEET_NAME]
    
    log, _ = execute_alocacao(new_ws, proposta.output, plant_data, allowed_cells)
    
    print("\n4. Salvando resultados...")
    salvar_auditoria("1_orquestrador", formatar_input_deps(deps, "Defina a estratégia macro."), estrategia.output.model_dump_json(indent=2), safe)
    salvar_auditoria("2_posicionador", formatar_input_deps(deps, prompt_pos), proposta.output.model_dump_json(indent=2), safe)
    
    new_wb.save(dest_file)
    
    # --- NOVO: RECALCULA O ESTADO FINAL DOS BLOCOS ---
    final_plant_data = scan_plant(new_ws, FORBIDDEN_PATTERNS)
    final_blocos_info = build_blocos_info(final_plant_data, new_ws.max_row, new_ws.max_column, new_ws)
    
    # Print no terminal
    print("\n================================================================================")
    print("ESTADO FINAL DOS BLOCOS APÓS AS MOVIMENTAÇÕES:")
    print("================================================================================")
    print(final_blocos_info)
    
    # Salva as auditorias originais dos agentes
    salvar_auditoria("1_orquestrador", formatar_input_deps(deps, "Defina a estratégia macro."), estrategia.output.model_dump_json(indent=2), safe)
    salvar_auditoria("2_posicionador", formatar_input_deps(deps, prompt_pos), proposta.output.model_dump_json(indent=2), safe)
    
    # Grava o relatório de mudanças principal
    write_report(ws, new_ws, proposta.output, log, f"propostas/proposta_{safe}_mudancas.txt")
    
    # --- NOVO: GRAVA O MAPA FINAL DE BLOCOS EM UM ARQUIVO .TXT INDEPENDENTE ---
    caminho_relatorio_final = f"propostas/proposta_{safe}_blocos_finais.txt"
    with open(caminho_relatorio_final, "w", encoding="utf-8") as f:
        f.write("================================================================================\n")
        f.write("RELATÓRIO DO ESTADO FINAL DOS BLOCOS APÓS AS MOVIMENTAÇÕES\n")
        f.write("================================================================================\n\n")
        f.write(final_blocos_info)
        
    print(f"✓ Processo concluído. Arquivos salvos em 'propostas/':")
    print(f"  - Excel modificado: proposta_{safe}.xlsx")
    print(f"  - Lista de mudanças: proposta_{safe}_mudancas.txt")
    print(f"  - Mapa final de blocos (TXT solicitado): proposta_{safe}_blocos_finais.txt")

if __name__ == '__main__':
    asyncio.run(main())
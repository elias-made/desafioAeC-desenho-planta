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
from ScannerPremissas import scan_orange_context, build_context_string_for_llm, normalize_val

# Importando dependências e agentes do Agents.py (Fluxo de 2 Agentes)
from Agents import PosicionadorDeps, OrganizadorDeps, posicionador, organizador
from BlockMapper import scan_plant, describe_for_llm
from pydantic import BaseModel, Field

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

# ══════════════════════════════════════════════════════════════════════════
# Estruturas de Validação de Dados Internas (Pydantic com coerção automática)
# ══════════════════════════════════════════════════════════════════════════

class Acao(BaseModel):
    """Representa a estrutura de uma ação de alteração física da planta"""
    tipo: str  # "liberar", "realocar" (posicionar) ou "transferir" (swap)
    cliente: str = ""
    quantidade: int = 0
    bloco: str = ""
    ambiente: str = ""
    novo_cliente: str = ""  # Mantido para compatibilidade
    cliente_a_liberar: str = ""  # Mantido para compatibilidade
    sala_lugares: int = 0
    
    # Atributos específicos para a função única de transferência (Organizador)
    cliente_a: str = ""
    bloco_a: str = ""
    ambiente_a: str = ""
    quantidade_a: int = 0
    cliente_b: str = ""
    bloco_b: str = ""
    ambiente_b: str = ""
    quantidade_b: int = 0

class PropostaMock:
    """Helper class para envelopar as ações em formato compatível com execute_alocacao"""
    def __init__(self, nome, custo_obras, acoes):
        self.nome = nome
        self.custo_obras = custo_obras
        self.acoes = acoes

# ══════════════════════════════════════════════════════════════════════════
# Funções Auxiliares de Limpeza e Geometria
# ══════════════════════════════════════════════════════════════════════════

def clean_json_string(s: str) -> str:
    """
    Remove marcações de bloco de código Markdown do texto retornado pela LLM
    e extrai estritamente a string contendo o objeto JSON {}
    """
    s = s.strip()
    match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', s, re.DOTALL | re.IGNORECASE)
    if match:
        return match.group(1).strip()
    s = re.sub(r'^```(?:json)?\s*', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\s*```$', '', s)
    return s.strip()

def get_bench_partner(c, ws):
    global _bench_partners_cache
    if not _bench_partners_cache:
        pa_cols = set()
        for r in range(1, min(200, ws.max_row + 1)):
            for col in range(1, ws.max_column + 1):
                val = ws.cell(r, col).value
                if val is not None:
                    v_str = str(val).strip().upper()
                    if v_str not in ('VAZIO', 'CT', 'SA', 'SALA', 'CW', '##', ''):
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
        
    macro_blocks = scan_orange_context('planta.xlsx', SHEET_NAME)
    
    # Identifica dinamicamente quais células são mesas válidas (evita corredores)
    allowed_cells = set()
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            v = cell.value
            if v is not None:
                v_str = str(v).strip()
                if v_str != "" and v_str.upper() not in ('CT', 'CATRACA', 'SA', 'CW', 'COWORKING', '##'):
                    allowed_cells.add((r, c))
                    
    all_cells_cache = {}
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            all_cells_cache[(r, c)] = ws.cell(row=r, column=c)
            
    lines = []
    for idx, block in enumerate(macro_blocks, start=1):
        b_id = f"vazio-{idx}"
        r_min, r_max, c_min, c_max = block['bounding_box']
        col_min = get_column_letter(c_min)
        col_max = get_column_letter(c_max)
        
        lines.append(f"-------------------- Bloco {idx} ({b_id}) --------------------")
        lines.append(f"Localização Macro: colunas {col_min}-{col_max}, linhas {r_min}-{r_max}")
        if block['texts']:
            lines.append(f"📌 Anotações na borda: {', '.join(block['texts'])}")
        lines.append("")
        
        for env in block.get('ambientes', []):
            env_id = env['id']
            env_r_min, env_r_max, env_c_min, env_c_max = env['bounding_box']
            env_col_min = get_column_letter(env_c_min)
            env_col_max = get_column_letter(env_c_max)
            
            client_counts = {}
            empty_count = 0
            
            for (r, c) in env['cells']:
                if (r, c) not in allowed_cells:
                    continue  # Filtra corredores e estruturas
                cell = all_cells_cache.get((r, c))
                val = cell.value if cell else None
                val_str = str(val).strip() if val is not None else ""
                
                if val_str.upper() in ('VAZIO', ''):
                    empty_count += 1
                elif val_str.upper() not in ('CT', 'CATRACA', 'SA', 'CW', 'COWORKING', '##'):
                    norm_val = normalize_val(val)
                    client_counts[norm_val] = client_counts.get(norm_val, 0) + 1
                    
            lines.append(f"  Ambiente {env_id} ({b_id}-{env_id}):")
            lines.append(f"    Limites da Bounding Box: colunas {env_col_min}-{env_col_max}, linhas {env_r_min}-{env_r_max}")
            lines.append(f"    Células de mesas físicas reais neste ambiente: {len([coord for coord in env['cells'] if coord in allowed_cells])}")
            lines.append(f"    Células sem clientes / em branco: {empty_count}")
            lines.append("    Clientes identificados:")
            if client_counts:
                for cli, qty in sorted(client_counts.items()):
                    lines.append(f"      - Cliente '{cli}': quantidade: {qty}")
            else:
                lines.append("      - Nenhum operador ativo")
            lines.append("")
            
        lines.append("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n")
        
    summary_lines = [
        "============================================================",
        "RESUMO TOTAL DE CLIENTES EM MESAS (PLANTA INTEIRA):",
        "============================================================"
    ]
    total_counts = {}
    ignored_on_summary = {'VAZIO', 'CT', 'SA', 'SALA', 'CW', '##', '', 'CATRACA', 'ESCANINHOS', 'SALA CLIENTE', 'COWORKING'}
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            if (r, c) not in allowed_cells:
                continue
            val = ws.cell(row=r, column=c).value
            if val is not None:
                val_str = str(int(val) if isinstance(val, float) and val == int(val) else val).strip()
                val_up = val_str.upper()
                if val_up not in ignored_on_summary:
                    total_counts[val_str] = total_counts.get(val_str, 0) + 1
    for client, qty in sorted(total_counts.items(), key=lambda x: x[0]):
        summary_lines.append(f"  Cliente '{client}': quantidade: {qty};")
    summary_lines.append("============================================================\n")
        
    return "\n".join(lines) + "\n" + "\n".join(summary_lines)

def parse_env_id(env_id_str: str) -> Tuple[str, str]:
    """
    Decodifica a nomenclatura de ambiente unificada de forma tolerante a sufixos explicativos.
    Ex: 'vazio-3-A (Zona Norte)' -> ('vazio-3', 'A')
    """
    match = re.search(r'(vazio-\d+)-([A-K])', env_id_str.strip(), re.IGNORECASE)
    if match:
        return match.group(1).lower(), match.group(2).upper()
    return "", ""

# ══════════════════════════════════════════════════════════════════════════
# Funções Lógicas de Execução Estrita com Transferência Assequível
# ══════════════════════════════════════════════════════════════════════════

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

    client_fills = {}
    client_fonts = {}
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            v = cell.value
            if v is not None:
                v_str = str(int(v) if isinstance(v, float) and v == int(v) else v).strip().upper()
                if v_str not in ('VAZIO', 'CT', 'SA', 'SALA', 'CW', '##', ''):
                    if cell.fill and cell.fill.patternType == 'solid':
                        client_fills[v_str] = copy(cell.fill)
                    if cell.font:
                        client_fonts[v_str] = copy(cell.font)

    default_new_fills = {
        'NOVO A': PatternFill(start_color='34495E', end_color='34495E', fill_type='solid'),
        'NOVO B': PatternFill(start_color='9B59B6', end_color='9B59B6', fill_type='solid'),
    }
    default_new_fonts = {
        'NOVO A': Font(color='FFFFFF', bold=True, size=8),
        'NOVO B': Font(color='FFFFFF', bold=True, size=8),
    }

    macro_blocks = scan_orange_context('planta.xlsx', SHEET_NAME)
    non_client_values = {'VAZIO', 'CT', 'SA', 'SALA', 'CW', '##', ''}
    active_clients_cache = {v.upper() for v in cell_values.values() if v.upper() not in non_client_values}
    
    acoes_liberar = [a for a in proposta.acoes if a.tipo.lower().strip() == 'liberar']
    acoes_realocar = [a for a in proposta.acoes if a.tipo.lower().strip() in ('realocar', 'posicionar')]
    acoes_transferir = [a for a in proposta.acoes if a.tipo.lower().strip() in ('transferir', 'swap', 'permutar')]
    
    reduced_clients = {str(a.cliente).strip().upper() for a in acoes_liberar}

    def is_safe_cell(r, c, target):
        if (r, c) not in allowed_cells: return False
        t_up = target.upper()
        current_val = cell_values.get((r, c), "").strip().upper()
        
        partner_col = get_bench_partner(c, ws)
        if partner_col:
            partner_val = cell_values.get((r, partner_col), "").strip().upper()
            if partner_val and partner_val not in ('VAZIO', '', t_up) and partner_val not in reduced_clients:
                if partner_val in active_clients_cache: return False
        
        return current_val in non_client_values or current_val == t_up or current_val in reduced_clients

    def get_env_cells(block_id_str: str, env_letter: str) -> List[Tuple[int, int]]:
        if not block_id_str or not env_letter:
            return []
        block_match = re.search(r'\d+', str(block_id_str))
        if not block_match:
            return []
        block_idx = int(block_match.group())
        if block_idx <= len(macro_blocks):
            block = macro_blocks[block_idx - 1]
            env = next((e for e in block.get('ambientes', []) if e['id'].upper() == env_letter.strip().upper()), None)
            if env:
                return list(env['cells'])
        return []

    def fill_bfs(disp, qtd):
        if not disp or qtd <= 0: return []
        alocadas = []
        for bloco in sorted(flood_fill(set(disp)), key=len, reverse=True):
            for atual in sorted(bloco, key=lambda coord: (coord[1], coord[0])):
                if len(alocadas) < qtd: alocadas.append(atual)
        if len(alocadas) < qtd: alocadas.extend([c for c in sorted(disp, key=lambda x: (x[1], x[0])) if c not in alocadas][:qtd - len(alocadas)])
        return alocadas

    # === FASE 1: FUNÇÃO DE REMOÇÃO ESTREITA (Posicionador) ===
    for acao in acoes_liberar:
        alvo = acao.cliente.strip()
        alvo_norm = alvo.upper()
        
        block_id = acao.bloco or (re.search(r'vazio-\d+', acao.novo_cliente).group() if re.search(r'vazio-\d+', acao.novo_cliente) else "")
        env_letter = acao.ambiente or (re.search(r'-([A-K])', acao.novo_cliente).group(1) if re.search(r'-([A-K])', acao.novo_cliente) else "")
        
        env_cells = get_env_cells(block_id, env_letter)
        if env_cells:
            cells_to_free = [coord for coord in env_cells if cell_values.get(coord, "").upper() == alvo_norm]
        else:
            cells_to_free = [k for k, v in cell_values.items() if v.upper() == alvo_norm]
            
        cells_to_free_sorted = sorted(cells_to_free, key=lambda coord: (coord[1], coord[0]))
        
        liberadas_count = 0
        for r, c in cells_to_free_sorted[:acao.quantidade]:
            ws.cell(r, c).value, cell_values[(r, c)] = 'vazio', 'vazio'
            ws.cell(r, c).fill, ws.cell(r, c).font = FILL_LIBERADO, FONT_SMALL
            log['liberadas'][(r, c)] = alvo
            liberadas_count += 1
            
        if liberadas_count < acao.quantidade:
            log['avisos'].append(f"Liberou apenas {liberadas_count} de {acao.quantidade} do cliente '{alvo}' no sub-ambiente {block_id}-{env_letter}.")

    # === FASE 2: FUNÇÃO DE POSICIONAMENTO ESTREITO E SEGURO (Posicionador) ===
    for acao in acoes_realocar:
        target = acao.cliente.strip()
        
        block_id = acao.bloco or (re.search(r'vazio-\d+', acao.novo_cliente).group() if re.search(r'vazio-\d+', acao.novo_cliente) else "")
        env_letter = acao.ambiente or (re.search(r'-([A-K])', acao.novo_cliente).group(1) if re.search(r'-([A-K])', acao.novo_cliente) else "")
        
        cli_clean = re.sub(r'-(complemento|parte|excedente|residuo)', '', target, flags=re.IGNORECASE).strip()
        m_novo = re.search(r'novo[-\s]*(?:cliente[-\s]*)?([a-zA-Z0-9]+)', cli_clean, flags=re.IGNORECASE)
        cli_clean = f"Novo {m_novo.group(1).upper()}" if m_novo else cli_clean[:15]
        
        env_cells = get_env_cells(block_id, env_letter)
        pool = []
        if env_cells:
            for (r, c) in env_cells:
                val = cell_values.get((r, c), "").upper()
                if val in ('VAZIO', '') and is_safe_cell(r, c, cli_clean):
                    pool.append((r, c))
                    
        # Restrição Estrita: Proíbe posicionar mais do que aquele ambiente físico permite
        dests = fill_bfs(pool, acao.quantidade)
        
        if len(dests) < acao.quantidade:
            log['avisos'].append(f"Capacidade física insuficiente no {block_id}-{env_letter}! Posicionou apenas {len(dests)} de {acao.quantidade} para '{cli_clean}'.")
            
        for dr, dc in dests:
            ws.cell(dr, dc).value, cell_values[(dr, dc)] = cli_clean, cli_clean
            fill_to_apply = client_fills.get(cli_clean.upper()) or default_new_fills.get(cli_clean.upper(), FILL_NEW_CLIENT)
            font_to_apply = FONT_WHITE if cli_clean.upper() in ['NOVO A', 'NOVO B'] else copy(ws.cell(dr, dc).font)
            
            ws.cell(dr, dc).fill = fill_to_apply
            if font_to_apply:
                ws.cell(dr, dc).font = font_to_apply
                
            active_clients_cache.add(cli_clean.upper())
            log['realocadas'].setdefault(f"{cli_clean} → {block_id}-{env_letter}", []).append((dr, dc))

    # === FASE 3: FUNÇÃO ÚNICA DE TRANSFERÊNCIA/PERMUTA ASSIMÉTRICA DIRETA (Organizador) ===
    for acao in acoes_transferir:
        cli_a = acao.cliente_a.strip()
        cli_b = acao.cliente_b.strip()
        
        cells_a = get_env_cells(acao.bloco_a, acao.ambiente_a)
        cells_b = get_env_cells(acao.bloco_b, acao.ambiente_b)
        
        # Localiza células ocupadas pelo Cliente A
        if cli_a.upper() in ('VAZIO', ''):
            occupied_a = [coord for coord in cells_a if cell_values.get(coord, "").upper() in ('VAZIO', '') and coord in allowed_cells]
        else:
            occupied_a = [coord for coord in cells_a if cell_values.get(coord, "").upper() == cli_a.upper()]
            
        # Localiza células ocupadas pelo Cliente B
        if cli_b.upper() in ('VAZIO', ''):
            occupied_b = [coord for coord in cells_b if cell_values.get(coord, "").upper() in ('VAZIO', '') and coord in allowed_cells]
        else:
            occupied_b = [coord for coord in cells_b if cell_values.get(coord, "").upper() == cli_b.upper()]
        
        occupied_a_sorted = sorted(occupied_a, key=lambda coord: (coord[1], coord[0]))
        occupied_b_sorted = sorted(occupied_b, key=lambda coord: (coord[1], coord[0]))
        
        targets_a = occupied_a_sorted[:acao.quantidade_a]
        targets_b = occupied_b_sorted[:acao.quantidade_b]
        
        # 1. Liberação mútua (esvazia as mesas de origem)
        for r, c in targets_a:
            ws.cell(r, c).value, cell_values[(r, c)] = 'vazio', 'vazio'
            ws.cell(r, c).fill, ws.cell(r, c).font = FILL_LIBERADO, FONT_SMALL
        for r, c in targets_b:
            ws.cell(r, c).value, cell_values[(r, c)] = 'vazio', 'vazio'
            ws.cell(r, c).fill, ws.cell(r, c).font = FILL_LIBERADO, FONT_SMALL
            
        # 2. Aloca Cliente A nas vagas disponíveis do Destino B (inclui as recém-liberadas por B)
        if cli_a.upper() not in ('VAZIO', ''):
            free_in_b = [coord for coord in cells_b if cell_values.get(coord, "").upper() in ('VAZIO', '') and coord in allowed_cells]
            free_in_b_sorted = sorted(free_in_b, key=lambda coord: (coord[1], coord[0]))
            
            dests_a = fill_bfs(free_in_b_sorted, len(targets_a))
            
            # TRAVA DE SEGURANÇA A: Alerta perda de posições do Cliente A
            if len(dests_a) < len(targets_a):
                log['avisos'].append(
                    f"ERRO DE CAPACIDADE: Tentou transferir {len(targets_a)} de '{cli_a}' para "
                    f"{acao.bloco_b}-{acao.ambiente_b}, mas só couberam {len(dests_a)}. "
                    f"Perda de {len(targets_a) - len(dests_a)} PAs!"
                )
                
            for dr, dc in dests_a:
                ws.cell(dr, dc).value, cell_values[(dr, dc)] = cli_a, cli_a
                fill_to_apply = client_fills.get(cli_a.upper()) or default_new_fills.get(cli_a.upper(), FILL_NEW_CLIENT)
                font_to_apply = FONT_WHITE if cli_a.upper() in ['NOVO A', 'NOVO B'] else copy(ws.cell(dr, dc).font)
                ws.cell(dr, dc).fill = fill_to_apply
                if font_to_apply: ws.cell(dr, dc).font = font_to_apply
                log['realocadas'].setdefault(f"{cli_a} → {acao.bloco_b}-{acao.ambiente_b}", []).append((dr, dc))
                
        # 3. Aloca Cliente B nas vagas disponíveis do Destino A (inclui as recém-liberadas por A)
        if cli_b.upper() not in ('VAZIO', ''):
            free_in_a = [coord for coord in cells_a if cell_values.get(coord, "").upper() in ('VAZIO', '') and coord in allowed_cells]
            free_in_a_sorted = sorted(free_in_a, key=lambda coord: (coord[1], coord[0]))
            
            dests_b = fill_bfs(free_in_a_sorted, len(targets_b))
            
            # TRAVA DE SEGURANÇA B: Alerta perda de posições do Cliente B
            if len(dests_b) < len(targets_b):
                log['avisos'].append(
                    f"ERRO DE CAPACIDADE: Tentou transferir {len(targets_b)} de '{cli_b}' para "
                    f"{acao.bloco_a}-{acao.ambiente_a}, mas só couberam {len(dests_b)}. "
                    f"Perda de {len(targets_b) - len(dests_b)} PAs!"
                )
                
            for dr, dc in dests_b:
                ws.cell(dr, dc).value, cell_values[(dr, dc)] = cli_b, cli_b
                fill_to_apply = client_fills.get(cli_b.upper()) or default_new_fills.get(cli_b.upper(), FILL_NEW_CLIENT)
                font_to_apply = FONT_WHITE if cli_b.upper() in ['NOVO A', 'NOVO B'] else copy(ws.cell(dr, dc).font)
                ws.cell(dr, dc).fill = fill_to_apply
                if font_to_apply: ws.cell(dr, dc).font = font_to_apply
                log['realocadas'].setdefault(f"{cli_b} → {acao.bloco_a}-{acao.ambiente_a}", []).append((dr, dc))

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

def salvar_auditoria(nome_agente: str, input_data: str, output_data: str, safe_name: str):
    dir_auditoria = f"propostas/auditoria_{safe_name}"
    os.makedirs(dir_auditoria, exist_ok=True)
    with open(f"{dir_auditoria}/{nome_agente}_input.txt", "w", encoding="utf-8") as f:
        f.write(input_data)
    with open(f"{dir_auditoria}/{nome_agente}_output.txt", "w", encoding="utf-8") as f:
        f.write(output_data)

def parse_premissas_dinamicas(premissas_txt: str) -> dict:
    """
    Analisa as premissas em texto e extrai dinamicamente os alvos matemáticos (Regex).
    """
    parametros = {
        "reducoes": {},       # cliente -> qtd a reduzir
        "novos_clientes": []   # lista de dicts {"nome": ..., "PAs": ...}
    }
    
    # 1. Busca reduções (ex: "Eliminar 270 posições cliente 1")
    for match in re.finditer(r'(?:eliminar|remover|reduzir)\s+(\d+)\s+posiç\w*\s+cliente\s+(\w+)', premissas_txt, re.IGNORECASE):
        qtd = int(match.group(1))
        cliente = normalize_val(match.group(2).strip())
        parametros["reducoes"][cliente] = qtd
        
    # 2. Busca criação de novos espaços (ex: "Criar um novo espaço apartado com 124 PAs")
    idx_novo = 0
    for match in re.finditer(r'(?:criar\s+um\s+novo\s+espaço|novo\s+espaço)\s+apartado\s+com\s+(\d+)\s+pas', premissas_txt, re.IGNORECASE):
        qtd = int(match.group(1))
        nome_cli = f"NOVO_{chr(65 + idx_novo)}" if idx_novo < 26 else f"NOVO_{idx_novo}"
        parametros["novos_clientes"].append({"nome": nome_cli, "PAs": qtd})
        idx_novo += 1
        
    return parametros

def normalizar_nome_cliente_novo(nome: str, novos_clientes: List[dict]) -> str:
    """
    Normaliza variações de nomes de novos clientes (ex: 'NOVO_ESPACO_A' -> 'NOVO_A')
    fazendo correspondência por sufixo de letra (A, B, C...) ou ordem.
    """
    nome_up = nome.upper().strip()
    if not nome_up.startswith("NOVO"):
        return nome
        
    # Extrai a letra do sufixo (ex: 'NOVO_ESPACO_A' -> 'A', 'NOVO_CLIENTE_B' -> 'B')
    match_letra = re.search(r'\b([A-Z])\b|[-_\s]([A-Z])$', nome_up)
    letra = None
    if match_letra:
        letra = match_letra.group(1) or match_letra.group(2)
        
    if letra:
        # Retorna o nome mapeado esperado para aquela letra (ex: 'NOVO_A' para 'A')
        for nc in novos_clientes:
            if nc["nome"].endswith(f"_{letra}"):
                return nc["nome"]
                
    return nome

def normalizar_acoes(acoes: List[Acao], novos_clientes: List[dict]) -> List[Acao]:
    for acao in acoes:
        # 1. Normaliza campo 'cliente' (usado em liberar e realocar)
        if acao.cliente:
            acao.cliente = normalizar_nome_cliente_novo(acao.cliente, novos_clientes)
        # 2. Normaliza campos 'cliente_a' e 'cliente_b' (usados em transferir)
        if acao.cliente_a:
            acao.cliente_a = normalizar_nome_cliente_novo(acao.cliente_a, novos_clientes)
        if acao.cliente_b:
            acao.cliente_b = normalizar_nome_cliente_novo(acao.cliente_b, novos_clientes)
    return acoes

def validar_inventario(ws_orig, ws_new, allowed_cells, parametros: dict) -> List[str]:
    """
    Garante a integridade do Excel antes de gravar fisicamente comparando dados dinâmicos.
    """
    def obter_contagem(ws):
        counts = {}
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                if (r, c) not in allowed_cells:
                    continue
                val = ws.cell(row=r, column=c).value
                if val is not None:
                    # Normaliza floats e strings usando a função do ScannerPremissas
                    norm_val = normalize_val(val)
                    if norm_val != "" and norm_val not in ('VAZIO', 'CT', 'CATRACA', 'SA', 'SALA', 'CW', 'COWORKING', '##'):
                        counts[norm_val] = counts.get(norm_val, 0) + 1
        return counts

    orig_counts = obter_contagem(ws_orig)
    new_counts = obter_contagem(ws_new)
    erros = []

    # Identifica dinamicamente quem são os clientes que sofrem alteração deliberada
    clientes_excluidos_da_preservacao = set(parametros["reducoes"].keys()) | {nc["nome"] for nc in parametros["novos_clientes"]}

    # 1. Valida preservação exata de todos os clientes estáveis (que não foram alterados nas premissas)
    for client, qty in orig_counts.items():
        if client in clientes_excluidos_da_preservacao: 
            continue
        qty_new = new_counts.get(client, 0)
        if qty_new != qty:
            erros.append(f"Inconsistência no Cliente estável '{client}': quantidade inicial era {qty}, mas agora é {qty_new} (Perda/Ganho de {qty_new - qty} PAs).")

    # 2. Valida alocação exata de cada novo cliente de forma dinâmica
    for nc in parametros["novos_clientes"]:
        nome = nc["nome"]
        esperado = nc["PAs"]
        atual = new_counts.get(nome, 0)
        if atual != esperado:
            erros.append(f"Inconsistência no '{nome}': quantidade exigida é {esperado} PAs, mas atualmente há {atual}.")

    # 3. Valida redução exata do Cliente modificado de forma dinâmica
    for cli, reducao in parametros["reducoes"].items():
        qty_orig = orig_counts.get(cli, 0)
        qty_new = new_counts.get(cli, 0)
        if qty_orig - qty_new != reducao:
            erros.append(f"Redução inválida do Cliente '{cli}': deveria eliminar exatamente {reducao} PAs (restante esperado: {qty_orig - reducao}), mas restaram {qty_new} (redução real de {qty_orig - qty_new} PAs).")

    return erros

# ══════════════════════════════════════════════════════════════════════════
# Fluxo de Execução Principal (Unificado de 2 Etapas com Auto-Correção)
# ══════════════════════════════════════════════════════════════════════════

async def main():
    with open('premissas.txt', encoding='utf-8') as f: 
        premissas_txt = f.read().strip()

    # Extrai dinamicamente os alvos numéricos e nomes de clientes do arquivo TXT
    parametros_premissas = parse_premissas_dinamicas(premissas_txt)

    # Constrói a lista explícita de nomes que a LLM DEVE usar para os novos clientes
    lista_novos_nomes = ""
    if parametros_premissas["novos_clientes"]:
        lista_novos_nomes = "\n".join(
            f"  - Para a nova operação de {nc['PAs']} PAs, use RIGOROSAMENTE o nome: '{nc['nome']}'"
            for nc in parametros_premissas["novos_clientes"]
        )

    regras_nomes_sistema = ""
    if lista_novos_nomes:
        regras_nomes_sistema = (
            f"\n=== DIRETRIZES DE NOMENCLATURA SISTÊMICA (OBRIGATÓRIO) ===\n"
            f"Ao criar ou movimentar os novos clientes, você DEVE usar EXATAMENTE estes nomes:\n"
            f"{lista_novos_nomes}\n"
            f"NÃO invente outros nomes como 'NOVO_CLIENTE_A' ou 'NOVO_CLIENTE_B'. Use apenas as strings exatas fornecidas acima.\n"
        )

    # Extrai o contexto visual das bordas laranjas e junta com as premissas do arquivo
    dados_laranjas = scan_orange_context('planta.xlsx', SHEET_NAME)
    premissas_visuais = build_context_string_for_llm(dados_laranjas)
    
    # Consolida as premissas injetando as diretrizes de nomes que a IA deve respeitar
    premissas_completas = f"{premissas_txt}\n{regras_nomes_sistema}\n{premissas_visuais}"

    # Carrega a planta original e escaneia os dados iniciais
    wb, ws = load_plant()
    plant_data = scan_plant(ws, FORBIDDEN_PATTERNS)
    
    plant_info_str = build_plant_info(plant_data)
    blocos_info_str = build_blocos_info(plant_data, ws.max_row, ws.max_column, ws)
    
    # ── PROTEÇÃO CONTRA USO DE CORREDORES (Apenas mesas físicas reais dentro dos ambientes laranjas) ──
    allowed_cells = set()
    
    # 1. Agrupa as coordenadas de todas as células pertencentes ao interior dos ambientes laranjas
    boundary_cells = set()
    for block in dados_laranjas:
        for env in block.get('ambientes', []):
            boundary_cells.update(env['cells'])
            
    # 2. Permite alocação apenas onde havia mesas válidas (não corredores) na planilha original
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            if (r, c) not in boundary_cells:
                continue  # Exclui cabeçalhos, somas, legendas e margens externas do Excel
                
            cell = ws.cell(row=r, column=c)
            v = cell.value
            if v is not None:
                v_str = str(v).strip()
                if v_str != "" and v_str.upper() not in ('CT', 'CATRACA', 'SA', 'CW', 'COWORKING', '##'):
                    allowed_cells.add((r, c))
    # ─────────────────────────────────────────────────────────────────────────────
        
    # ══════════════════════════════════════════════════════════════════════════
    # ETAPA 1: Execução do Posicionador (Alocação Bruto Direto)
    # ══════════════════════════════════════════════════════════════════════════
    print("1. Posicionador gerando ações primárias (alocação bruta)...")
    pos_deps = PosicionadorDeps(
        plant_info=plant_info_str,
        blocos_info=blocos_info_str,
        premissas=premissas_completas
    )
    res_posicionador = await posicionador.run("Gere a proposta de alocação inicial baseada nas premissas de negócio.", deps=pos_deps)
    rascunho_layout_json = clean_json_string(res_posicionador.output)
    
    # strict=False desativa erro ao encontrar quebras de linha reais nas strings da IA
    pos_data = json.loads(rascunho_layout_json, strict=False)
    
    # Configuração estrita de salvamento sem variação de nomes (proposta_final.xlsx)
    dest_file = "propostas/proposta_final.xlsx"
    safe_name = "final"
    
    os.makedirs('propostas', exist_ok=True)
    shutil.copy("planta.xlsx", dest_file)
    
    new_wb = openpyxl.load_workbook(dest_file)
    new_ws = new_wb[SHEET_NAME]
    
    # Executa primeiro as ações físicas brutas do Posicionador
    acoes_totais = []
    for ac in pos_data.get("acoes_primarias", []):
        acoes_totais.append(Acao(**ac))
        
    # Normaliza as ações primárias antes de executar fisicamente
    acoes_totais = normalizar_acoes(acoes_totais, parametros_premissas["novos_clientes"])
        
    proposta_inicial = PropostaMock(
        nome=pos_data.get("nome", "Alocacao Bruta"),
        custo_obras="Baixo",
        acoes=acoes_totais
    )
    
    print("Executing initial raw layout physically...")
    log_acumulado, _ = execute_alocacao(new_ws, proposta_inicial, plant_data, allowed_cells)
    
    # Salva auditoria inicial do posicionador
    pos_input_str = (
        f"[PREMISSAS DO TXT]\n{pos_deps.premissas}\n\n"
        f"[BLOCOS INFO]\n{pos_deps.blocos_info}\n\n"
        f"[PLANT INFO]\n{pos_deps.plant_info}"
    )
    salvar_auditoria("1_posicionador", pos_input_str, res_posicionador.output, safe_name)

    # Executa validação de integridade física no estado resultante com parâmetros dinâmicos
    erros_validacao = validar_inventario(ws, new_ws, allowed_cells, parametros_premissas)
    
    if not erros_validacao:
        print("✓ Inicial layout validado e gravado com sucesso!")
    else:
        print("⚠ Layout inicial possui erros de inventário. Detalhes:")
        for err in erros_validacao:
            print(f"   ❌ {err}")
        print("\n   👉 Ações que a LLM tentou executar nesta rodada:")
        for ac in pos_data.get("acoes_primarias", []):
            tipo = ac.get('tipo', '').strip().lower()
            if tipo in ('liberar', 'realocar', 'posicionar'):
                print(f"      - [{tipo.upper()}] Cliente '{ac.get('cliente', '')}': Qtd {ac.get('quantidade', 0)} no Bloco {ac.get('bloco', '')}-{ac.get('ambiente', '')}")
        print("\nIniciando loop de auto-correção...")

    # ══════════════════════════════════════════════════════════════════════════
    # ETAPA 2: Loop de Auto-Correção e Swaps Estáveis com o Organizador (Máx 3 Iterações)
    # ══════════════════════════════════════════════════════════════════════════
    print("\n2. Iniciando Loop de Auto-Correção do Organizador (Máximo 3 Iterações)...")
    
    for iteracao in range(1, 4):
        print(f"--- Iteração {iteracao}/3 ---")
        
        # Rescan real-time da planilha física modificada gravada em disco (new_ws)
        current_plant_data = scan_plant(new_ws, FORBIDDEN_PATTERNS)
        current_plant_info = build_plant_info(current_plant_data)
        current_blocos_info = build_blocos_info(current_plant_data, new_ws.max_row, new_ws.max_column, new_ws)
        
        org_deps = OrganizadorDeps(
            plant_info=current_plant_info,
            blocos_info=current_blocos_info,
            premissas=premissas_completas,
            rascunho_layout=rascunho_layout_json
        )
        
        # Injeta os erros matemáticos de forma direta para a auto-correção da IA
        feedback_erros = ""
        if erros_validacao:
            feedback_erros = (
                "\n\n⚠️ INCONSISTÊNCIAS DE SOMA OPERACIONAL (Seu planejamento quebrou o inventário físico nas seguintes regras):\n"
                + "\n".join(f"- {err}" for err in erros_validacao)
                + "\n\nAplique ações de 'transferir' para reequilibrar as quantidades exatas de volta!"
            )
        
        res_organizador = await organizador.run(
            f"Avalie a planta física atual. Se houver qualquer violação de premissas (como unificação, exclusividade, Cliente 7 compartilhado), utilize a função 'transferir' para corrigir. Se estiver tudo certo, retorne a lista vazia.{feedback_erros}",
            deps=org_deps
        )
        reorganizacao_json = clean_json_string(res_organizador.output)
        
        org_input_str = (
            f"[RASCUNHO DO LAYOUT]\n{org_deps.rascunho_layout}\n\n"
            f"[PREMISSAS DO TXT]\n{org_deps.premissas}\n\n"
            f"[BLOCOS INFO ATUALIZADOS NA PLANILHA FISICA]\n{org_deps.blocos_info}\n\n"
            f"[PLANT INFO ATUALIZADA]\n{org_deps.plant_info}"
        )
        salvar_auditoria(f"2_organizador_iter{iteracao}", org_input_str, res_organizador.output, safe_name)
        
        # strict=False desativa erro ao encontrar quebras de linha reais nas strings da IA
        org_data = json.loads(reorganizacao_json, strict=False)
        acoes_org = org_data.get("acoes_organizacao", [])
        
        if not acoes_org and not erros_validacao:
            print(f"✓ Sistema auditado e validado. Todas as premissas e inventários estão plenamente satisfatórios!")
            break
            
        print(f"⚠ Executando ações de correção em sandbox para validação de integridade...")
        
        # Converte as ações de swap sugeridas para validação em sandbox
        acoes_iteracao = []
        for ac in acoes_org:
            acoes_iteracao.append(Acao(**ac))
            
        # Normaliza as ações do organizador antes de executar fisicamente
        acoes_iteracao = normalizar_acoes(acoes_iteracao, parametros_premissas["novos_clientes"])
            
        proposta_correcao = PropostaMock(f"Swaps Iteracao {iteracao}", "Baixo", acoes_iteracao)
        
        # Executa em sandbox temporário (clone do estado atual gravado em disco)
        sandbox_wb, sandbox_ws = clone_ws(new_ws)
        log_iter, _ = execute_alocacao(sandbox_ws, proposta_correcao, current_plant_data, allowed_cells)
        
        # Roda validação matemática dinâmica no sandbox pós-correção
        erros_validacao_temp = validar_inventario(ws, sandbox_ws, allowed_cells, parametros_premissas)
        
        if not erros_validacao_temp:
            # Commit: Se o sandbox passou nas regras de integridade, aplica as alterações na planilha master física
            print(f"✓ Correção da Iteração {iteracao} validada com sucesso! Gravando no Excel master...")
            log_iter, _ = execute_alocacao(new_ws, proposta_correcao, current_plant_data, allowed_cells)
            
            # Mescla os logs de auditoria finais
            for k, v in log_iter.get('realocadas', {}).items():
                log_acumulado.setdefault('realocadas', {}).setdefault(k, []).extend(v)
            for k, v in log_iter.get('liberadas', {}).items():
                log_acumulado.setdefault('liberadas', {})[k] = v
            log_acumulado.setdefault('avisos', []).extend(log_iter.get('avisos', []))
            acoes_totais.extend(acoes_iteracao)
            erros_validacao = []  # Zera a fila de erros
        else:
            # Rollback: Mantém a planilha íntegra do loop anterior e detalha as falhas no console
            print(f"❌ Correção inválida na Iteração {iteracao}. Executando rollback. Detalhes:")
            for err in erros_validacao_temp:
                print(f"   -> {err}")
            print("\n   👉 Ações que a LLM tentou executar nesta rodada:")
            for ac in acoes_org:
                tipo = ac.get('tipo', '').strip().lower()
                if tipo in ('liberar', 'realocar', 'posicionar'):
                    print(f"      - [{tipo.upper()}] Cliente '{ac.get('cliente', '')}': Qtd {ac.get('quantidade', 0)} no Bloco {ac.get('bloco', '')}-{ac.get('ambiente', '')}")
                elif tipo in ('transferir', 'swap', 'permutar'):
                    print(f"      - [TRANSFERIR] '{ac.get('cliente_a', '')}' ({ac.get('quantidade_a', 0)} PAs de {ac.get('bloco_a', '')}-{ac.get('ambiente_a', '')}) <-> '{ac.get('cliente_b', '')}' ({ac.get('quantidade_b', 0)} PAs de {ac.get('bloco_b', '')}-{ac.get('ambiente_b', '')})")
            erros_validacao = erros_validacao_temp  # Mantém os erros ativos para a próxima iteração

    # ══════════════════════════════════════════════════════════════════════════
    # ETAPA 3: Finalização e Escrita de Relatórios
    # ══════════════════════════════════════════════════════════════════════════
    print("\n3. Gravando resultados e relatórios finais...")
    new_wb.save(dest_file)
    
    # Recalcula o relatório final de blocos
    final_plant_data = scan_plant(new_ws, FORBIDDEN_PATTERNS)
    final_blocos_info = build_blocos_info(final_plant_data, new_ws.max_row, new_ws.max_column, new_ws)
    
    print("\n================================================================================")
    print("ESTADO FINAL DOS BLOCOS APÓS TODAS AS MOVIMENTAÇÕES E CORREÇÕES:")
    print("================================================================================")
    print(final_blocos_info)
    
    proposta_final = PropostaMock(
        nome=pos_data.get("nome", "Reorganizacao Consolidada"),
        custo_obras="Baixo",
        acoes=acoes_totais
    )
    
    write_report(ws, new_ws, proposta_final, log_acumulado, "propostas/proposta_final_mudancas.txt")
    
    caminho_relatorio_final = "propostas/proposta_final_blocos_finais.txt"
    with open(caminho_relatorio_final, "w", encoding="utf-8") as f:
        f.write("================================================================================\n")
        f.write("RELATÓRIO DO ESTADO FINAL DOS BLOCOS APÓS AS MOVIMENTAÇÕES\n")
        f.write("================================================================================\n\n")
        f.write(final_blocos_info)
        
    print(f"✓ Processo concluído com sucesso. Arquivos gravados em 'propostas/':")
    print(f"  - Excel modificado: proposta_final.xlsx")
    print(f"  - Lista de mudanças: proposta_final_mudancas.txt")
    print(f"  - Mapa final de blocos: proposta_final_blocos_finais.txt")

if __name__ == '__main__':
    asyncio.run(main())
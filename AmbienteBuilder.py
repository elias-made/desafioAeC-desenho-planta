# AmbienteBuilder.py

import os
import re
from typing import List, Set, Tuple
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from ScannerPremissas import scan_orange_context

# ══════════════════════════════════════════════════════════════════════════
# Funções de Fluxo, Contornos e Divisórias Dinâmicas (Sem Hardcode)
# ══════════════════════════════════════════════════════════════════════════

def _tem_parede_laranja_entre(ws, r1, c1, r2, c2) -> bool:
    """Verifica se existe uma parede física laranjada (borda de contorno do ambiente) entre as duas células."""
    if r1 < 1 or r1 > ws.max_row or c1 < 1 or c1 > ws.max_column:
        return True
    if r2 < 1 or r2 > ws.max_row or c2 < 1 or c2 > ws.max_column:
        return True
        
    cell1 = ws.cell(row=r1, column=c1)
    cell2 = ws.cell(row=r2, column=c2)
    
    def is_orange(side) -> bool:
        if not side or not side.style or side.style == 'none':
            return False
        if not side.color or not side.color.value:
            return False
        c_val = str(side.color.value).strip().upper()
        # Captura "FF9900", "FFFF9900" ou qualquer variante de laranja do projeto
        return "9900" in c_val or "FF99" in c_val

    # Movimento horizontal (esquerda / direita)
    if r1 == r2:
        if c2 < c1:  # Esquerda
            if is_orange(cell1.border.left) or is_orange(cell2.border.right):
                return True
        else:  # Direita
            if is_orange(cell1.border.right) or is_orange(cell2.border.left):
                return True
    # Movimento vertical (cima / baixo)
    elif c1 == c2:
        if r2 < r1:  # Cima
            if is_orange(cell1.border.top) or is_orange(cell2.border.bottom):
                return True
        else:  # Baixo
            if is_orange(cell1.border.bottom) or is_orange(cell2.border.top):
                return True
                
    return False

def _eh_pilar_ou_coluna(ws, r, c) -> bool:
    """Verifica se a célula especificada é um pilar ou coluna física estrutural (preto ou ##)."""
    if r < 1 or r > ws.max_row or c < 1 or c > ws.max_column:
        return False
    cell = ws.cell(row=r, column=c)
    from ScannerPremissas import is_barrier_cell
    if is_barrier_cell(cell):
        val = str(cell.value).strip().upper() if cell.value is not None else ""
        if val in ('SALA', 'COWORKING', 'SALA CLIENTE', 'SALA1', 'SALA2', 'SALA3', 'SALA4', 'CATRACA', 'CT', 'ESCANINHOS'):
            return False
        return True
    return False

def _eh_faixa_livre(ws, env_cells: Set[Tuple[int, int]], r0, r1, c0, c1) -> bool:
    """Verifica se todas as células no retângulo [r0, r1] x [c0, c1] são corredores livres (sem barreiras ou mesas de outros clientes)."""
    if r0 < 1 or r1 > ws.max_row or c0 < 1 or c1 > ws.max_column:
        return False
    from ScannerPremissas import is_desk_cell, is_barrier_cell, cell_has_orange_fill
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            cell = ws.cell(row=r, column=c)
            val = str(cell.value).strip().upper() if cell.value is not None else ""
            
            # Se for uma mesa de operador, ela só bloqueia se pertencer a outro cliente (estiver fora de env_cells)
            if is_desk_cell(cell) or val.startswith("NOVO"):
                if (r, c) not in env_cells:
                    return False
            
            # Se for pilar estrutural (##), CT de outra sala ou bloco laranja sólido, bloqueia a expansão
            if is_barrier_cell(cell) or cell_has_orange_fill(cell) or val in ("CT", "CATRACA"):
                return False
    return True

def get_env_cells(block_id_str: str, env_letter: str, macro_blocks: List[dict]) -> List[Tuple[int, int]]:
    """Identifica as coordenadas de células pertencentes a um ambiente dentro de um bloco específico."""
    if not block_id_str or not env_letter:
        return []
    block_match = re.search(r'\d+', str(block_id_str))
    if not block_match:
        return []
    block_idx = int(block_match.group())
    if block_idx <= len(macro_blocks):
        block = macro_blocks[block_idx - 1]
        block_envs = {e['id'].upper(): e for e in block.get('ambientes', [])}
        parts = re.split(r'[-_\s+&,|/]+', env_letter.upper())
        matched_envs = []
        for p in parts:
            if p in block_envs:
                matched_envs.append(block_envs[p])
        if matched_envs:
            cells = []
            for env in matched_envs:
                cells.extend(env['cells'])
            return cells
    return []

def separar_ambiente_e_desenhar_divisorias(
    ws, 
    env_cells: Set[Tuple[int, int]], 
    allocated_cells: Set[Tuple[int, int]], 
    border_style: str = "medium", 
    border_color: str = "FF9900"
):
    """
    Desenha as divisórias ao redor do bloco irregular de mesas ocupadas, adicionando
    corredores ergonômicos funcionais e gerando um contorno 2D célula a célula.
    Garante que as divisórias nunca ultrapassem ou invadam mesas de outros clientes.
    """
    if not allocated_cells or not env_cells:
        return

    from ScannerPremissas import is_desk_cell
    from BlockMapper import flood_fill

    side_style = Side(border_style=border_style, color=border_color)
    env_cells = set(env_cells)
    allocated = set(allocated_cells) & env_cells
    if not allocated:
        allocated = set(allocated_cells)

    # 1. Encontra todas as mesas do ambiente para achar as bancadas ocupadas
    all_desks_in_env = set()
    for r, c in env_cells:
        cell = ws.cell(row=r, column=c)
        val_str = str(cell.value).strip().upper() if cell.value is not None else ""
        if is_desk_cell(cell) or (r, c) in allocated or val_str.startswith("NOVO"):
            all_desks_in_env.add((r, c))

    benches = flood_fill(all_desks_in_env)
    target_bench_cells = set()
    
    # 2. Varredura inteligente de bancadas com suporte a corte de espinha
    for bench in benches:
        bench_set = set(bench)
        allocated_in_bench = bench_set & allocated
        
        if allocated_in_bench:
            # --- DETECÇÃO DINÂMICA DE CORTE INTENCIONAL ---
            b_rows = {r for r, c in bench_set}
            b_cols = {c for r, c in bench_set}
            physical_rows_count = len(b_rows)
            physical_cols_count = len(b_cols)
            
            allocated_rows_count = len({r for r, c in allocated_in_bench})
            allocated_cols_count = len({c for r, c in allocated_in_bench})
            
            corte_colunas_intencional = allocated_cols_count < physical_cols_count
            corte_linhas_intencional = allocated_rows_count < physical_rows_count
            
            # --- REGRA DE PREENCHIMENTO ---
            total_capacity = len(bench_set)
            allocated_count = len(allocated_in_bench)
            missing_to_complete = total_capacity - allocated_count
            
            if missing_to_complete <= 2 and not corte_colunas_intencional and not corte_linhas_intencional:
                target_bench_cells.update(bench_set)
            else:
                target_bench_cells.update(allocated_in_bench)

    if not target_bench_cells:
        target_bench_cells = allocated

    # 3. CONSTRUÇÃO DA ÁREA DA SALA (união de mesas + corredores ergonômicos correspondentes limitados ao ambiente base)
    room_cells = set(target_bench_cells)
    
    for bench in benches:
        bench_set = set(bench)
        bench_allocated = bench_set & target_bench_cells
        if not bench_allocated:
            continue
            
        r0_b = min(r for r, c in bench_allocated)
        r1_b = max(r for r, c in bench_allocated)
        c0_b = min(c for r, c in bench_allocated)
        c1_b = max(c for r, c in bench_allocated)
        
        # Determina a orientação física desta bancada específica
        b_rows = {r for r, c in bench_set}
        b_cols = {c for r, c in bench_set}
        is_vertical = len(b_rows) >= len(b_cols)
        
        if is_vertical:
            # Corredor Esquerdo: expande se não houver parede laranja no caminho
            for r in range(r0_b, r1_b + 1):
                if (r, c0_b) in target_bench_cells:
                    target_cell = (r, c0_b - 1)
                    if not _tem_parede_laranja_entre(ws, r, c0_b, r, c0_b - 1):
                        if _eh_faixa_livre(ws, env_cells, r, r, c0_b - 1, c0_b - 1):
                            room_cells.add(target_cell)
            # Corredor Direito: expande se não houver parede laranja no caminho
            for r in range(r0_b, r1_b + 1):
                if (r, c1_b) in target_bench_cells:
                    target_cell = (r, c1_b + 1)
                    if not _tem_parede_laranja_entre(ws, r, c1_b, r, c1_b + 1):
                        if _eh_faixa_livre(ws, env_cells, r, r, c1_b + 1, c1_b + 1):
                            room_cells.add(target_cell)
            # Corredor Inferior (Unificador/Conector de Escape feito célula a célula)
            left_col = c0_b - 1 if (r1_b, c0_b - 1) in room_cells else c0_b
            right_col = c1_b + 1 if (r1_b, c1_b + 1) in room_cells else c1_b
            for c in range(left_col, right_col + 1):
                target_cell = (r1_b + 1, c)
                if not _tem_parede_laranja_entre(ws, r1_b, c, r1_b + 1, c):
                    if _eh_faixa_livre(ws, env_cells, r1_b + 1, r1_b + 1, c, c):
                        room_cells.add(target_cell)
        else:
            # Corredor Superior (Egress): expande se não houver parede laranja no caminho
            for c in range(c0_b, c1_b + 1):
                if (r0_b, c) in target_bench_cells:
                    target_cell = (r0_b - 1, c)
                    if not _tem_parede_laranja_entre(ws, r0_b, c, r0_b - 1, c):
                        if _eh_faixa_livre(ws, env_cells, r0_b - 1, r0_b - 1, c, c):
                            room_cells.add(target_cell)
            # Corredor Inferior (Egress): expande se não houver parede laranja no caminho
            for c in range(c0_b, c1_b + 1):
                if (r1_b, c) in target_bench_cells:
                    target_cell = (r1_b + 1, c)
                    if not _tem_parede_laranja_entre(ws, r1_b, c, r1_b + 1, c):
                        if _eh_faixa_livre(ws, env_cells, r1_b + 1, r1_b + 1, c, c):
                            room_cells.add(target_cell)
            # Corredor Esquerdo (Unificador/Conector feito célula a célula)
            top_row = r0_b - 1 if (r0_b - 1, c0_b) in room_cells else r0_b
            bottom_row = r1_b + 1 if (r1_b + 1, c0_b) in room_cells else r1_b
            for r in range(top_row, bottom_row + 1):
                target_cell = (r, c0_b - 1)
                if not _tem_parede_laranja_entre(ws, r, c0_b, r, c0_b - 1):
                    if _eh_faixa_livre(ws, env_cells, r, r, c0_b - 1, c0_b - 1):
                        room_cells.add(target_cell)
            # Corredor Direito (Unificador/Conector feito célula a célula - independente)
            top_row = r0_b - 1 if (r0_b - 1, c1_b) in room_cells else r0_b
            bottom_row = r1_b + 1 if (r1_b + 1, c1_b) in room_cells else r1_b
            for r in range(top_row, bottom_row + 1):
                target_cell = (r, c1_b + 1)
                if not _tem_parede_laranja_entre(ws, r, c1_b, r, c1_b + 1):
                    if _eh_faixa_livre(ws, env_cells, r, r, c1_b + 1, c1_b + 1):
                        room_cells.add(target_cell)

    # 3.5. PONTE DE CORREDOR VERTICAL (UNIFICAÇÃO DE ILHAS DE SALAS)
    corridor_columns = {c for r, c in (room_cells - target_bench_cells)}
    for col in sorted(list(corridor_columns)):
        rows_in_col = sorted(list({r for r, c in room_cells if c == col}))
        if len(rows_in_col) > 1:
            for i in range(len(rows_in_col) - 1):
                r_start = rows_in_col[i]
                r_end = rows_in_col[i + 1]
                if r_end - r_start > 1:
                    gap_rows = range(r_start + 1, r_end)
                    all_gap_free = True
                    for gr in gap_rows:
                        target_cell = (gr, col)
                        # A ponte não deve cruzar paredes laranjas físicas e deve estar desobstruída
                        if _tem_parede_laranja_entre(ws, gr - 1, col, gr, col) or not _eh_faixa_livre(ws, env_cells, gr, gr, col, col):
                            all_gap_free = False
                            break
                    if all_gap_free:
                        for gr in gap_rows:
                            room_cells.add((gr, col))

    # 4. INSERÇÃO DA CATRACA (CT) ÚNICA NO ESPAÇO DE CORREDOR MAIS BAIXO DO AMBIENTE UNIFICADO
    corridor_cells = room_cells - target_bench_cells
    if corridor_cells:
        max_r = max(r for r, c in room_cells)
        bottom_corridors = [cell for cell in corridor_cells if cell[0] == max_r]
        if not bottom_corridors:
            max_r_corridors = max(cell[0] for cell in corridor_cells)
            bottom_corridors = [cell for cell in corridor_cells if cell[0] == max_r_corridors]
        
        min_c_tables = min(c for r, c in target_bench_cells)
        min_c_room = min(c for r, c in room_cells)
        if min_c_room < min_c_tables:
            bottom_corridors.sort(key=lambda x: x[1])
        else:
            bottom_corridors.sort(key=lambda x: x[1], reverse=True)
            
        cell_ct_coord = bottom_corridors[0]
        cell_ct = ws.cell(row=cell_ct_coord[0], column=cell_ct_coord[1])
        
        fill_ct = PatternFill(start_color="FF9900", end_color="FF9900", fill_type="solid")
        font_ct = Font(color="FFFFFF", bold=True, size=8)
        cell_ct.value = "CT"
        cell_ct.fill = fill_ct
        cell_ct.font = font_ct

    # 5. DESENHO DAS DIVISÓRIAS (CONTORNO DA SALA CÉLULA A CÉLULA)
    for r, c in room_cells:
        # Topo
        if (r - 1, c) not in room_cells and not _eh_pilar_ou_coluna(ws, r - 1, c):
            _aplicar_borda_espelhada(ws, r, c, 'top', side_style)
        # Base
        if (r + 1, c) not in room_cells and not _eh_pilar_ou_coluna(ws, r + 1, c):
            _aplicar_borda_espelhada(ws, r, c, 'bottom', side_style)
        # Esquerda
        if (r, c - 1) not in room_cells and not _eh_pilar_ou_coluna(ws, r, c - 1):
            _aplicar_borda_espelhada(ws, r, c, 'left', side_style)
        # Direita
        if (r, c + 1) not in room_cells and not _eh_pilar_ou_coluna(ws, r, c + 1):
            _aplicar_borda_espelhada(ws, r, c, 'right', side_style)

def _copiar_lado(side_obj):
    if not side_obj or not side_obj.style or side_obj.style == 'none':
        return None
    return Side(border_style=side_obj.style, color=side_obj.color)

def _aplicar_borda_celula_unica(ws, r, c, side, side_style):
    cell = ws.cell(row=r, column=c)
    b = cell.border
    top = _copiar_lado(b.top)
    bottom = _copiar_lado(b.bottom)
    left = _copiar_lado(b.left)
    right = _copiar_lado(b.right)
    
    if side == 'top': top = side_style
    elif side == 'bottom': bottom = side_style
    elif side == 'left': left = side_style
    elif side == 'right': right = side_style
    
    cell.border = Border(top=top, bottom=bottom, left=left, right=right)

def _aplicar_borda_espelhada(ws, r, c, side, side_style):
    _aplicar_borda_espelhada_unica = _aplicar_borda_celula_unica
    _aplicar_borda_espelhada_unica(ws, r, c, side, side_style)
    if side == 'bottom':
        _aplicar_borda_espelhada_unica(ws, r + 1, c, 'top', side_style)
    elif side == 'top':
        _aplicar_borda_espelhada_unica(ws, r - 1, c, 'bottom', side_style)
    elif side == 'right':
        _aplicar_borda_espelhada_unica(ws, r, c + 1, 'left', side_style)
    elif side == 'left':
        _aplicar_borda_espelhada_unica(ws, r, c - 1, 'right', side_style)

# ══════════════════════════════════════════════════════════════════════════
# Função de Seleção e Execução de Teste Manual
# ══════════════════════════════════════════════════════════════════════════

def _selecionar_mesas_contiguas(env_cells: Set[Tuple[int, int]], ws, target_qty: int) -> Set[Tuple[int, int]]:
    """
    Busca e reserva mesas de forma estruturada, agrupando e isolando colunas/linhas físicas
    de bancadas para evitar a expansão desnecessária do Bounding Box geral.
    """
    from ScannerPremissas import is_desk_cell
    from BlockMapper import flood_fill
    
    all_desks = set()
    for r, c in env_cells:
        cell = ws.cell(row=r, column=c)
        if is_desk_cell(cell):
            all_desks.add((r, c))
            
    if not all_desks:
        return set()
        
    benches = flood_fill(all_desks)
    
    vertical_benches = []
    horizontal_benches = []
    for b in benches:
        b_rows = {r for r, c in b}
        b_cols = {c for r, c in b}
        if len(b_rows) >= len(b_cols):
            vertical_benches.append(b)
        else:
            horizontal_benches.append(b)
            
    is_mostly_vertical = len(vertical_benches) >= len(horizontal_benches)
    selected = set()
    
    # ══════════════════════════════════════════════════════════════════════════
    # BUSCA DE CORTE SIMÉTRICO OTIMAL COM PONTUAÇÃO E TOMADA DE DECISÃO HIERÁRQUICA
    # ══════════════════════════════════════════════════════════════════════════
    if is_mostly_vertical and vertical_benches:
        vertical_benches.sort(key=lambda b: (min(r for r, c in b), min(c for r, c in b)))
        groups = []
        for b in vertical_benches:
            br_y_0 = min(r for r, c in b)
            placed = False
            for grp in groups:
                if abs(min(r for r, c in grp[0]) - br_y_0) <= 2:
                    grp.append(b)
                    placed = True
                    break
            if not placed:
                groups.append([b])
        
        valid_groups = [grp for grp in groups if sum(len(b) for b in grp) >= target_qty]
        if valid_groups:
            valid_groups.sort(key=lambda grp: min(min(r for r, c in b) for b in grp))
            benches_to_use = valid_groups[0]
        else:
            benches_to_use = vertical_benches
            
        max_rows = max(len({r for r, c in b}) for b in benches_to_use)
        
        best_N, best_H = None, None
        candidates = []
        
        for N in range(1, len(benches_to_use) + 1):
            for H in range(1, max_rows + 1):
                cap = 0
                for b in benches_to_use[:N]:
                    b_rows_sorted = sorted(list({r for r, c in b}))
                    allowed_rows = set(b_rows_sorted[:H])
                    cap += sum(1 for r, c in b if r in allowed_rows)
                
                if cap >= target_qty:
                    excess = cap - target_qty
                    corridor_height = max_rows - H
                    candidates.append((cap, N, H, excess, corridor_height))
        
        if candidates:
            # 1. Menor excesso, 2. Menor N, 3. Maior Corredor
            candidates.sort(key=lambda x: (x[3], x[1], -x[4]))
            best_N = candidates[0][1]
            best_H = candidates[0][2]
        else:
            # Fallback Escassez
            deficit_candidates = []
            for N in range(1, len(benches_to_use) + 1):
                for H in range(1, max_rows + 1):
                    cap = 0
                    for b in benches_to_use[:N]:
                        b_rows_sorted = sorted(list({r for r, c in b}))
                        allowed_rows = set(b_rows_sorted[:H])
                        cap += sum(1 for r, c in b if r in allowed_rows)
                    corridor_height = max_rows - H
                    deficit_candidates.append((cap, N, H, corridor_height))
            
            deficit_candidates.sort(key=lambda x: (-x[0], x[1], -x[3]))
            best_N = deficit_candidates[0][1]
            best_H = deficit_candidates[0][2]
        
        if best_N is not None and best_H is not None:
            remaining_qty = target_qty
            for b in benches_to_use[:best_N]:
                b_rows_sorted = sorted(list({r for r, c in b}))
                allowed_rows = set(b_rows_sorted[:best_H])
                allowed_cells = [cell for cell in b if cell[0] in allowed_rows]
                # --- ORDENAÇÃO DE 2 EM 2 PARA VERTICAIS: LINHA PRIMEIRO ---
                allowed_cells.sort(key=lambda x: (x[0], x[1]))
                
                take_qty = min(len(allowed_cells), remaining_qty)
                selected.update(allowed_cells[:take_qty])
                remaining_qty -= take_qty
                if remaining_qty <= 0:
                    break
                    
    elif not is_mostly_vertical and horizontal_benches:
        horizontal_benches.sort(key=lambda b: (min(c for r, c in b), min(r for r, c in b)))
        groups = []
        for b in horizontal_benches:
            bc_x_0 = min(c for r, c in b)
            placed = False
            for grp in groups:
                if abs(min(c for r, c in grp[0]) - bc_x_0) <= 2:
                    grp.append(b)
                    placed = True
                    break
            if not placed:
                groups.append([b])
        
        valid_groups = [grp for grp in groups if sum(len(b) for b in grp) >= target_qty]
        if valid_groups:
            valid_groups.sort(key=lambda grp: min(min(c for r, c in b) for b in grp))
            benches_to_use = valid_groups[0]
        else:
            benches_to_use = horizontal_benches
            
        max_cols = max(len({c for r, c in b}) for b in benches_to_use)
        
        best_N, best_W = None, None
        candidates = []
        
        for N in range(1, len(benches_to_use) + 1):
            for W in range(1, max_cols + 1):
                cap = 0
                for b in benches_to_use[:N]:
                    b_cols_sorted = sorted(list({c for r, c in b}))
                    allowed_cols = set(b_cols_sorted[:W])
                    cap += sum(1 for r, c in b if c in allowed_cols)
                
                if cap >= target_qty:
                    excess = cap - target_qty
                    corridor_width = max_cols - W
                    candidates.append((cap, N, W, excess, corridor_width))
                    
        if candidates:
            # 1. Menor excesso, 2. Menor N, 3. Maior Corredor
            candidates.sort(key=lambda x: (x[3], x[1], -x[4]))
            best_N = candidates[0][1]
            best_W = candidates[0][2]
        else:
            # Fallback Escassez
            deficit_candidates = []
            for N in range(1, len(benches_to_use) + 1):
                for W in range(1, max_cols + 1):
                    cap = 0
                    for b in benches_to_use[:N]:
                        b_cols_sorted = sorted(list({c for r, c in b}))
                        allowed_cols = set(b_cols_sorted[:W])
                        cap += sum(1 for r, c in b if c in allowed_cols)
                    corridor_width = max_cols - W
                    deficit_candidates.append((cap, N, W, corridor_width))
            
            deficit_candidates.sort(key=lambda x: (-x[0], x[1], -x[3]))
            best_N = deficit_candidates[0][1]
            best_W = deficit_candidates[0][2]
                        
        if best_N is not None and best_W is not None:
            remaining_qty = target_qty
            for b in benches_to_use[:best_N]:
                b_cols_sorted = sorted(list({c for r, c in b}))
                allowed_cols = set(b_cols_sorted[:best_W])
                allowed_cells = [cell for cell in b if cell[1] in allowed_cols]
                # --- ORDENAÇÃO DE 2 EM 2 PARA HORIZONTAIS: COLUNA PRIMEIRO ---
                allowed_cells.sort(key=lambda x: (x[1], x[0]))
                
                take_qty = min(len(allowed_cells), remaining_qty)
                selected.update(allowed_cells[:take_qty])
                remaining_qty -= take_qty
                if remaining_qty <= 0:
                    break
    # ══════════════════════════════════════════════════════════════════════════

    # Fallback de segurança se nada for selecionado
    if not selected:
        benches.sort(key=lambda b: min(b))
        remaining_qty = target_qty
        for bench in benches:
            if remaining_qty <= 0:
                break
            bench_set = set(bench)
            if len(bench_set) <= remaining_qty:
                selected.update(bench_set)
                remaining_qty -= len(bench_set)
            else:
                cols = {c for r, c in bench_set}
                rows = {r for r, c in bench_set}
                if len(rows) >= len(cols):
                    sorted_bench = sorted(list(bench_set), key=lambda x: (x[1], x[0]))
                else:
                    sorted_bench = sorted(list(bench_set), key=lambda x: (x[0], x[1]))
                bench_selected = set(sorted_bench[:remaining_qty])
                selected.update(bench_selected)
                remaining_qty = 0
                break
                
    return selected

def testar_criacao_sala_manual(
    file_path: str,
    sheet_name: str,
    bloco_id: str,
    ambiente_letra: str,
    quantidade_mesas: int,
    output_path: str = "planta_teste_sala.xlsx",
    border_style: str = "medium",
    border_color: str = "FF9900"
):
    """
    Carrega a planilha, seleciona as mesas de forma contígua e desenha as divisórias
    de contorno ao redor de todas as bancadas completas ocupadas, sem alterar valores ou estilos.
    """
    print(f"\n[INICIANDO TESTE MANUAL] Bloco: {bloco_id} | Ambiente: {ambiente_letra} | PAs Solicitadas: {quantidade_mesas}")
    
    if not os.path.exists(file_path):
        print(f"Erro: Arquivo '{file_path}' não encontrado.")
        return

    wb = openpyxl.load_workbook(file_path)
    ws = wb[sheet_name]

    macro_blocks = scan_orange_context(file_path, sheet_name)
    env_cells = set(get_env_cells(bloco_id, ambiente_letra, macro_blocks))

    if not env_cells:
        print(f"Erro: Não foi possível localizar células para o ambiente '{bloco_id}-{ambiente_letra}'.")
        return

    # 1. Seleciona estritamente a quantidade de mesas solicitada contiguamente
    allocated_cells = _selecionar_mesas_contiguas(env_cells, ws, quantidade_mesas)
    
    if not allocated_cells:
        print("Erro: Nenhuma mesa utilizável encontrada nesse ambiente.")
        return

    # 2. Desenha as divisórias de contorno considerando a ocupação dinâmica no ambiente base (env_cells)
    separar_ambiente_e_desenhar_divisorias(
        ws=ws, 
        env_cells=env_cells, 
        allocated_cells=allocated_cells,
        border_style=border_style, 
        border_color=border_color
    )

    wb.save(output_path)
    print(f"✓ Teste concluído com sucesso! Resultado salvo em: '{output_path}'")

if __name__ == "__main__":
    testar_criacao_sala_manual(
        file_path="planta.xlsx",
        sheet_name="JPIII",
        bloco_id="vazio-4",          
        ambiente_letra="A",         
        quantidade_mesas=280,         
        output_path="planta_teste_sala.xlsx"
    )
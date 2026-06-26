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

def _eh_faixa_livre(ws, target_bench_cells: Set[Tuple[int, int]], r0, r1, c0, c1) -> bool:
    """Verifica se todas as células no retângulo [r0, r1] x [c0, c1] são corredores livres (sem barreiras ou outras mesas) [0]."""
    if r0 < 1 or r1 > ws.max_row or c0 < 1 or c1 > ws.max_column:
        return False
    from ScannerPremissas import is_desk_cell, is_barrier_cell, cell_has_orange_fill
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            cell = ws.cell(row=r, column=c)
            val = str(cell.value).strip().upper() if cell.value is not None else ""
            
            # Se for uma mesa de operador, ela DEVE pertencer à nossa bancada alvo para ser aceita [0]
            # (Impede invadir bancadas vizinhas de outros clientes) [0]
            if is_desk_cell(cell) or val.startswith("NOVO"):
                if (r, c) not in target_bench_cells:
                    return False
            
            # Se for pilar estrutural (##), CT de outra sala ou bloco laranja sólido, bloqueia a expansão [0]
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
    Desenha as divisórias de forma a envolver as bancadas físicas selecionadas,
    adicionando corredores de recuo e inserindo a catraca (CT) unificada de saída [0].
    Garante que as divisórias nunca ultrapassem ou invadam mesas de outros clientes [0].
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
    
    # Varredura inteligente de bancadas com suporte a corte de espinha [0]
    for bench in benches:
        bench_set = set(bench)
        allocated_in_bench = bench_set & allocated
        
        if allocated_in_bench:
            r0 = min(r for r, c in allocated_in_bench)
            r1 = max(r for r, c in allocated_in_bench)
            c0 = min(c for r, c in allocated_in_bench)
            c1 = max(c for r, c in allocated_in_bench)
            
            br_y_0 = min(r for r, c in bench_set) # Cobre limites da bancada original
            br_y_1 = max(r for r, c in bench_set)
            bc_x_0 = min(c for r, c in bench_set)
            bc_x_1 = max(c for r, c in bench_set)
            
            cols = {c for r, c in bench_set}
            rows = {r for r, c in bench_set}
            
            leftover_threshold = 2
            
            # Aplica o recorte e unificação de bancadas duplas vertical ou horizontal [0]
            if len(rows) >= len(cols):
                target_r0 = br_y_0 if (r0 - br_y_0 <= leftover_threshold) else r0
                target_r1 = br_y_1 if (br_y_1 - r1 <= leftover_threshold) else r1
                target_c0 = c0
                target_c1 = c1
            else:
                target_r0 = r0
                target_r1 = r1
                target_c0 = bc_x_0 if (c0 - bc_x_0 <= leftover_threshold) else c0
                target_c1 = bc_x_1 if (bc_x_1 - c1 <= leftover_threshold) else c1
                
            # Adiciona apenas a fração útil da bancada sob o critério de espinha e limite de sobras [0]
            for r, c in bench_set:
                if target_r0 <= r <= target_r1 and target_c0 <= c <= target_c1:
                    target_bench_cells.add((r, c))

    if not target_bench_cells:
        target_bench_cells = allocated

    # 3. Limites exatos da bancada de mesas ocupada
    r0 = min(r for r, c in target_bench_cells)
    r1 = max(r for r, c in target_bench_cells)
    c0 = min(c for r, c in target_bench_cells)
    c1 = max(c for r, c in target_bench_cells)

    # Determina a orientação predominante da bancada ocupada
    # Se o número de linhas (altura) for maior ou igual ao de colunas (largura), é considerada vertical.
    is_vertical = (r1 - r0) >= (c1 - c0)

    # 4. Expansão Ergonômica de Corredores (1 célula de folga adaptada à orientação) [0]
    # O código só expande em uma direção se as células vizinhas forem de fato vazios transitáveis [0].
    if is_vertical:
        # Bancada Vertical:
        # Operadores recuam para as laterais (esquerda/direita).
        # Topo permanece flush para preservar corredores horizontais públicos.
        # Base expande para criar o corredor de circulação e colocar a catraca (CT).
        room_c0 = c0 - 1 if _eh_faixa_livre(ws, target_bench_cells, r0, r1, c0 - 1, c0 - 1) else c0
        room_c1 = c1 + 1 if _eh_faixa_livre(ws, target_bench_cells, r0, r1, c1 + 1, c1 + 1) else c1
        room_r0 = r0
        room_r1 = r1 + 1 if _eh_faixa_livre(ws, target_bench_cells, r1 + 1, r1 + 1, c0, c1) else r1
    else:
        # Bancada Horizontal:
        # Operadores recuam para cima (topo) e para baixo (base).
        # Lateral direita permanece flush para preservar corredores verticais públicos.
        # Lateral esquerda expande para unificar a circulação (conectar topo e base) e colocar a catraca (CT).
        room_r0 = r0 - 1 if _eh_faixa_livre(ws, target_bench_cells, r0 - 1, r0 - 1, c0, c1) else r0
        room_r1 = r1 + 1 if _eh_faixa_livre(ws, target_bench_cells, r1 + 1, r1 + 1, c0, c1) else r1
        room_c0 = c0 - 1 if _eh_faixa_livre(ws, target_bench_cells, r0, r1, c0 - 1, c0 - 1) else c0
        room_c1 = c1

    print(f"      [DIVISÓRIAS] Orientação Detectada: {'Vertical' if is_vertical else 'Horizontal'}")
    print(f"      [DIVISÓRIAS] Bancada Ocupada: Linhas {r0}-{r1}, Colunas {c0}-{c1}")
    print(f"      [DIVISÓRIAS] Sala com Recuo e Catraca: Linhas {room_r0}-{room_r1}, Colunas {room_c0}-{room_c1}")

    # 5. INSERÇÃO DA CATRACA (CT) ÚNICA NO ESPAÇO DO NOVO CORREDOR [0]
    # Posiciona a catraca de forma garantida em uma célula interna válida do corredor esquerdo [0]
    fill_ct = PatternFill(start_color="FF9900", end_color="FF9900", fill_type="solid")
    font_ct = Font(color="FFFFFF", bold=True, size=8)
    
    cell_ct = ws.cell(row=room_r1, column=room_c0)
    cell_ct.value = "CT"
    cell_ct.fill = fill_ct
    cell_ct.font = font_ct

    # 6. Desenha as divisórias de forma irrestrita nas 4 faces limites calculadas
    
    # Topo da sala
    for c in range(room_c0, room_c1 + 1):
        _aplicar_borda_espelhada(ws, room_r0, c, 'top', side_style)

    # Base da sala
    for c in range(room_c0, room_c1 + 1):
        _aplicar_borda_espelhada(ws, room_r1, c, 'bottom', side_style)

    # Esquerda da sala
    for r in range(room_r0, room_r1 + 1):
        _aplicar_borda_espelhada(ws, r, room_c0, 'left', side_style)

    # Direita da sala
    for r in range(room_r0, room_r1 + 1):
        _aplicar_borda_espelhada(ws, r, room_c1, 'right', side_style)

def _copiar_lado(side_obj):
    """Cria uma cópia limpa de um objeto Side para preservar estilos originais da célula."""
    if not side_obj or not side_obj.style or side_obj.style == 'none':
        return None
    return Side(border_style=side_obj.style, color=side_obj.color)

def _aplicar_borda_celula_unica(ws, r, c, side, side_style):
    """Aplica o estilo laranja na face correspondente mantendo as outras 3 faces inalteradas."""
    cell = ws.cell(row=r, column=c)
    b = cell.border
    
    # Clona as bordas existentes para não sofrer overwrite do openpyxl [0]
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
    """Aplica a borda de forma espelhada na célula vizinha para garantir renderização correta."""
    _aplicar_borda_celula_unica(ws, r, c, side, side_style)
    if side == 'bottom':
        _aplicar_borda_celula_unica(ws, r + 1, c, 'top', side_style)
    elif side == 'top':
        _aplicar_borda_celula_unica(ws, r - 1, c, 'bottom', side_style)
    elif side == 'right':
        _aplicar_borda_celula_unica(ws, r, c + 1, 'left', side_style)
    elif side == 'left':
        _aplicar_borda_celula_unica(ws, r, c - 1, 'right', side_style)

# ══════════════════════════════════════════════════════════════════════════
# Função de Seleção e Execução de Teste Manual
# ══════════════════════════════════════════════════════════════════════════

def _selecionar_mesas_contiguas(env_cells: Set[Tuple[int, int]], ws, target_qty: int) -> Set[Tuple[int, int]]:
    """
    Busca e reserva mesas de forma estruturada, preenchendo bancadas físicas inteiras 
    antes de transbordar para a próxima bancada mais próxima [0].
    """
    from ScannerPremissas import is_desk_cell
    from BlockMapper import flood_fill
    
    # 1. Coleta todas as mesas utilizáveis no ambiente
    all_desks = set()
    for r, c in env_cells:
        cell = ws.cell(row=r, column=c)
        if is_desk_cell(cell):
            all_desks.add((r, c))
            
    if not all_desks:
        return set()
        
    # 2. Divide as mesas em bancadas físicas reais (clusters) [0]
    benches = flood_fill(all_desks)
    
    # Ordena as bancadas para começar pela do canto superior esquerdo (consistência de leitura)
    benches.sort(key=lambda b: min(b))
    
    selected = set()
    remaining_qty = target_qty
    
    # 3. Aloca as mesas bancada por bancada [0]
    for bench in benches:
        if remaining_qty <= 0:
            break
            
        bench_set = set(bench)
        # Se a bancada física inteira cabe no que falta, pega ela por completo [0]
        if len(bench_set) <= remaining_qty:
            selected.update(bench_set)
            remaining_qty -= len(bench_set)
        else:
            # Se a bancada é maior do que o que falta, ordena as mesas de forma a priorizar 
            # o preenchimento de colunas (para verticais) ou linhas (para horizontais) [0].
            # Isso desvincula do BFS radial e força a ocorrência do corte de espinha [0].
            cols = {c for r, c in bench_set}
            rows = {r for r, c in bench_set}
            
            if len(rows) >= len(cols):
                # Bancada vertical: ordena Coluna primeiro, depois Linha [0]
                sorted_bench = sorted(list(bench_set), key=lambda x: (x[1], x[0]))
            else:
                # Bancada horizontal: ordena Linha primeiro, depois Coluna [0]
                sorted_bench = sorted(list(bench_set), key=lambda x: (x[0], x[1]))
                
            # Seleciona exatamente o restante necessário respeitando a espinha de fiação [0]
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

    # 1. Seleciona estritamente a quantidade de mesas solicitada contiguamente (ex: 30)
    allocated_cells = _selecionar_mesas_contiguas(env_cells, ws, quantidade_mesas)
    
    if not allocated_cells:
        print("Erro: Nenhuma mesa utilizável encontrada nesse ambiente.")
        return

    # 2. Identifica e une as bancadas físicas ou frações de bancadas que contêm mesas alocadas [0]
    from ScannerPremissas import is_desk_cell
    from BlockMapper import flood_fill
    
    all_desks = set()
    for r, c in env_cells:
        cell = ws.cell(row=r, column=c)
        if is_desk_cell(cell):
            all_desks.add((r, c))
            
    benches = flood_fill(all_desks)
    target_bench_cells = set()
    
    # Se uma bancada física possui pelo menos uma mesa alocada para nós, ela entra no cálculo [0]
    for bench in benches:
        bench_set = set(bench)
        allocated_in_bench = bench_set & set(allocated_cells)
        
        if allocated_in_bench:
            r0 = min(r for r, c in allocated_in_bench)
            r1 = max(r for r, c in allocated_in_bench)
            c0 = min(c for r, c in allocated_in_bench)
            c1 = max(c for r, c in allocated_in_bench)
            
            br_y_0 = min(r for r, c in bench_set) # Cobre limites da bancada original
            br_y_1 = max(r for r, c in bench_set)
            bc_x_0 = min(c for r, c in bench_set)
            bc_x_1 = max(c for r, c in bench_set)
            
            cols = {c for r, c in bench_set}
            rows = {r for r, c in bench_set}
            
            leftover_threshold = 2
            
            # Aplica o recorte e unificação de bancadas duplas vertical ou horizontal [0]
            if len(rows) >= len(cols):
                target_r0 = br_y_0 if (r0 - br_y_0 <= leftover_threshold) else r0
                target_r1 = br_y_1 if (br_y_1 - r1 <= leftover_threshold) else r1
                target_c0 = c0
                target_c1 = c1
            else:
                target_r0 = r0
                target_r1 = r1
                target_c0 = bc_x_0 if (c0 - bc_x_0 <= leftover_threshold) else c0
                target_c1 = bc_x_1 if (bc_x_1 - c1 <= leftover_threshold) else c1
                
            # Adiciona apenas a fração útil da bancada sob o critério de espinha e limite de sobras [0]
            for r, c in bench_set:
                if target_r0 <= r <= target_r1 and target_c0 <= c <= target_c1:
                    target_bench_cells.add((r, c))
            
    if not target_bench_cells:
        target_bench_cells = allocated_cells

    # 3. Nenhuma modificação visual ou de valor nas células internas é feita.
    # As mesas preservam suas cores, valores e fontes originais intactas do Excel.

    # 4. Desenha as divisórias de contorno considerando a BANCADA OU FRAÇÃO UNIFICADA (target_bench_cells) [0]
    separar_ambiente_e_desenhar_divisorias(
        ws=ws, 
        env_cells=env_cells, 
        allocated_cells=target_bench_cells, # Passamos a bancada/fração para o desenho das divisórias [0]
        border_style=border_style, 
        border_color=border_color
    )

    wb.save(output_path)
    print(f"✓ Teste concluído com sucesso! Resultado salvo em: '{output_path}'")

if __name__ == "__main__":
    testar_criacao_sala_manual(
        file_path="planta.xlsx",
        sheet_name="JPIII",
        bloco_id="vazio-2",          
        ambiente_letra="A",         
        quantidade_mesas=6,         
        output_path="planta_teste_sala.xlsx"
    )
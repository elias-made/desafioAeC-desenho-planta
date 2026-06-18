import openpyxl
from collections import defaultdict
from typing import Dict, List, Tuple

def manhattan_distance(c1: Tuple[int, int], c2: Tuple[int, int]) -> int:
    return abs(c1[0] - c2[0]) + abs(c1[1] - c2[1])

def is_color_orange(hex_color: str) -> bool:
    if not hex_color or len(hex_color) < 6:
        return False
    hex_color = hex_color[-6:].upper()
    
    # Cores de laranja padrão e variações comuns do Excel
    if hex_color in ['FF9900', 'FFC000', 'ED7D31', 'F79646', 'FFB347', 'FF8C00', 'E26B0A']:
        return True
        
    # Heurística para tons personalizados de laranja (R alto, G médio, B baixo)
    try:
        r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        if r > 200 and 100 <= g <= 200 and b < 100:
            return True
    except:
        pass
    return False

def scan_orange_context(file_path: str = 'planta.xlsx', sheet_name: str = 'JPIII', max_gap: int = 2) -> List[Dict]:
    """
    Identifica as molduras laranjas que delimitam os macro blocos físicos de forma adaptável.
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb[sheet_name]
    
    orange_cells = set()
    cell_values = {}
    
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            
            fill = cell.fill
            if fill and fill.patternType == 'solid' and fill.start_color:
                rgb = str(fill.start_color.rgb)
                if rgb and rgb != '00000000' and is_color_orange(rgb):
                    orange_cells.add((r, c))
                    
            if cell.value is not None:
                val_str = str(cell.value).strip()
                if val_str and len(val_str) > 3 and val_str.upper() not in ('VAZIO', 'SALA', 'CATRACA', 'COWORKING'):
                    cell_values[(r, c)] = val_str

    if not orange_cells:
        print("⚠️ AVISO: Nenhuma moldura ou borda laranja encontrada.")
        return []

    orange_list = list(orange_cells)
    n = len(orange_list)
    parent = list(range(n))
    
    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x
        
    def union(i, j):
        root_i = find(i)
        root_j = find(j)
        if root_i != root_j:
            parent[root_i] = root_j
            
    # Usa a variável max_gap de forma adaptável (evita hardcode de distâncias altas)
    for i in range(n):
        for j in range(i + 1, n):
            if manhattan_distance(orange_list[i], orange_list[j]) <= max_gap:
                union(i, j)
                
    groups = defaultdict(list)
    for i, coord in enumerate(orange_list):
        groups[find(i)].append(coord)
        
    macro_blocks = []
    
    # Filtra ruídos menores e ordena os blocos de cima para baixo na planta
    valid_groups = [coords for coords in groups.values() if len(coords) >= 10]
    valid_groups.sort(key=lambda coords: min(r for r, c in coords))
    
    for idx, coords in enumerate(valid_groups, start=1):
        r_min, r_max = min(r for r, c in coords), max(r for r, c in coords)
        c_min, c_max = min(c for r, c in coords), max(c for r, c in coords)
        
        block_texts = set()
        pad = 3
        for (tr, tc), txt in cell_values.items():
            if (r_min - pad) <= tr <= (r_max + pad) and (c_min - pad) <= tc <= (c_max + pad):
                block_texts.add(txt)
                        
        macro_blocks.append({
            'id': f"Macro_Bloco_{idx}",
            'texts': list(block_texts),
            'bounding_box': (r_min, r_max, c_min, c_max)
        })
        
    return macro_blocks

def build_context_string_for_llm(macro_blocks: List[Dict]) -> str:
    if not macro_blocks:
        return "Nenhum macro-bloco laranja encontrado."
        
    linhas = ["=== PREMISSAS E CONTEXTO EXTRAÍDOS DAS BORDAS LARANJAS ==="]
    for b in macro_blocks:
        linhas.append(f"\n{b['id']}:")
        if b['texts']:
            for txt in b['texts']:
                linhas.append(f"  - {txt}")
        else:
            linhas.append("  - (Sem anotações de texto nesta borda)")
            
    return "\n".join(linhas)
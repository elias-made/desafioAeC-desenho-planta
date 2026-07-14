"""
BlockMapper — funções puras de geometria para plantas de call center.
"""

from collections import deque, defaultdict
from typing import Dict, List, Set, Tuple

Cell = Tuple[int, int]


# ── Primitivas de geometria ─────────────────────────────────────────────

def flood_fill(cells: Set[Cell]) -> List[List[Cell]]:
    """Retorna blocos contíguos (4-vizinhos), ordenados do maior para o menor."""
    visited: Set[Cell] = set()
    blocks: List[List[Cell]] = []
    for seed in sorted(cells):
        if seed in visited:
            continue
        block: List[Cell] = []
        q = deque([seed])
        while q:
            cur = q.popleft()
            if cur in visited or cur not in cells:
                continue
            visited.add(cur)
            block.append(cur)
            r, c = cur
            for dr, dc in ((-1, 0), (1, 0), (0, -1), (0, 1)):
                nb = (r + dr, c + dc)
                if nb not in visited and nb in cells:
                    q.append(nb)
        blocks.append(sorted(block))
    return sorted(blocks, key=len, reverse=True)


def _block_gap(b1: List[Cell], b2: List[Cell]) -> int:
    """Distância mínima de Manhattan entre as bordas de dois blocos."""
    def border(blk):
        s = set(blk)
        return [(r, c) for r, c in blk
                if any((r + dr, c + dc) not in s
                       for dr, dc in ((-1, 0), (1, 0), (0, -1), (0, 1)))]
    best = float('inf')
    for r1, c1 in border(b1):
        for r2, c2 in border(b2):
            d = abs(r1 - r2) + abs(c1 - c2)
            if d < best:
                best = d
            if best <= 1:
                return 1
    return int(best)


def group_zones(blocks: List[List[Cell]], gap: int) -> List[List[Cell]]:
    """Agrupa blocos cujo gap ≤ `gap` em zonas (Union-Find)."""
    n = len(blocks)
    parent = list(range(n))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for i in range(n):
        for j in range(i + 1, n):
            if find(i) != find(j) and _block_gap(blocks[i], blocks[j]) <= gap:
                parent[find(i)] = find(j)

    groups: Dict[int, List[Cell]] = defaultdict(list)
    for i, blk in enumerate(blocks):
        groups[find(i)].extend(blk)

    return sorted(groups.values(), key=len, reverse=True)


# ── Leitura da planta ────────────────────────────────────────────────────

def scan_plant(ws, forbidden_patterns: Set[str]) -> Dict:
    client_cells: Dict[str, Set[Cell]] = {}
    forbidden: Set[Cell] = set()

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                v_str = ""
            elif isinstance(v, float) and v == int(v):
                v_str = str(int(v))
            else:
                v_str = str(v).strip()
            v_up = v_str.upper()

            if any(pat in v_up for pat in forbidden_patterns):
                forbidden.add((r, c))
            else:
                client_cells.setdefault(v_str, set()).add((r, c))

    return {'client_cells': client_cells, 'forbidden': forbidden}

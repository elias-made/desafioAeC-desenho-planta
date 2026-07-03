# AmbienteBuilder.py

import os
import re
from typing import List, Set, Tuple
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell

from ScannerPremissas import scan_orange_context


def _eh_celula_mesclada(ws, r: int, c: int) -> bool:
    """Retorna True se a célula (r, c) faz parte de um range mesclado e não pode ser
    escrita individualmente (openpyxl só permite escrever na âncora/topo-esquerda)."""
    if r < 1 or r > ws.max_row or c < 1 or c > ws.max_column:
        return False
    return isinstance(ws.cell(row=r, column=c), MergedCell)

# ══════════════════════════════════════════════════════════════════════════
# Funções de Fluxo, Contornos e Divisórias Dinâmicas (Sem Hardcode)
# ══════════════════════════════════════════════════════════════════════════

def _eh_celula_de_mesa_local(cell) -> bool:
    """Verifica de forma dinâmica e sem hardcode se a célula representa uma mesa de operador."""
    if cell is None or cell.value is None:
        return False
    from ScannerPremissas import is_desk_cell, is_barrier_cell, cell_has_orange_fill
    
    # Se o scanner base já identificar como mesa, valida imediatamente
    if is_desk_cell(cell):
        return True
        
    val = str(cell.value).strip().upper()
    if val:
        # Se for barreira física ou bloco preenchido de laranja, não é mesa
        if is_barrier_cell(cell) or cell_has_orange_fill(cell):
            return False
        # Evita tratar marcadores de corredores/salas conhecidos como mesas
        if val in ('SALA', 'COWORKING', 'SALA CLIENTE', 'SALA1', 'SALA2', 'SALA3', 'SALA4', 'CATRACA', 'CT', 'ESCANINHOS', '##'):
            return False
        # Evita números altos de anotações (metas como 120, 144) se forem puramente numéricos maiores que 9
        try:
            val_num = float(val)
            if val_num > 9:
                return False
        except ValueError:
            pass
        # Qualquer outro texto (como RENT, ADM, etc.) em célula não-barreira é uma mesa dinâmica
        return True
    return False

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

def _eh_catraca_do_nosso_ambiente(ws, env_cells: Set[Tuple[int, int]], r: int, c: int) -> bool:
    """Verifica se a célula (r, c) é uma catraca que está fisicamente encostada (sem parede laranja) no nosso ambiente base."""
    if (r, c) in env_cells:
        return True
    for dr, dc in ((-1, 0), (1, 0), (0, -1), (0, 1)):
        nr, nc = r + dr, c + dc
        if (nr, nc) in env_cells:
            if not _tem_parede_laranja_entre(ws, r, c, nr, nc):
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
    """Verifica se todas as células no retângulo [r0, r1] x [c0, c1] são corredores livres (sem barreiras ou mesas)."""
    if r0 < 1 or r1 > ws.max_row or c0 < 1 or c1 > ws.max_column:
        return False
    from ScannerPremissas import is_barrier_cell, cell_has_orange_fill
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            cell = ws.cell(row=r, column=c)
            val = str(cell.value).strip().upper() if cell.value is not None else ""
            
            # Bloqueio absoluto: qualquer mesa física (local ou não) não pode ser tratada como corredor de circulação
            if _eh_celula_de_mesa_local(cell) or val.startswith("N_"):
                return False
            
            eh_nossa_catraca = False
            if val in ("CT", "CATRACA"):
                if _eh_catraca_do_nosso_ambiente(ws, env_cells, r, c):
                    eh_nossa_catraca = True
            
            if not eh_nossa_catraca:
                if is_barrier_cell(cell):
                    if val in ('SALA', 'COWORKING', 'SALA CLIENTE', 'SALA1', 'SALA2', 'SALA3', 'SALA4', 'ESCANINHOS'):
                        if (r, c) not in env_cells:
                            return False
                    else:
                        return False
                        
                if cell_has_orange_fill(cell):
                    return False
                    
                if val in ("CT", "CATRACA"):
                    return False
    return True

def get_env_cells(block_id_str: str, env_letter: str, macro_blocks: List[dict]) -> List[Tuple[int, int]]:
    """Identifica as coordenadas de células de forma tolerante a alucinações de letras."""
    if not block_id_str or not env_letter:
        return []
    block_match = re.search(r'\d+', str(block_id_str))
    if not block_match:
        return []
    block_idx = int(block_match.group())
    if block_idx <= len(macro_blocks):
        block = macro_blocks[block_idx - 1]
        block_envs = {e['id'].upper(): e for e in block.get('ambientes', [])}
        
        target_env = env_letter.upper()
        if target_env not in block_envs and len(block_envs) == 1:
            target_env = list(block_envs.keys())[0]
            
        parts = re.split(r'[-_\s+&,|/]+', target_env)
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

# ══════════════════════════════════════════════════════════════════════════
# Geração de Layout de Sala Estruturada com Circulação Interna Própria
# ══════════════════════════════════════════════════════════════════════════

def _calcular_altura_sala(W: int, N: int) -> int:
    desks_per_row = 4 if W == 5 else 2
    h_desks = (N + desks_per_row - 1) // desks_per_row
    h_corridors = (h_desks - 1) // 3 if h_desks > 3 else 0
    return h_desks + h_corridors

def _celulas_contorno_do_ambiente(ws, env_cells: Set[Tuple[int, int]]) -> Set[Tuple[int, int]]:
    """
    Retorna as células do anel de contorno (parede laranja) do ambiente que são piso
    utilizável, e não pilares sólidos nem mesas/rótulos de outros ambientes.

    O scanner (get_interior_cells) remove do env_cells todo o anel de células que carrega
    a borda laranja do contorno, pois elas entram em border_cells. Isso encolhe o ambiente
    em 1 célula em toda a volta e faz com que qualquer sala nova pare a 1 célula da parede
    real. Aqui recuperamos justamente essas células de perímetro (vazias, com borda laranja,
    sem preenchimento sólido) para que a sala possa assentar rente à parede do ambiente.
    """
    from ScannerPremissas import cell_has_orange_fill, cell_has_orange_border
    extra: Set[Tuple[int, int]] = set()
    # Inclui as 8 direções (ortogonais + diagonais). As diagonais são essenciais para
    # recuperar as QUINAS do contorno (ex.: a célula de canto que tem parede laranja em cima
    # e à esquerda, mas que só toca o interior na diagonal). Sem elas, o medidor de gap trata
    # a coluna/linha da parede como inválida e a sala para 1 célula antes da parede real.
    for (r, c) in env_cells:
        for dr, dc in ((-1, 0), (1, 0), (0, -1), (0, 1),
                       (-1, -1), (-1, 1), (1, -1), (1, 1)):
            nr, nc = r + dr, c + dc
            if (nr, nc) in env_cells:
                continue
            if nr < 1 or nr > ws.max_row or nc < 1 or nc > ws.max_column:
                continue
            cell = ws.cell(row=nr, column=nc)
            # Células mescladas (paredes desenhadas como range) não podem ser escritas/estilizadas
            # individualmente pelo openpyxl; não as puxamos para a sala.
            if isinstance(cell, MergedCell):
                continue
            # Pilar/bloco sólido laranja não é piso: nunca ocupar
            if cell_has_orange_fill(cell):
                continue
            # Não engolir mesas nem rótulos (SALA, CT, etc.) de outros ambientes
            val = str(cell.value).strip().upper() if cell.value is not None else ""
            if val:
                continue
            # Só recupera a célula se ela é de fato o contorno laranja do ambiente
            if cell_has_orange_border(cell):
                extra.add((nr, nc))
    return extra

def _eh_celula_valida_para_sala(ws, r: int, c: int, env_cells: Set[Tuple[int, int]]) -> bool:
    """Verifica se a célula está dentro do conjunto real do bloco, evitando pilares
    e outras salas preexistentes, mas permitindo colar nas bordas externas de contorno."""
    if (r, c) not in env_cells:
        return False
        
    if _eh_pilar_ou_coluna(ws, r, c):
        return False
        
    # Permite células com borda laranja do próprio contorno, evitando apenas preenchimentos sólidos (pilares)
    from ScannerPremissas import cell_has_orange_fill
    cell = ws.cell(row=r, column=c)
    if cell_has_orange_fill(cell):
        return False
        
    # Evita textos como SALA, COWORKING, etc. que indiquem outras salas
    val_str = str(cell.value).strip().upper() if cell.value is not None else ""
    if val_str in ('SALA', 'COWORKING', 'SALA CLIENTE', 'SALA1', 'SALA2', 'SALA3', 'SALA4', 'CATRACA', 'CT', 'ESCANINHOS'):
        return False
        
    return True

def _contar_gap_borda(ws, env_cells: Set[Tuple[int, int]], r_start: int, r_end: int, c_start: int, c_end: int, direcao: str) -> int:
    """
    Conta quantas células contíguas e válidas (dentro do ambiente, sem pilares/paredes/outras salas)
    existem a partir da borda do retângulo candidato, na direção especificada, até encontrar
    a primeira célula inválida (barreira real).

    Isso mede a distância física real até a próxima barreira, em vez de usar min()/max() sobre
    todo o ambiente (que confundia corredores distantes e outras alas com a "borda" do bloco).
    Corredores vazios contam como válidos (fazem parte de env_cells), então o gap avança
    naturalmente por eles até topar em algo físico de fato.
    """
    gap = 0
    limite = 3  # não precisamos contar além do limiar usado na validação (>=3 já é aceitável)
    if direcao in ('left', 'right'):
        c = c_start if direcao == 'left' else c_end
        step = -1 if direcao == 'left' else 1
        while gap < limite:
            c += step
            if not all(_eh_celula_valida_para_sala(ws, r, c, env_cells) for r in range(r_start, r_end + 1)):
                break
            gap += 1
    else:
        r = r_start if direcao == 'top' else r_end
        step = -1 if direcao == 'top' else 1
        while gap < limite:
            r += step
            if not all(_eh_celula_valida_para_sala(ws, r, c, env_cells) for c in range(c_start, c_end + 1)):
                break
            gap += 1
    return gap

def _remocao_preserva_conectividade(env_cells: Set[Tuple[int, int]], rect_cells: Set[Tuple[int, int]]) -> bool:
    """
    Verifica se, ao remover o retângulo candidato (rect_cells) do ambiente, o espaço restante
    continua totalmente conectado. Isso evita que a sala seja posicionada de forma a cortar
    o ambiente em dois pedaços, isolando mesas do outro lado sem rota de saída até a catraca.

    Se o restante ficar dividido em mais de um componente conectado, o retângulo é rejeitado,
    mesmo que geometricamente ele fosse válido de outra forma.
    """
    from BlockMapper import flood_fill

    restante = env_cells - rect_cells
    if not restante:
        return True  # Não sobrou nada para verificar (a sala tomou o ambiente inteiro)

    componentes = flood_fill(restante)
    return len(componentes) <= 1

def _celula_eh_corredor_circulacao(ws, env_cells: Set[Tuple[int, int]], r: int, c: int) -> bool:
    """
    True se a célula (r, c) é um corredor de circulação real: está dentro do ambiente,
    é uma faixa livre (sem mesa/barreira/catraca) e NÃO é uma célula de parede-contorno
    (borda ou preenchimento laranja). Isso distingue corredor de circulação de uma parede
    do ambiente, já que o env aumentado inclui o anel de contorno.
    """
    if (r, c) not in env_cells:
        return False
    from ScannerPremissas import cell_has_orange_border, cell_has_orange_fill
    cell = ws.cell(row=r, column=c)
    if cell_has_orange_border(cell) or cell_has_orange_fill(cell):
        return False
    return _eh_faixa_livre(ws, env_cells, r, r, c, c)

def _sala_conecta_corredor(ws, env_cells: Set[Tuple[int, int]], r_start: int, r_end: int, c_start: int, c_end: int) -> bool:
    """
    True se pelo menos um lado do retângulo da sala é adjacente a um corredor de circulação,
    sem parede laranja entre a sala e esse corredor. Garante que a sala tenha uma porta viável
    (mesmo que a porta não seja desenhada nem marcada) ligada à circulação do ambiente.
    """
    # Topo e base
    for c in range(c_start, c_end + 1):
        if _celula_eh_corredor_circulacao(ws, env_cells, r_start - 1, c) and not _tem_parede_laranja_entre(ws, r_start, c, r_start - 1, c):
            return True
        if _celula_eh_corredor_circulacao(ws, env_cells, r_end + 1, c) and not _tem_parede_laranja_entre(ws, r_end, c, r_end + 1, c):
            return True
    # Esquerda e direita
    for r in range(r_start, r_end + 1):
        if _celula_eh_corredor_circulacao(ws, env_cells, r, c_start - 1) and not _tem_parede_laranja_entre(ws, r, c_start, r, c_start - 1):
            return True
        if _celula_eh_corredor_circulacao(ws, env_cells, r, c_end + 1) and not _tem_parede_laranja_entre(ws, r, c_end, r, c_end + 1):
            return True
    return False

def _encontrar_melhor_retangulo_sala(ws, env_cells: Set[Tuple[int, int]], W: int, H: int) -> Tuple[int, int, Set[Tuple[int, int]]]:
    """
    Encontra o melhor retângulo W x H para a sala dentro do bloco,
    priorizando encostar nas paredes externas e minimizar a sobreposição de mesas existentes,
    sem nunca invadir salas preexistentes ou deixar mesas presas nas laterais de forma irregular.
    """
    if not env_cells:
        return None, None, set()
        
    r_max = max(r for r, c in env_cells)
    c_max = max(c for r, c in env_cells)
    
    rows = sorted(list({r for r, c in env_cells}))
    cols = sorted(list({c for r, c in env_cells}))
    
    candidates = []
    
    for r_start in rows:
        if r_start + H - 1 > r_max:
            break
        for c_start in cols:
            if c_start + W - 1 > c_max:
                break
                
            valido = True
            rect_cells = set()
            desks_destroyed = 0
            
            r_end = r_start + H - 1
            c_end = c_start + W - 1
            
            # Impedir recuos estreitos de 1 ou 2 células medindo a distância REAL até a próxima
            # barreira física, célula a célula, a partir de cada lado do retângulo candidato.
            # gap == 0 significa que já encosta na parede/limite (ótimo).
            # gap >= 3 significa que há espaço suficiente para um corredor de circulação (aceitável).
            # 0 < gap < 3 é o caso problemático: sobraria uma faixa inutilizável -> inválido.
            gap_left = _contar_gap_borda(ws, env_cells, r_start, r_end, c_start, c_end, 'left')
            if 0 < gap_left < 3:
                valido = False
            gap_right = _contar_gap_borda(ws, env_cells, r_start, r_end, c_start, c_end, 'right') if valido else 0
            if valido and 0 < gap_right < 3:
                valido = False
            gap_top = _contar_gap_borda(ws, env_cells, r_start, r_end, c_start, c_end, 'top') if valido else 0
            if valido and 0 < gap_top < 3:
                valido = False
            gap_bottom = _contar_gap_borda(ws, env_cells, r_start, r_end, c_start, c_end, 'bottom') if valido else 0
            if valido and 0 < gap_bottom < 3:
                valido = False
                
            if valido:
                for r in range(r_start, r_start + H):
                    for c in range(c_start, c_start + W):
                        if not _eh_celula_valida_para_sala(ws, r, c, env_cells):
                            valido = False
                            break
                        rect_cells.add((r, c))
                        cell = ws.cell(row=r, column=c)
                        if _eh_celula_de_mesa_local(cell):
                            desks_destroyed += 1
                    if not valido:
                        break
            
            # Verifica se retirar este retângulo não isola mesas do restante do ambiente
            # (não pode deixar mesas presas sem rota de saída para a catraca/corredor principal).
            if valido and not _remocao_preserva_conectividade(env_cells, rect_cells):
                valido = False
                        
            if valido:
                touching_sides = 0
                if any(not _eh_celula_valida_para_sala(ws, r_start - 1, c, env_cells) for c in range(c_start, c_start + W)):
                    touching_sides += 1
                if any(not _eh_celula_valida_para_sala(ws, r_start + H, c, env_cells) for c in range(c_start, c_start + W)):
                    touching_sides += 1
                if any(not _eh_celula_valida_para_sala(ws, r, c_start - 1, env_cells) for r in range(r_start, r_start + H)):
                    touching_sides += 1
                if any(not _eh_celula_valida_para_sala(ws, r, c_start + W, env_cells) for r in range(r_start, r_start + H)):
                    touching_sides += 1

                # Preferência principal: a sala deve estar ligada a um corredor de circulação
                # do ambiente (porta única viável), mesmo que a porta não seja desenhada.
                conecta = 1 if _sala_conecta_corredor(ws, env_cells, r_start, r_end, c_start, c_end) else 0

                candidates.append((conecta, touching_sides, desks_destroyed, r_start, c_start, rect_cells))
                
    if candidates:
        # 1º liga em corredor; 2º encosta em mais paredes; 3º destrói menos mesas; depois ordem estável
        candidates.sort(key=lambda x: (-x[0], -x[1], x[2], x[3], x[4]))
        best = candidates[0]
        return best[3], best[4], best[5]
        
    return None, None, set()

def _gerar_bancadas_dinamicas(ws, env_cells: Set[Tuple[int, int]], N: int) -> Set[Tuple[int, int]]:
    """Gera bancadas simples caso falhe a estruturação retangular ideal."""
    usable = {coord for coord in env_cells if not _eh_pilar_ou_coluna(ws, coord[0], coord[1])}
    if len(usable) < N:
        return set(list(usable)[:N])
        
    rows = sorted(list({r for r, c in usable}))
    cols = sorted(list({c for r, c in usable}))
    
    col_pairs = []
    i = 0
    while i < len(cols) - 1:
        if cols[i+1] == cols[i] + 1:
            col_pairs.append((cols[i], cols[i+1]))
            i += 3
        else:
            i += 1
            
    row_pairs = []
    i = 0
    while i < len(rows) - 1:
        if rows[i+1] == rows[i] + 1:
            row_pairs.append((rows[i], rows[i+1]))
            i += 3
        else:
            i += 1
            
    desks = set()
    
    if len(col_pairs) >= len(row_pairs) and col_pairs:
        for c1, c2 in col_pairs:
            if len(desks) >= N:
                break
            common_rows = sorted([r for r in rows if (r, c1) in usable and (r, c2) in usable])
            for r in common_rows:
                if len(desks) >= N:
                    break
                desks.add((r, c1))
                if len(desks) >= N:
                    break
                desks.add((r, c2))
    elif row_pairs:
        for r1, r2 in row_pairs:
            if len(desks) >= N:
                break
            common_cols = sorted([c for c in cols if (r1, c) in usable and (r2, c) in usable])
            for c in common_cols:
                if len(desks) >= N:
                    break
                desks.add((r1, c))
                if len(desks) >= N:
                    break
                desks.add((r2, c))
    else:
        sorted_usable = sorted(list(usable), key=lambda x: (x[1], x[0]))
        desks = set(sorted_usable[:N])
        
    return desks

def _gerar_layout_sala_estruturado(ws, env_cells: Set[Tuple[int, int]], N: int) -> Tuple[Set[Tuple[int, int]], Set[Tuple[int, int]]]:
    """
    Gera um layout de sala estruturado e ADAPTATIVO à quantidade de mesas (N), com
    circulação interna em "U" conectada (um único componente de circulação).

    Regras atendidas:
      - Mesas em bancadas de 2 (back-to-back), com corredor de circulação dos dois lados
        das bancadas (não prende operador contra a parede).
      - Os dois corredores laterais são unidos por uma faixa livre numa das pontas,
        formando uma circulação única (a sala é "pensada em 1 porta"). A porta em si não
        é desenhada nem marcada; a sala pode ser só parede em volta.
      - A sala é dimensionada para caber EXATAMENTE N mesas mais essa circulação.
      - O posicionamento (via _encontrar_melhor_retangulo_sala) prioriza encostar a sala
        num corredor de circulação do ambiente.

    Geometria:
      - Vertical: 4 colunas [corredor | mesa | mesa | corredor] x (ceil(N/2) + 1) linhas,
        onde a última linha é a faixa livre que liga os corredores esquerdo e direito.
      - Horizontal: 4 linhas [corredor / mesa / mesa / corredor] x (ceil(N/2) + 1) colunas,
        onde a última coluna é a faixa livre que liga os corredores superior e inferior.
    """
    # Aumenta o ambiente com o anel de contorno (parede laranja) para que a sala
    # possa assentar rente à parede real, e não parar a 1 célula dela.
    env_cells = set(env_cells) | _celulas_contorno_do_ambiente(ws, env_cells)

    if N <= 0:
        return set(), set()

    linhas_mesas = (N + 1) // 2  # ceil(N/2): nº de linhas (vertical) / colunas (horizontal) de bancada

    # Vertical: 4 colunas x (linhas_mesas + 1 faixa de ligação)
    W_v, H_v = 4, linhas_mesas + 1
    r_v, c_v, _rect_v = _encontrar_melhor_retangulo_sala(ws, env_cells, W_v, H_v)

    # Horizontal: (linhas_mesas + 1 faixa de ligação) colunas x 4 linhas
    H_h, W_h = 4, linhas_mesas + 1
    r_h, c_h, _rect_h = _encontrar_melhor_retangulo_sala(ws, env_cells, W_h, H_h)

    opcoes = []
    if r_v is not None:
        opcoes.append(("v", W_v * H_v, r_v, c_v, W_v, H_v))
    if r_h is not None:
        opcoes.append(("h", W_h * H_h, r_h, c_h, W_h, H_h))

    if opcoes:
        opcoes.sort(key=lambda x: x[1])  # mais compacto primeiro
        layout, _area, r_start, c_start, W, H = opcoes[0]

        room_all_cells = {(r, c)
                          for r in range(r_start, r_start + H)
                          for c in range(c_start, c_start + W)}

        desks_set = set()
        placed = 0
        if layout == "v":
            # Mesas nas 2 colunas centrais; última linha fica livre (faixa de ligação em U)
            desk_cols = (c_start + 1, c_start + 2)
            for r in range(r_start, r_start + H - 1):
                for c in desk_cols:
                    if placed < N:
                        desks_set.add((r, c))
                        placed += 1
        else:
            # Mesas nas 2 linhas centrais; última coluna fica livre (faixa de ligação em U)
            desk_rows = (r_start + 1, r_start + 2)
            for c in range(c_start, c_start + W - 1):
                for r in desk_rows:
                    if placed < N:
                        desks_set.add((r, c))
                        placed += 1

        if placed >= N:
            return desks_set, room_all_cells

    # Fallback: não coube o retângulo estruturado -> bancadas dinâmicas simples
    allocated = _gerar_bancadas_dinamicas(ws, env_cells, N)
    return allocated, allocated

# ══════════════════════════════════════════════════════════════════════════
# Funções de Desenho, Limpeza e Formatação
# ══════════════════════════════════════════════════════════════════════════

def separar_ambiente_e_desenhar_divisorias(
    ws, 
    env_cells: Set[Tuple[int, int]], 
    allocated_cells: Set[Tuple[int, int]], 
    border_style: str = "medium", 
    border_color: str = "FF9900",
    reconstruir_sala: bool = False,
    room_cells_override: Set[Tuple[int, int]] = None  # Parâmetro para controle total da área da sala
):
    """
    Desenha as divisórias ao redor do bloco irregular de mesas ocupadas, adicionando
    corredores ergonômicos funcionais e gerando um contorno 2D célula a célula.
    Garante que as divisórias nunca ultrapassem ou invadam mesas de outros clientes.
    """
    if not allocated_cells or not env_cells:
        return

    from BlockMapper import flood_fill

    side_style = Side(border_style=border_style, color=border_color)
    env_cells = set(env_cells)
    allocated = set(allocated_cells) & env_cells
    if not allocated:
        allocated = set(allocated_cells)

    # Se houver um override estruturado de sala, usamos a grade pré-definida diretamente.
    # A sala é um espaço totalmente controlado (definido por room_cells_override): as mesas
    # NÃO devem ser filtradas pelo env_cells encolhido do scanner, senão as PAs que encostam
    # na parede (recuperadas via anel de contorno) seriam descartadas e a sala ficaria com
    # menos mesas do que o solicitado.
    if room_cells_override is not None:
        room_cells = set(room_cells_override)
        target_bench_cells = set(allocated_cells) & room_cells
        if not target_bench_cells:
            target_bench_cells = set(allocated_cells)
        
        if reconstruir_sala:
            fill_desk = PatternFill(start_color="BDC3C7", end_color="BDC3C7", fill_type="solid")
            font_desk = Font(name="Calibri", size=9)
            for r, c in target_bench_cells:
                if _eh_celula_mesclada(ws, r, c):
                    continue
                cell = ws.cell(row=r, column=c)
                cell.value = "vazio"
                cell.fill = fill_desk
                cell.font = font_desk
    else:
        # 1. Encontra as mesas para as bancadas ocupadas (modo padrão original)
        all_desks_in_env = set()
        for r, c in env_cells:
            cell = ws.cell(row=r, column=c)
            val_str = str(cell.value).strip().upper() if cell.value is not None else ""
            if _eh_celula_de_mesa_local(cell) or (r, c) in allocated or val_str.startswith("N_"):
                all_desks_in_env.add((r, c))

        benches = flood_fill(all_desks_in_env)
        target_bench_cells = set()
        
        for bench in benches:
            bench_set = set(bench)
            allocated_in_bench = bench_set & allocated
            
            if allocated_in_bench:
                b_rows = {r for r, c in bench_set}
                b_cols = {c for r, c in bench_set}
                physical_rows_count = len(b_rows)
                physical_cols_count = len(b_cols)
                
                allocated_rows_count = len({r for r, c in allocated_in_bench})
                allocated_cols_count = len({c for r, c in allocated_in_bench})
                
                corte_colunas_intencional = allocated_cols_count < physical_cols_count
                corte_linhas_intencional = allocated_rows_count < physical_rows_count
                
                total_capacity = len(bench_set)
                allocated_count = len(allocated_in_bench)
                missing_to_complete = total_capacity - allocated_count
                
                excesso_potencial_global = (len(target_bench_cells) + missing_to_complete) - len(allocated)
                max_excesso_permitido = 0 if len(allocated) >= len(allocated_cells) else 2
                
                if (missing_to_complete <= 2 
                        and not corte_colunas_intencional 
                        and not corte_linhas_intencional 
                        and excesso_potencial_global <= max_excesso_permitido):
                    target_bench_cells.update(bench_set)
                else:
                    target_bench_cells.update(allocated_in_bench)

        if not target_bench_cells:
            target_bench_cells = allocated

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
            
            b_rows = {r for r, c in bench_set}
            b_cols = {c for r, c in bench_set}
            is_vertical = len(b_rows) >= len(b_cols)
            
            if is_vertical:
                # Corredor Esquerdo
                for r in range(r0_b, r1_b + 1):
                    if (r, c0_b) in target_bench_cells:
                        target_cell = (r, c0_b - 1)
                        if not _tem_parede_laranja_entre(ws, r, c0_b, r, c0_b - 1):
                            if _eh_faixa_livre(ws, env_cells, r, r, c0_b - 1, c0_b - 1):
                                room_cells.add(target_cell)
                # Corredor Direito
                for r in range(r0_b, r1_b + 1):
                    if (r, c1_b) in target_bench_cells:
                        target_cell = (r, c1_b + 1)
                        if not _tem_parede_laranja_entre(ws, r, c1_b, r, c1_b + 1):
                            if _eh_faixa_livre(ws, env_cells, r, r, c1_b + 1, c1_b + 1):
                                room_cells.add(target_cell)
                # Corredor Inferior
                left_col = c0_b - 1 if (r1_b, c0_b - 1) in room_cells else c0_b
                right_col = c1_b + 1 if (r1_b, c1_b + 1) in room_cells else c1_b
                
                miolo_livre_inferior = True
                for c in range(c0_b, c1_b + 1):
                    if not _eh_faixa_livre(ws, env_cells, r1_b + 1, r1_b + 1, c, c) or _tem_parede_laranja_entre(ws, r1_b, c, r1_b + 1, c):
                        miolo_livre_inferior = False
                        break
                        
                if miolo_livre_inferior:
                    for c in range(left_col, right_col + 1):
                        target_cell = (r1_b + 1, c)
                        if not _tem_parede_laranja_entre(ws, r1_b, c, r1_b + 1, c):
                            if _eh_faixa_livre(ws, env_cells, r1_b + 1, r1_b + 1, c, c):
                                room_cells.add(target_cell)
            else:
                # Corredor Superior
                for c in range(c0_b, c1_b + 1):
                    if (r0_b, c) in target_bench_cells:
                        target_cell = (r0_b - 1, c)
                        if not _tem_parede_laranja_entre(ws, r0_b, c, r0_b - 1, c):
                            if _eh_faixa_livre(ws, env_cells, r0_b - 1, r0_b - 1, c, c):
                                room_cells.add(target_cell)
                # Corredor Inferior
                for c in range(c0_b, c1_b + 1):
                    if (r1_b, c) in target_bench_cells:
                        target_cell = (r1_b + 1, c)
                        if not _tem_parede_laranja_entre(ws, r1_b, c, r1_b + 1, c):
                            if _eh_faixa_livre(ws, env_cells, r1_b + 1, r1_b + 1, c, c):
                                room_cells.add(target_cell)
                # Corredor Esquerdo
                top_row = r0_b - 1 if (r0_b - 1, c0_b) in room_cells else r0_b
                bottom_row = r1_b + 1 if (r1_b + 1, c0_b) in room_cells else r1_b
                
                miolo_livre_esquerdo = True
                for r in range(r0_b, r1_b + 1):
                    if not _eh_faixa_livre(ws, env_cells, r, r, c0_b - 1, c0_b - 1) or _tem_parede_laranja_entre(ws, r, c0_b, r, c0_b - 1):
                        miolo_livre_esquerdo = False
                        break
                        
                if miolo_livre_esquerdo:
                    for r in range(top_row, bottom_row + 1):
                        target_cell = (r, c0_b - 1)
                        if not _tem_parede_laranja_entre(ws, r, c0_b, r, c0_b - 1):
                            if _eh_faixa_livre(ws, env_cells, r, r, c0_b - 1, c0_b - 1):
                                room_cells.add(target_cell)
                                
                # Corredor Direito
                top_row = r0_b - 1 if (r0_b - 1, c1_b) in room_cells else r0_b
                bottom_row = r1_b + 1 if (r1_b + 1, c1_b) in room_cells else r1_b
                
                miolo_livre_direito = True
                for r in range(r0_b, r1_b + 1):
                    if not _eh_faixa_livre(ws, env_cells, r, r, c1_b + 1, c1_b + 1) or _tem_parede_laranja_entre(ws, r, c1_b, r, c1_b + 1):
                        miolo_livre_direito = False
                        break
                        
                    if miolo_livre_direito:
                        for r in range(top_row, bottom_row + 1):
                            target_cell = (r, c1_b + 1)
                            if not _tem_parede_laranja_entre(ws, r, c1_b, r, c1_b + 1):
                                if _eh_faixa_livre(ws, env_cells, r, r, c1_b + 1, c1_b + 1):
                                    room_cells.add(target_cell)

        # 3.5. PONTE DE CORREDOR VERTICAL
        corridor_columns = {c for r, c in (room_cells - target_bench_cells)}
        for col in sorted(list(corridor_columns)):
            rows_in_col = sorted(list({r for r, c in room_cells if c == col}))
            if len(rows_in_col) > 1:
                for i in range(len(rows_in_col) - 1):
                    r_start_idx = rows_in_col[i]
                    r_end_idx = rows_in_col[i + 1]
                    if r_end_idx - r_start_idx > 1:
                        gap_rows = range(r_start_idx + 1, r_end_idx)
                        all_gap_free = True
                        for gr in gap_rows:
                            target_cell = (gr, col)
                            if _tem_parede_laranja_entre(ws, gr - 1, col, gr, col) or not _eh_faixa_livre(ws, env_cells, gr, gr, col, col):
                                all_gap_free = False
                                break
                        if all_gap_free:
                            for gr in gap_rows:
                                room_cells.add((gr, col))

    # 4. INSERÇÃO DA CATRACA (CT) ÚNICA
    corridor_cells = room_cells - target_bench_cells
    cell_ct_coord = None
    if corridor_cells and not reconstruir_sala:  # Se for sala fechada interna (reconstruir_sala=True), omite a catraca (CT)
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
            
        # Ignora corredores em células mescladas ao escolher onde vai a catraca
        bottom_corridors = [cell for cell in bottom_corridors if not _eh_celula_mesclada(ws, cell[0], cell[1])]
        if bottom_corridors:
            cell_ct_coord = bottom_corridors[0]
            cell_ct = ws.cell(row=cell_ct_coord[0], column=cell_ct_coord[1])

            fill_ct = PatternFill(start_color="FF9900", end_color="FF9900", fill_type="solid")
            font_ct = Font(color="FFFFFF", bold=True, size=8)
            cell_ct.value = "CT"
            cell_ct.fill = fill_ct
            cell_ct.font = font_ct

    # 4.5. Limpa os corredores internos da sala criada
    if reconstruir_sala:
        for r, c in room_cells:
            if (r, c) not in target_bench_cells and (r, c) != cell_ct_coord:
                if not _eh_pilar_ou_coluna(ws, r, c) and not _eh_celula_mesclada(ws, r, c):
                    cell = ws.cell(row=r, column=c)
                    cell.value = ""
                    cell.fill = PatternFill(fill_type=None)

    # 5. DESENHO DAS DIVISÓRIAS
    for r, c in room_cells:
        if (r - 1, c) not in room_cells:
            _aplicar_borda_espelhada(ws, r, c, 'top', side_style)
        if (r + 1, c) not in room_cells:
            _aplicar_borda_espelhada(ws, r, c, 'bottom', side_style)
        if (r, c - 1) not in room_cells:
            _aplicar_borda_espelhada(ws, r, c, 'left', side_style)
        if (r, c + 1) not in room_cells:
            _aplicar_borda_espelhada(ws, r, c, 'right', side_style)

def _copiar_lado(side_obj):
    if not side_obj or not side_obj.style or side_obj.style == 'none':
        return None
    return Side(border_style=side_obj.style, color=side_obj.color)

def _aplicar_borda_celula_unica(ws, r, c, side, side_style):
    if _eh_celula_mesclada(ws, r, c):
        return
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
    """Aplica a borda de forma espelhada e dinâmica validando os limites da planilha."""
    _aplicar_borda_celula_unica(ws, r, c, side, side_style)
    if side == 'bottom' and r + 1 <= ws.max_row:
        _aplicar_borda_celula_unica(ws, r + 1, c, 'top', side_style)
    elif side == 'top' and r - 1 >= 1:
        _aplicar_borda_celula_unica(ws, r - 1, c, 'bottom', side_style)
    elif side == 'right' and c + 1 <= ws.max_column:
        _aplicar_borda_celula_unica(ws, r, c + 1, 'left', side_style)
    elif side == 'left' and c - 1 >= 1:
        _aplicar_borda_celula_unica(ws, r, c - 1, 'right', side_style)

# ══════════════════════════════════════════════════════════════════════════
# Função de Seleção e Execução de Teste Manual
# ══════════════════════════════════════════════════════════════════════════

def _selecionar_mesas_contiguas(env_cells: Set[Tuple[int, int]], ws, target_qty: int) -> Set[Tuple[int, int]]:
    """
    Busca e reserva mesas de forma estruturada, agrupando as bancadas em Macro-Clusters espaciais.
    Mantém a ordem exata de preenchimento para garantir que truncamentos de metas ímpares
    ocorram estritamente na ponta da última bancada em uso, preservando os blocos anteriores.
    """
    all_desks = set()
    for r, c in env_cells:
        cell = ws.cell(row=r, column=c)
        if _eh_celula_de_mesa_local(cell):
            all_desks.add((r, c))
            
    if not all_desks:
        return set()
        
    from BlockMapper import flood_fill
    benches = flood_fill(all_desks)
    n_benches = len(benches)
    if n_benches == 0:
        return set()
        
    parent = list(range(n_benches))
    
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
            
    for i in range(n_benches):
        for j in range(i + 1, n_benches):
            min_dist = min(
                abs(r1 - r2) + abs(c1 - c2)
                for r1, c1 in benches[i]
                for r2, c2 in benches[j]
            )
            if min_dist <= 4:
                union(i, j)
                
    from collections import defaultdict
    clusters_map = defaultdict(list)
    for i in range(n_benches):
        clusters_map[find(i)].append(benches[i])
        
    clusters = list(clusters_map.values())
    
    adjusted_target_qty = target_qty
    possui_bancada_dupla = False
    for cl in clusters:
        for b in cl:
            b_cols = {c for r, c in b}
            b_rows = {r for r, c in b}
            if len(b_cols) == 2 or len(b_rows) == 2:
                possui_bancada_dupla = True
                break
        if possui_bancada_dupla:
            break
            
    if possui_bancada_dupla and adjusted_target_qty % 2 != 0:
        adjusted_target_qty += 1

    valid_clusters = []
    for cl in clusters:
        cap = sum(len(b) for b in cl)
        if cap >= adjusted_target_qty:
            valid_clusters.append((cl, cap))
            
    if valid_clusters:
        valid_clusters.sort(key=lambda x: x[1])
        selected_benches = valid_clusters[0][0]
    else:
        clusters_with_cap = []
        for cl in clusters:
            cap = sum(len(b) for b in cl)
            clusters_with_cap.append((cl, cap))
        clusters_with_cap.sort(key=lambda x: -x[1])
        
        selected_benches = []
        for cl, cap in clusters_with_cap:
            selected_benches.extend(cl)
            
    selected_set = set()
    selected_list = []
    
    max_bench_cap = max(len(b) for b in selected_benches) if selected_benches else 0
    
    if adjusted_target_qty > max_bench_cap:
        selected_benches_sorted = sorted(selected_benches, key=lambda b: (min(r for r, c in b), min(c for r, c in b)))
        
        remaining_qty = adjusted_target_qty
        for b in selected_benches_sorted:
            if remaining_qty <= 0:
                break
                
            b_cols = {c for r, c in b}
            b_rows = {r for r, c in b}
            if len(b_rows) >= len(b_cols):  # Vertical
                sorted_cells = sorted(list(b), key=lambda x: (x[0], x[1]))
            else:  # Horizontal
                sorted_cells = sorted(list(b), key=lambda x: (x[1], x[0]))
                
            take_qty = min(len(sorted_cells), remaining_qty)
            for cell in sorted_cells[:take_qty]:
                if cell not in selected_set:
                    selected_set.add(cell)
                    selected_list.append(cell)
            remaining_qty -= take_qty
            
    else:
        vertical_benches = []
        horizontal_benches = []
        for b in selected_benches:
            b_rows = {r for r, c in b}
            b_cols = {c for r, c in b}
            if len(b_rows) >= len(b_cols):
                vertical_benches.append(b)
            else:
                horizontal_benches.append(b)
                
        is_mostly_vertical = len(vertical_benches) >= len(horizontal_benches)
        
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
            
            valid_groups = [grp for grp in groups if sum(len(b) for b in grp) >= adjusted_target_qty]
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
                    
                    if cap >= adjusted_target_qty:
                        excess = cap - adjusted_target_qty
                        corridor_height = max_rows - H
                        candidates.append((cap, N, H, excess, corridor_height))
            
            if candidates:
                candidates.sort(key=lambda x: (x[3], x[1], -x[4]))
                best_N = candidates[0][1]
                best_H = candidates[0][2]
            else:
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
                remaining_qty = adjusted_target_qty
                for b in benches_to_use[:best_N]:
                    b_cols = {c for r, c in b}
                    physical_cols_count = len(b_cols)
                    
                    if physical_cols_count == 2 and remaining_qty % 2 != 0:
                        remaining_qty += 1
                        
                    b_rows_sorted = sorted(list({r for r, c in b}))
                    allowed_rows = set(b_rows_sorted[:best_H])
                    allowed_cells = [cell for cell in b if cell[0] in allowed_rows]
                    allowed_cells.sort(key=lambda x: (x[0], x[1]))
                    
                    take_qty = min(len(allowed_cells), remaining_qty)
                    for cell in allowed_cells[:take_qty]:
                        if cell not in selected_set:
                            selected_set.add(cell)
                            selected_list.append(cell)
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
                    if abs(min(c for grp_b in grp for r, c in grp_b) - bc_x_0) <= 2:
                        grp.append(b)
                        placed = True
                        break
                if not placed:
                    groups.append([b])
            
            valid_groups = [grp for grp in groups if sum(len(b) for b in grp) >= adjusted_target_qty]
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
                    
                    if cap >= adjusted_target_qty:
                        excess = cap - adjusted_target_qty
                        corridor_width = max_cols - W
                        candidates.append((cap, N, W, excess, corridor_width))
                        
            if candidates:
                candidates.sort(key=lambda x: (x[3], x[1], -x[4]))
                best_N = candidates[0][1]
                best_W = candidates[0][2]
            else:
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
                remaining_qty = adjusted_target_qty
                for b in benches_to_use[:best_N]:
                    b_rows = {r for r, c in b}
                    physical_rows_count = len(b_rows)
                    
                    if physical_rows_count == 2 and remaining_qty % 2 != 0:
                        remaining_qty += 1
                        
                    b_cols_sorted = sorted(list({c for r, c in b}))
                    allowed_cols = set(b_cols_sorted[:best_W])
                    allocated_cells = [cell for cell in b if cell[1] in allowed_cols]
                    allocated_cells.sort(key=lambda x: (x[1], x[0]))
                    
                    take_qty = min(len(allocated_cells), remaining_qty)
                    for cell in allocated_cells[:take_qty]:
                        if cell not in selected_set:
                            selected_set.add(cell)
                            selected_list.append(cell)
                    remaining_qty -= take_qty
                    if remaining_qty <= 0:
                        break

    if not selected_list:
        benches.sort(key=lambda b: min(b))
        remaining_qty = adjusted_target_qty
        for bench in benches:
            if remaining_qty <= 0:
                break
            bench_set = set(bench)
            cols = {c for r, c in bench_set}
            rows = {r for r, c in bench_set}
            if len(rows) >= len(cols):
                sorted_bench = sorted(list(bench_set), key=lambda x: (x[0], x[1]))
            else:
                sorted_bench = sorted(list(bench_set), key=lambda x: (x[1], x[0]))
                
            if len(bench_set) <= remaining_qty:
                for cell in sorted_bench:
                    if cell not in selected_set:
                        selected_set.add(cell)
                        selected_list.append(cell)
                remaining_qty -= len(bench_set)
            else:
                for cell in sorted_bench[:remaining_qty]:
                    if cell not in selected_set:
                        selected_set.add(cell)
                        selected_list.append(cell)
                remaining_qty = 0
                break

    if len(selected_list) > target_qty:
        selected_list = selected_list[:target_qty]
                
    return set(selected_list)

def testar_criacao_sala_manual(
    file_path: str,
    sheet_name: str,
    bloco_id: str,
    ambiente_letra: str,
    quantidade_mesas: int,
    quantidade_mesas_sala: int = None,  # Opcional para ativação da sala interna
    output_path: str = "planta_teste_sala.xlsx",
    border_style: str = "medium",
    border_color: str = "FF9900"
):
    """
    Carrega a planilha, seleciona a equipe principal de forma tradicional desenhando o
    perímetro original do ambiente, e depois cria a sala interna dentro dele sem catraca.
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

    allocated_sala = set()
    room_cells_override = None
    
    # 1. Se solicitada a sala fechada, calcula sua área e células primeiro
    if quantidade_mesas_sala is not None:
        print(f"🛠️ Planejando Sala Fechada Estruturada de {quantidade_mesas_sala} mesas dentro do ambiente...")
        allocated_sala, room_cells_override = _gerar_layout_sala_estruturado(ws, env_cells, quantidade_mesas_sala)

    # 2. Aloca a equipe principal estritamente no espaço restante (env_cells - room_cells_override)
    # Isso implementa a compensação: se a sala consumiu mesas existentes, a equipe principal se expandirá no espaço livre remanescente.
    available_env_cells = env_cells - (room_cells_override if room_cells_override else set())
    allocated_ambiente = _selecionar_mesas_contiguas(available_env_cells, ws, quantidade_mesas)
    
    if not allocated_ambiente:
        print("Erro: Nenhuma mesa utilizável encontrada nesse ambiente para a equipe principal.")
        return

    # 3. Desenha as divisórias do contorno externo do Ambiente principal considerando a união de ambos (evita dentes na parede)
    allocated_total = allocated_ambiente | (room_cells_override if room_cells_override else set())
    separar_ambiente_e_desenhar_divisorias(
        ws=ws, 
        env_cells=env_cells, 
        allocated_cells=allocated_total,
        border_style=border_style, 
        border_color=border_color,
        reconstruir_sala=False
    )

    # 4. Desenha as divisórias internas da sala fechada e reconstrói as mesas e corredores apenas nessa área isolada
    if quantidade_mesas_sala is not None:
        separar_ambiente_e_desenhar_divisorias(
            ws=ws, 
            env_cells=env_cells, 
            allocated_cells=allocated_sala,
            border_style=border_style, 
            border_color=border_color,
            reconstruir_sala=True,
            room_cells_override=room_cells_override
        )

    wb.save(output_path)
    print(f"✓ Teste concluído com sucesso! Resultado salvo em: '{output_path}'")

if __name__ == "__main__":
    testar_criacao_sala_manual(
        file_path="planta.xlsx",
        sheet_name="JPIII",
        bloco_id="Bloco_2",          # Note que o bloco_id no scan original de JPIII para a planta real é Bloco_2 (com o caractere underscore)
        ambiente_letra="A",         
        quantidade_mesas=120, 
        quantidade_mesas_sala=10,  # Exemplo de criação de sala dinâmica para 10 mesas com corredor
        output_path="planta_teste_sala.xlsx"
    )
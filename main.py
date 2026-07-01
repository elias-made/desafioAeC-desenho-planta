# main.py

import asyncio
import os
import shutil
import openpyxl
import traceback
from copy import copy
from dotenv import load_dotenv
from openpyxl.styles import PatternFill, Font

# Importando dependências do LayoutEngine
from LayoutEngine import (
    SHEET_NAME, FORBIDDEN_PATTERNS,
    Acao, PropostaMock, clean_json_string, extrair_e_carregar_json,
    load_plant, clone_ws, build_plant_info, build_blocos_info,
    execute_alocacao, write_report, salvar_auditoria,
    normalizar_acoes, validar_inventario
)

# Importando dependências do AmbienteBuilder
from AmbienteBuilder import separar_ambiente_e_desenhar_divisorias, _selecionar_mesas_contiguas, get_env_cells

# Importando dependências e agentes do Agents.py (Fluxo de 2 Agentes)
from Agents import PosicionadorDeps, OrganizadorDeps, posicionador, organizador
from ScannerPremissas import scan_orange_context, normalize_val
from BlockMapper import scan_plant

load_dotenv()

# ══════════════════════════════════════════════════════════════════════════
# Fluxo de Execução Principal (Unificado de 2 Etapas com Auto-Correção)
# ══════════════════════════════════════════════════════════════════════════

async def main():
    # 1. Carrega as premissas brutas do arquivo de texto
    with open('premissas.txt', encoding='utf-8') as f: 
        premissas_txt = f.read().strip()

    # Diretrizes de nomenclatura inseridas diretamente via prompt
    regras_nomes_sistema = (
        "\n=== DIRETRIZES DE NOMENCLATURA SISTÊMICA (OBRIGATÓRIO) ===\n"
        "Ao criar novas operações sugeridas nas premissas, você DEVE nomeá-las de forma sequencial:\n"
        "  - A primeira operação criada deve ser nomeada rigorosamente como: 'N_1'\n"
        "  - A segunda operação criada deve ser nomeada rigorosamente como: 'N_2'\n"
        "  - A terceira operação criada deve ser nomeada rigorosamente como: 'N_3'\n"
        "E assim por diante. Use sempre este padrão e mantenha a consistência nominal.\n"
    )

    # Extrai o contexto visual das bordas laranjas (necessário para o mapeamento de células permitidas)
    dados_laranjas = scan_orange_context('planta.xlsx', SHEET_NAME)
    
    # Consolida as diretrizes do sistema com o texto de entrada (as anotações visuais vão direto no bloco)
    premissas_completas = f"{premissas_txt}\n{regras_nomes_sistema}"

    # Carrega a planta original e escaneia os dados iniciais
    wb, ws = load_plant()
    
    # --- SNAPSHOT DA PLANILHA ORIGINAL (Restauração Perfeita de Inventário) ---
    original_snapshot = {}
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            original_snapshot[(r, c)] = {
                'value': cell.value,
                'fill': copy(cell.fill) if cell.has_style else None,
                'font': copy(cell.font) if cell.has_style else None
            }

    plant_data = scan_plant(ws, FORBIDDEN_PATTERNS)
    plant_info_str = build_plant_info(plant_data)
    blocos_info_str = build_blocos_info(plant_data, ws.max_row, ws.max_column, ws, file_path='planta.xlsx')
    
    # ── PROTEÇÃO CONTRA USO DE CORREDORES ──
    allowed_cells = set()
    boundary_cells = set()
    for block in dados_laranjas:
        for env in block.get('ambientes', []):
            boundary_cells.update(env['cells'])
            
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            if (r, c) not in boundary_cells:
                continue  
                
            cell = ws.cell(row=r, column=c)
            v = cell.value
            if v is not None:
                v_str = str(v).strip()
                if v_str != "" and v_str.upper() not in ('CT', 'CATRACA', 'SA', 'CW', 'COWORKING', '##'):
                    allowed_cells.add((r, c))
    # ─────────────────────────────────────────────────────────────────────────────

    # ══════════════════════════════════════════════════════════════════════════
    # ETAPA 1: Execução Direta do Posicionador (Sem Loops Desnecessários)
    # ══════════════════════════════════════════════════════════════════════════
    print("1. Posicionador gerando ações primárias (alocação bruta)...")
    
    dest_file = "propostas/proposta_final.xlsx"
    safe_name = "final"
    os.makedirs('propostas', exist_ok=True)
    shutil.copy("planta.xlsx", dest_file)
    new_wb = openpyxl.load_workbook(dest_file)
    new_ws = new_wb[SHEET_NAME]

    pos_deps = PosicionadorDeps(
        plant_info=plant_info_str,
        blocos_info=blocos_info_str,
        premissas=premissas_completas
    )
    
    # Execução única e direta do Posicionador para gerar o rascunho inicial
    res_posicionador = await posicionador.run(
        "Gere a proposta de alocação inicial baseada nas premissas de negócio.", 
        deps=pos_deps
    )
    rascunho_layout_json = clean_json_string(res_posicionador.output)
    pos_data = extrair_e_carregar_json(rascunho_layout_json)
    
    # Executa a alocação bruta na planilha master
    acoes_totais = []
    for ac in pos_data.get("acoes_primarias", []):
        acoes_totais.append(Acao(**ac))
        
    gabarito_dados = pos_data.get("gabarito", {})
    
    # FALLBACK: Se o bloco "gabarito" for omitido, reconstrói a partir das ações planejadas
    if not gabarito_dados and "acoes_primarias" in pos_data:
        reducoes_fallback = {}
        novos_fallback = []
        for ac in pos_data.get("acoes_primarias", []):
            tipo = ac.get("tipo", "").strip().lower()
            cli = ac.get("cliente", "").strip()
            qtd = ac.get("quantidade", 0)
            if tipo == "liberar":
                reducoes_fallback[cli] = reducoes_fallback.get(cli, 0) + qtd
            elif tipo in ("realocar", "posicionar") and cli.upper().startswith("N_"):
                existente = next((nc for nc in novos_fallback if nc["nome"] == cli.upper()), None)
                if existente:
                    existente["PAs"] += qtd
                else:
                    novos_fallback.append({"nome": cli.upper(), "PAs": qtd})
        gabarito_dados = {"reducoes": reducoes_fallback, "novos_clientes": novos_fallback}
        
    novos_clientes_brutos = gabarito_dados.get("novos_clientes", [])
    novos_clientes_extraidos = []
    for nc in novos_clientes_brutos:
        if isinstance(nc, dict):
            nome_orig = nc.get("nome", "")
            pas = nc.get("PAs", 0)
        else:
            nome_orig = str(nc)
            pas = 0
        nome_norm = nome_orig.upper().strip().replace(" ", "_")
        novos_clientes_extraidos.append({
            "nome": nome_norm,
            "PAs": pas
        })
        
    parametros_premissas = {
        "reducoes": {normalize_val(k): v for k, v in gabarito_dados.get("reducoes", {}).items()},
        "novos_clientes": novos_clientes_extraidos
    }
        
    # Normaliza as ações primárias antes de executar
    acoes_totais = normalizar_acoes(acoes_totais, parametros_premissas["novos_clientes"])
        
    proposta_inicial = PropostaMock(
        nome=pos_data.get("nome", "Alocacao Bruta"),
        custo_obras="Baixo",
        acoes=acoes_totais
    )
    
    # Executa o rascunho de posicionamento bruto de soma-zero na planilha master
    log_acumulado, _ = execute_alocacao(new_ws, proposta_inicial, plant_data, allowed_cells, file_path='planta.xlsx')
    criar_ambientes_solicitados = pos_data.get("criar_ambientes", [])
    erros_validacao = validar_inventario(ws, new_ws, allowed_cells, parametros_premissas)

    # === INTEGRAÇÃO DO AMBIENTEBUILDER ===
    ambientes_criados_info = []
    if criar_ambientes_solicitados:
        print("🛠️ Iniciando AmbienteBuilder para desenhar divisórias físicas...")
        macro_blocks_atual = scan_orange_context(dest_file, SHEET_NAME)
        for amb in criar_ambientes_solicitados:
            bloco_id = amb.get("bloco")
            ambiente_letra = amb.get("ambiente")
            qtd_mesas = amb.get("quantidade_mesas")
            cliente_dest = amb.get("cliente_destinado")
            
            env_cells = set(get_env_cells(bloco_id, ambiente_letra, macro_blocks_atual))
            if env_cells:
                allocated_cells = _selecionar_mesas_contiguas(env_cells, new_ws, qtd_mesas)
                if allocated_cells:
                    # Desenha as divisórias físicas da sala
                    separar_ambiente_e_desenhar_divisorias(new_ws, env_cells, allocated_cells)
                    
                    # --- SINCRONIZAÇÃO DE CORES E VALORES (EVITA VAZAMENTO DE MESAS) ---
                    # A. Recupera de forma dinâmica o estilo visual aplicado ao novo cliente
                    fill_to_apply = None
                    font_to_apply = None
                    for r, c in env_cells:
                        cell = new_ws.cell(row=r, column=c)
                        if cell.value == cliente_dest:
                            fill_to_apply = copy(cell.fill) if cell.has_style else None
                            font_to_apply = copy(cell.font) if cell.has_style else None
                            break
                    if not fill_to_apply:
                        fill_to_apply = PatternFill(start_color='BDC3C7', end_color='BDC3C7', fill_type='solid')
                        font_to_apply = Font(color='FFFFFF', bold=True, size=8)

                    # B. Alinha as posições de mesas internas e restaura as externas ao estado original
                    for r, c in env_cells:
                        cell = new_ws.cell(row=r, column=c)
                        if cell.value == "CT":
                            continue
                            
                        if (r, c) in allocated_cells:
                            cell.value = cliente_dest
                            cell.fill = fill_to_apply
                            cell.font = font_to_apply
                        else:
                            snap = original_snapshot.get((r, c))
                            if snap:
                                cell.value = snap['value']
                                if snap['fill']: cell.fill = snap['fill']
                                if snap['font']: cell.font = snap['font']

                    ambientes_criados_info.append(
                        f"- Criado Novo Ambiente Fechado no Bloco '{bloco_id}', Ambiente '{ambiente_letra}' "
                        f"com {len(allocated_cells)} PAs para o cliente '{cliente_dest}'."
                    )
                    print(f"   ✓ Divisórias físicas desenhadas no bloco {bloco_id}-{ambiente_letra}!")
                else:
                    print(f"   ⚠️ Nenhuma mesa contígua encontrada para criar sala no {bloco_id}-{ambiente_letra}.")
            else:
                print(f"   ⚠️ Bloco/Ambiente '{bloco_id}-{ambiente_letra}' não localizado para o AmbienteBuilder.")

    # Salva a planilha intermediária (para garantir as alterações físicas das paredes)
    new_wb.save(dest_file)

    # Re-scan dos blocos e paredes laranjas (ScannerPremissas) para forçar o remapeamento dinâmico
    import ScannerPremissas
    ScannerPremissas._orange_context_cache = None  # Limpa o cache estático do scanner
    
    ambientes_criados_str = "\n".join(ambientes_criados_info) if ambientes_criados_info else "Nenhum ambiente físico criado nesta rodada."

    # Salva auditoria do posicionador
    salvar_auditoria("1_posicionador", (
        f"[PREMISSAS DO TXT]\n{pos_deps.premissas}\n\n"
        f"[BLOCOS INFO]\n{pos_deps.blocos_info}\n\n"
        f"[PLANT INFO]\n{pos_deps.plant_info}\n\n"
        f"[ACOES PLANEJADAS]\n{rascunho_layout_json}"
    ), res_posicionador.output, safe_name)

    # ══════════════════════════════════════════════════════════════════════════
    # ETAPA 2: Loop de Auto-Correção e Swaps via Organizador
    # ══════════════════════════════════════════════════════════════════════════
    print("2. Organizador iniciando loop de alinhamento e swaps...")

    for iteracao in range(1, 4):
        print(f"   -> Tentativa do Organizador {iteracao}/3...")
        
        try:
            # O scan de blocos físicos é refeito de forma dinâmica na planilha modificada pelo AmbienteBuilder
            current_wb = openpyxl.load_workbook(dest_file)
            current_ws = current_wb[SHEET_NAME]
            
            current_plant_data = scan_plant(current_ws, FORBIDDEN_PATTERNS)
            current_plant_info = build_plant_info(current_plant_data)
            
            # Chamada passando o caminho dinâmico do arquivo modificado
            current_blocos_info = build_blocos_info(
                current_plant_data, 
                current_ws.max_row, 
                current_ws.max_column, 
                current_ws, 
                file_path=dest_file
            )
        except Exception as e:
            print(f"\n❌ ERRO DETECTADO NA PREPARAÇÃO DA TENTATIVA {iteracao}:")
            traceback.print_exc()
            return
        
        org_deps = OrganizadorDeps(
            plant_info=current_plant_info,
            blocos_info=current_blocos_info,
            premissas=premissas_completas,
            rascunho_layout=rascunho_layout_json,
            ambientes_criados=ambientes_criados_str
        )
        
        feedback_erros = ""
        if erros_validacao:
            feedback_erros = (
                "\n\n⚠️ INCONSISTÊNCIAS DE SOMA OPERACIONAL NO LAYOUT ATUAL:\n"
                + "\n".join(f"- {err}" for err in erros_validacao)
                + "\n\nAplique ações de 'transferir' (permutas/swaps) para reequilibrar as quantidades exatas de volta!"
            )
        
        res_organizador = await organizador.run(
            f"Avalie a planta física atual. Se houver qualquer violação de premissas (como unificação, exclusividade, novo cliente desalinhado de sua sala física), utilize a função 'transferir' para corrigir. Se estiver tudo certo, retorne a lista vazia.{feedback_erros}",
            deps=org_deps
        )
        reorganizacao_json = clean_json_string(res_organizador.output)
        
        salvar_auditoria(f"2_organizador_iter{iteracao}", (
            f"[RASCUNHO DO LAYOUT]\n{org_deps.rascunho_layout}\n\n"
            f"[AMBIENTES RECENTEMENTE CRIADOS]\n{org_deps.ambientes_criados}\n\n"
            f"[PREMISSAS DO TXT]\n{org_deps.premissas}\n\n"
            f"[BLOCOS INFO ATUALIZADOS NA PLANILHA FISICA]\n{org_deps.blocos_info}\n\n"
            f"[PLANT INFO ATUALIZADA]\n{org_deps.plant_info}"
        ), res_organizador.output, safe_name)
        
        org_data = extrair_e_carregar_json(reorganizacao_json)
        acoes_org = org_data.get("acoes_organizacao", [])
        
        if not acoes_org and not erros_validacao:
            print(f"✓ Sistema auditado e validado. Todas as premissas e inventários estão plenamente satisfatórios!")
            break
            
        print(f"   ⚠ Validando ações corretivas em sandbox...")
        
        acoes_iteracao = []
        for ac in acoes_org:
            acoes_iteracao.append(Acao(**ac))
            
        acoes_iteracao = normalizar_acoes(acoes_iteracao, parametros_premissas["novos_clientes"])
            
        proposta_correcao = PropostaMock(f"Swaps Iteracao {iteracao}", "Baixo", acoes_iteracao)
        
        sandbox_wb, sandbox_ws = clone_ws(new_ws)
        log_iter, _ = execute_alocacao(sandbox_ws, proposta_correcao, current_plant_data, allowed_cells, file_path=dest_file)
        
        erros_validacao_temp = validar_inventario(ws, sandbox_ws, allowed_cells, parametros_premissas)
        
        if not erros_validacao_temp:
            print(f"   ✓ Correção da Iteração {iteracao} validada com sucesso! Consolidando alterações físicas...")
            # Consolida usando o caminho físico dinâmico
            log_iter, _ = execute_alocacao(new_ws, proposta_correcao, current_plant_data, allowed_cells, file_path=dest_file)
            
            for k, v in log_iter.get('realocadas', {}).items():
                log_acumulado.setdefault('realocadas', {}).setdefault(k, []).extend(v)
            for k, v in log_iter.get('liberadas', {}).items():
                log_acumulado.setdefault('liberadas', {})[k] = v
            log_acumulado.setdefault('avisos', []).extend(log_iter.get('avisos', []))
            acoes_totais.extend(acoes_iteracao)
            erros_validacao = []  
            
            # Salva o arquivo intermediário após a rodada bem sucedida
            new_wb.save(dest_file)
        else:
            print(f"   ❌ Correção inválida na Iteração {iteracao}. Executando rollback...")
            print("\n      ⚠️ DETALHES DAS INCONSISTÊNCIAS NO LAYOUT DO ORGANIZADOR:")
            for err in erros_validacao_temp:
                print(f"      -> {err}")
            print()
            erros_validacao = erros_validacao_temp  

    # ══════════════════════════════════════════════════════════════════════════
    # ETAPA 3: Finalização e Escrita de Relatórios
    # ══════════════════════════════════════════════════════════════════════════
    print("\n3. Gravando resultados e relatórios finais...")
    new_wb.save(dest_file)
    
    final_plant_data = scan_plant(new_ws, FORBIDDEN_PATTERNS)
    final_blocos_info = build_blocos_info(final_plant_data, new_ws.max_row, new_ws.max_column, new_ws, file_path=dest_file)
    
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
        
    print(f"✓ Processo concluído com sucesso. Arquivos gravados em 'propostas/'.")

if __name__ == '__main__':
    asyncio.run(main())
# main.py

import asyncio
import os
import shutil
import openpyxl
from dotenv import load_dotenv

# Importando dependências do LayoutEngine modularizado
from LayoutEngine import (
    SHEET_NAME, FORBIDDEN_PATTERNS,
    Acao, PropostaMock, clean_json_string, extrair_e_carregar_json,
    load_plant, clone_ws, build_plant_info, build_blocos_info,
    execute_alocacao, write_report, salvar_auditoria,
    normalizar_acoes, validar_inventario
)

# Importando dependências e agentes do Agents.py (Fluxo de 2 Agentes)
from Agents import PosicionadorDeps, OrganizadorDeps, posicionador, organizador
from ScannerPremissas import scan_orange_context, build_context_string_for_llm, normalize_val
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
        "  - A primeira operação criada deve ser nomeada rigorosamente como: 'NOVO_A'\n"
        "  - A segunda operação criada deve ser nomeada rigorosamente como: 'NOVO_B'\n"
        "  - A terceira operação criada deve ser nomeada rigorosamente como: 'NOVO_C'\n"
        "E assim por diante. Use sempre letras maiúsculas e mantenha a consistência nominal.\n"
    )

    # Extrai o contexto visual das bordas laranjas e junta com as premissas do arquivo
    dados_laranjas = scan_orange_context('planta.xlsx', SHEET_NAME)
    premissas_visuais = build_context_string_for_llm(dados_laranjas)
    
    # Consolida as diretrizes do sistema com o texto de entrada
    premissas_completas = f"{premissas_txt}\n{regras_nomes_sistema}\n{premissas_visuais}"

    # Carrega a planta original e escaneia os dados iniciais
    wb, ws = load_plant()
    plant_data = scan_plant(ws, FORBIDDEN_PATTERNS)
    
    plant_info_str = build_plant_info(plant_data)
    blocos_info_str = build_blocos_info(plant_data, ws.max_row, ws.max_column, ws)
    
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
    # ETAPA 1: Execução do Posicionador com Loop de Auto-Correção Interno
    # ══════════════════════════════════════════════════════════════════════════
    print("1. Posicionador gerando ações primárias (alocação bruta)...")
    
    dest_file = "propostas/proposta_final.xlsx"
    safe_name = "final"
    os.makedirs('propostas', exist_ok=True)
    shutil.copy("planta.xlsx", dest_file)
    new_wb = openpyxl.load_workbook(dest_file)
    new_ws = new_wb[SHEET_NAME]

    erros_validacao = []
    rascunho_layout_json = ""
    pos_data = {}
    acoes_totais = []
    log_acumulado = {}
    proposta_inicial = None
    parametros_premissas = {"reducoes": {}, "novos_clientes": []}

    for pos_iter in range(1, 4):
        print(f"   -> Tentativa do Posicionador {pos_iter}/3...")
        pos_deps = PosicionadorDeps(
            plant_info=plant_info_str,
            blocos_info=blocos_info_str,
            premissas=premissas_completas
        )
        
        feedback_erros = ""
        if erros_validacao:
            feedback_erros = (
                "\n\n⚠️ INCONSISTÊNCIAS DE SOMA OPERACIONAL NO SEU LAYOUT ANTERIOR:\n"
                + "\n".join(f"- {err}" for err in erros_validacao)
                + "\n\nPor favor, recalcule as capacidades de destino, libere espaço e corrija suas quantidades!"
            )
            
        res_posicionador = await posicionador.run(
            f"Gere a proposta de alocação inicial baseada nas premissas de negócio.{feedback_erros}", 
            deps=pos_deps
        )
        rascunho_layout_json = clean_json_string(res_posicionador.output)
        pos_data = extrair_e_carregar_json(rascunho_layout_json)
        
        # Executa em sandbox temporário para validação estrita
        sandbox_wb, sandbox_ws = clone_ws(ws)
        acoes_totais = []
        for ac in pos_data.get("acoes_primarias", []):
            acoes_totais.append(Acao(**ac))
            
        # Extração inteligente do gabarito informado pela IA
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
                elif tipo in ("realocar", "posicionar") and cli.upper().startswith("NOVO"):
                    existente = next((nc for nc in novos_fallback if nc["nome"] == cli.upper()), None)
                    if existente:
                        existente["PAs"] += qtd
                    else:
                        novos_fallback.append({"nome": cli.upper(), "PAs": qtd})
            gabarito_dados = {"reducoes": reducoes_fallback, "novos_clientes": novos_fallback}
            
        # Normaliza e formata os dados do gabarito para a auditoria (programação defensiva)
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
        
        log_acumulado, _ = execute_alocacao(sandbox_ws, proposta_inicial, plant_data, allowed_cells)
        erros_validacao = validar_inventario(ws, sandbox_ws, allowed_cells, parametros_premissas)
        
        if not erros_validacao:
            print("   ✓ Layout inicial validado matematicamente com sucesso!")
            _, _ = execute_alocacao(new_ws, proposta_inicial, plant_data, allowed_cells)
            break
        else:
            print(f"   ❌ Tentativa {pos_iter} falhou na validação de inventário físico. Detalhes:")
            for err in erros_validacao:
                print(f"      -> {err}")
            print("\n      👉 Ações tentadas nesta rodada:")
            for ac in acoes_totais:
                tipo = ac.tipo.upper()
                if tipo in ('LIBERAR', 'REALOCAR', 'POSICIONAR'):
                    print(f"         - [{tipo}] Cliente '{ac.cliente}': Qtd {ac.quantidade} no Bloco {ac.bloco}-{ac.ambiente}")
            print()
            
    # Se todas as tentativas falharem, commita o rascunho imperfeito para o Organizador corrigir
    if erros_validacao and proposta_inicial is not None:
        print("   ⚠️ Gravando layout rascunho com inconsistências temporárias para que o Organizador realize as correções finas...")
        _, _ = execute_alocacao(new_ws, proposta_inicial, plant_data, allowed_cells)

    # Salva auditoria do posicionador
    salvar_auditoria("1_posicionador", (
        f"[PREMISSAS DO TXT]\n{pos_deps.premissas}\n\n"
        f"[BLOCOS INFO]\n{pos_deps.blocos_info}\n\n"
        f"[PLANT INFO]\n{pos_deps.plant_info}"
    ), res_posicionador.output, safe_name)

    if erros_validacao:
        print("⚠ Layout inicial possui erros não resolvidos pelo Posicionador. Entrando no loop do Organizador...")
    else:
        print("Iniciando loop de otimização espacial do Organizador...")

    # ══════════════════════════════════════════════════════════════════════════
    # ETAPA 2: Loop de Auto-Correção e Swaps Estáveis com o Organizador (Máx 3 Iterações)
    # ══════════════════════════════════════════════════════════════════════════
    print("\n2. Iniciando Loop de Auto-Correção do Organizador (Máximo 3 Iterações)...")
    
    for iteracao in range(1, 4):
        print(f"--- Iteração {iteracao}/3 ---")
        
        current_plant_data = scan_plant(new_ws, FORBIDDEN_PATTERNS)
        current_plant_info = build_plant_info(current_plant_data)
        current_blocos_info = build_blocos_info(current_plant_data, new_ws.max_row, new_ws.max_column, new_ws)
        
        org_deps = OrganizadorDeps(
            plant_info=current_plant_info,
            blocos_info=current_blocos_info,
            premissas=premissas_completas,
            rascunho_layout=rascunho_layout_json
        )
        
        feedback_erros = ""
        if erros_validacao:
            feedback_erros = (
                "\n\n⚠️ INCONSISTÊNCIAS DE SOMA OPERACIONAL NO LAYOUT ATUAL:\n"
                + "\n".join(f"- {err}" for err in erros_validacao)
                + "\n\nAplique ações de 'transferir' para reequilibrar as quantidades exatas de volta!"
            )
        
        res_organizador = await organizador.run(
            f"Avalie a planta física atual. Se houver qualquer violação de premissas (como unificação, exclusividade, Cliente 7 compartilhado), utilize a função 'transferir' para corrigir. Se estiver tudo certo, retorne a lista vazia.{feedback_erros}",
            deps=org_deps
        )
        reorganizacao_json = clean_json_string(res_organizador.output)
        
        salvar_auditoria(f"2_organizador_iter{iteracao}", (
            f"[RASCUNHO DO LAYOUT]\n{org_deps.rascunho_layout}\n\n"
            f"[PREMISSAS DO TXT]\n{org_deps.premissas}\n\n"
            f"[BLOCOS INFO ATUALIZADOS NA PLANILHA FISICA]\n{org_deps.blocos_info}\n\n"
            f"[PLANT INFO ATUALIZADA]\n{org_deps.plant_info}"
        ), res_organizador.output, safe_name)
        
        org_data = extrair_e_carregar_json(reorganizacao_json)
        acoes_org = org_data.get("acoes_organizacao", [])
        
        if not acoes_org and not erros_validacao:
            print(f"✓ Sistema auditado e validado. Todas as premissas e inventários estão plenamente satisfatórios!")
            break
            
        print(f"⚠ Executando ações de correção em sandbox para validação de integridade...")
        
        acoes_iteracao = []
        for ac in acoes_org:
            acoes_iteracao.append(Acao(**ac))
            
        acoes_iteracao = normalizar_acoes(acoes_iteracao, parametros_premissas["novos_clientes"])
            
        proposta_correcao = PropostaMock(f"Swaps Iteracao {iteracao}", "Baixo", acoes_iteracao)
        
        sandbox_wb, sandbox_ws = clone_ws(new_ws)
        log_iter, _ = execute_alocacao(sandbox_ws, proposta_correcao, current_plant_data, allowed_cells)
        
        erros_validacao_temp = validar_inventario(ws, sandbox_ws, allowed_cells, parametros_premissas)
        
        if not erros_validacao_temp:
            print(f"✓ Correção da Iteração {iteracao} validada com sucesso! Gravando no Excel master...")
            log_iter, _ = execute_alocacao(new_ws, proposta_correcao, current_plant_data, allowed_cells)
            
            for k, v in log_iter.get('realocadas', {}).items():
                log_acumulado.setdefault('realocadas', {}).setdefault(k, []).extend(v)
            for k, v in log_iter.get('liberadas', {}).items():
                log_acumulado.setdefault('liberadas', {})[k] = v
            log_acumulado.setdefault('avisos', []).extend(log_iter.get('avisos', []))
            acoes_totais.extend(acoes_iteracao)
            erros_validacao = []  
        else:
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
            erros_validacao = erros_validacao_temp  

    # ══════════════════════════════════════════════════════════════════════════
    # ETAPA 3: Finalização e Escrita de Relatórios
    # ══════════════════════════════════════════════════════════════════════════
    print("\n3. Gravando resultados e relatórios finais...")
    new_wb.save(dest_file)
    
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
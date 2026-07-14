"""Normalizacao deterministica das solicitacoes produzidas pelo posicionador."""

import re

from ScannerPremissas import normalize_val


def normalizar_nomes_novos_por_ordem(pos_data, premissas_txt):
    """Vincula N_1, N_2... a ordem das solicitacoes no texto, nao a ordem da LLM."""
    solicitacoes = []
    for linha in premissas_txt.splitlines():
        if not re.search(r"\bcriar\b", linha, re.IGNORECASE):
            continue
        match_pas = re.search(r"(\d+)(?:\s+nov[oa]s?)?\s*PAs?\b", linha, re.IGNORECASE)
        if not match_pas:
            continue
        match_sala = re.search(r"sala\s+(?:com\s+)?(\d+)\s*lug", linha, re.IGNORECASE)
        solicitacoes.append((int(match_pas.group(1)), int(match_sala.group(1)) if match_sala else 0))

    ambientes = pos_data.get("criar_ambientes", [])
    usados = set()
    renomeacoes = {}
    for indice, (pas_operacionais, sala_lugares) in enumerate(solicitacoes, start=1):
        for posicao, ambiente in enumerate(ambientes):
            if posicao in usados:
                continue
            qtd = int(ambiente.get("quantidade_mesas", 0) or 0)
            sala = int(ambiente.get("sala_lugares", 0) or 0)
            if sala == sala_lugares and qtd in (pas_operacionais, pas_operacionais + sala_lugares):
                nome_antigo = normalize_val(ambiente.get("cliente_destinado"))
                renomeacoes[nome_antigo] = f"N_{indice}"
                usados.add(posicao)
                break

    for ambiente in ambientes:
        nome = normalize_val(ambiente.get("cliente_destinado"))
        if nome in renomeacoes:
            ambiente["cliente_destinado"] = renomeacoes[nome]
    for cliente in pos_data.get("gabarito", {}).get("novos_clientes", []):
        if isinstance(cliente, dict):
            nome = normalize_val(cliente.get("nome"))
            if nome in renomeacoes:
                cliente["nome"] = renomeacoes[nome]
    for acao in pos_data.get("acoes_primarias", []):
        nome = normalize_val(acao.get("cliente"))
        if nome in renomeacoes:
            acao["cliente"] = renomeacoes[nome]
    return pos_data


def normalizar_quantidades_ambientes(criar_ambientes, novos_clientes, pas_explicitos):
    """Garante que quantidade_mesas exclua os lugares da sala interna."""
    clientes = {nc["nome"]: nc for nc in novos_clientes}
    for ambiente in criar_ambientes:
        cliente = normalize_val(ambiente.get("cliente_destinado"))
        sala = max(0, int(ambiente.get("sala_lugares", 0) or 0))
        ambiente["sala_lugares"] = sala
        if cliente in pas_explicitos:
            ambiente["quantidade_mesas"] = pas_explicitos[cliente]
            continue
        quantidade = max(0, int(ambiente.get("quantidade_mesas", 0) or 0))
        cadastro = clientes.get(cliente)
        if sala > 0 and cadastro is not None and cadastro["PAs"] == quantidade:
            quantidade = max(0, quantidade - sala)
            cadastro["PAs"] = quantidade
        ambiente["quantidade_mesas"] = quantidade
    return criar_ambientes
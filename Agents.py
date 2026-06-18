from dataclasses import dataclass
from typing import List
import re

from pydantic import BaseModel, Field
from pydantic_ai import Agent, RunContext, ModelRetry

import Prompts
from LLM import planner_model

# ══════════════════════════════════════════════════════════════════════════
# Schemas de Comunicação entre Agentes
# ══════════════════════════════════════════════════════════════════════════

@dataclass
class PlannerDeps:
    plant_info: str
    mapa_2d: str
    premissas: str
    blocos_info: str
    estrategia_macro: str = ""

class Estrategia(BaseModel):
    analise_premissas: str = Field(description="Interpretação das premissas")
    clientes_a_reduzir: List[str] = Field(description="Lista de clientes a reduzir com NOME EXATO (ex: ['1'] ou ['Cliente 1'])")
    nome_exato_cliente_principal: str = Field(description="Nome EXATO como aparece na planilha do cliente principal a reduzir (ex: '1' ou 'Cliente 1')")
    diretrizes_alocacao: str = Field(description="Diretriz de onde posicionar os novos clientes")

class Acao(BaseModel):
    tipo: str = Field(description="Deve ser estritamente 'liberar' ou 'realocar'.")
    cliente: str = Field(description="Nome do cliente afetado (ex: 'Cliente 1', 'Novo A')")
    quantidade: int = Field(description="Quantidade de PAs físicas")
    novo_cliente: str = Field(default="", description="Bloco destino (ex: 'vazio-1' ou 'vazio-3')")
    cliente_a_liberar: str = Field(default="", description="Se for realocar ou liberar em bloco específico, indique o cliente de origem (ex: '1' ou '2')")
    sala_lugares: int = Field(default=0, description="Quantidade de cadeiras adjacentes a serem convertidas em sala (ex: 4 para Novo A, 1 para Novo B)")

class Proposta(BaseModel):
    proposta: int
    nome: str
    descricao: str
    acoes: List[Acao]
    premissas_atendidas: List[str]
    catracas_novas: int
    custo_obras: str
    observacoes: str

# ══════════════════════════════════════════════════════════════════════════
# Agentes
# ══════════════════════════════════════════════════════════════════════════

orquestrador = Agent(model=planner_model, deps_type=PlannerDeps, output_type=Estrategia)

@orquestrador.system_prompt
def orquestrador_prompt(ctx: RunContext[PlannerDeps]) -> str:
    return Prompts.PLANNER_TMPL.format(
        plant_info=ctx.deps.plant_info, 
        mapa_2d="[Mapa Omitido]",
        premissas=ctx.deps.premissas, 
        blocos_info=ctx.deps.blocos_info
    ) + """

== INSTRUÇÃO FINAL OBRIGATÓRIA (LEI DE PRESERVAÇÃO DE POSIÇÕES) ==

Você é o ORQUESTRADOR. Defina a estratégia macro com estas regras matemáticas absolutas:

1. MAPEAR EXATAMENTE A REDUÇÃO:
   - A premissa pede para eliminar exatamente 270 posições do Cliente '1'.
   - Você deve planejar ações para reduzir EXATAMENTE 270 posições do Cliente '1' globais (nem mais, nem menos).

2. REGRA DE MÍNIMO IMPACTO (NÃO MOVER CLIENTES ESTÁVEIS):
   - **NUNCA movimente ou realoque clientes estáveis** (como o Cliente '2', Cliente '3', etc.) para outros blocos. Mover clientes estáveis gera custos de obra civil e risco de fragmentação secundária.
   - Sempre prefira balancear e ajustar as quantidades exatas de redução do Cliente '1' entre o Bloco 3 e o Bloco 6 para abrir espaço exato para o Novo A (124 PAs) e Novo B (165 PAs), mantendo os demais clientes 100% intocados em suas posições de origem.

3. USO DE LIBERAÇÕES CIRÚRGICAS POR BLOCO:
   - Divida as liberações do Cliente '1' indicando os blocos exatos e as quantidades exatas de redução (ex: liberar 137 no bloco 'vazio-3' e liberar 133 no bloco 'vazio-6'). Isso evita transbordos e garante o balanço de 270 exato.
"""

posicionador = Agent(model=planner_model, deps_type=PlannerDeps, output_type=Proposta, retries=5)

@posicionador.system_prompt
def posicionador_prompt(ctx: RunContext[PlannerDeps]) -> str:
    import json
    try:
        estrategia = json.loads(ctx.deps.estrategia_macro)
        nome_cliente = estrategia.get('nome_exato_cliente_principal', '1')
    except:
        nome_cliente = '1'
    
    return f"""Você é o POSICIONADOR. Siga RIGOROSAMENTE esta estratégia do Orquestrador:
{ctx.deps.estrategia_macro}

⚠️  AVISO CRÍTICO — NOMENCLATURA DOS NOVOS CLIENTES:
═════════════════════════════════════════════════════════
Os nomes DOS NOVOS CLIENTES são EXATAMENTE: "Novo A" e "Novo B"
NÃO use variações como: "Novo Cliente A", "Novo Cliente B"

ERRADO:  {{"cliente": "Novo Cliente A", ...}}
CORRETO: {{"cliente": "Novo A", ...}}

REGRAS OBRIGATÓRIAS DE TRANSAÇÃO (SOMA ZERO):

1. ORDEM DE EXECUÇÃO:
   - As ações de "liberar" sempre serão executadas na Fase 1 (criando o espaço livre '0').
   - As ações de "realocar" sempre serão executadas na Fase 2 (posicionando apenas sobre espaço vazio ou '0').

2. REDUÇÃO CIRÚRGICA DO CLIENTE PRINCIPAL '{nome_cliente}' POR BLOCO:
   - Divida as ações de 'liberar' indicando os blocos exatos e as quantidades exatas de redução (ex: liberar 137 no bloco 'vazio-3' e liberar 133 no bloco 'vazio-6').
   - O campo 'novo_cliente' na ação 'liberar' deve receber o ID do bloco de origem correspondente (ex: 'vazio-3' ou 'vazio-6').
   - A soma de todas as ações de 'liberar' de '{nome_cliente}' deve totalizar exatamente 270 PAs.

3. PARA OS NOVOS CLIENTES ("Novo A" e "Novo B"):
   - Gere ações de 'realocar' indicando o bloco de destino correto:
     - Use o ID do bloco laranjas correspondente (ex: 'vazio-6' ou 'vazio-3').
     - O campo 'cliente_a_liberar' deve receber o cliente de origem ('1') se estiver usando o espaço reduzido dele naquele bloco.
     - Nunca misture Novo A e Novo B no mesmo bloco!

4. REGRA DE BALANÇO:
   - Para alocar um cliente em um bloco, certifique-se de que a quantidade de vagas vazias originais daquele bloco somada às vagas que você liberou nele seja MAIOR ou IGUAL à demanda do novo cliente. O script de alocação não faz despejos automáticos.

5. MATEMÁTICA:
   - Novo A: aloque apenas as 124 PAs operacionais (as salas de reunião devem ser ignoradas fisicamente na planilha).
   - Novo B: aloque apenas as 165 PAs operacionais (as salas de reunião devem ser ignoradas fisicamente na planilha).

RETORNE JSON estrito. Use EXATAMENTE "Novo A" ou "Novo B" nos respectivos campos de novos clientes."""


@posicionador.output_validator
def validate_proposta(ctx: RunContext[PlannerDeps], proposta: Proposta) -> Proposta:
    import re
    # 1. Extrai a capacidade total real (física) e as vagas vazias de cada bloco laranjas
    block_sections = re.split(r'--------------------', ctx.deps.blocos_info)
    capacidades_totais = {}
    vagas_vazias_originais = {}
    
    for section in block_sections:
        match_id = re.search(r'\(vazio-(\d+)\)', section)
        match_total = re.search(r'Tamanho total do bloco:\s*(\d+)', section)
        match_empty = re.search(r'Células sem clientes:\s*(\d+)', section)
        
        if match_id:
            b_id = f"vazio-{match_id.group(1)}"
            if match_total:
                capacidades_totais[b_id] = int(match_total.group(1))
            if match_empty:
                vagas_vazias_originais[b_id] = int(match_empty.group(1))
            
    valid_new_clients = {"NOVO A", "NOVO B"}
    invalid_names = []
    
    for acao in proposta.acoes:
        if acao.tipo in ('realocar', 'alocar'):
            cliente_upper = str(acao.cliente).strip().upper()
            if 'NOVO' in cliente_upper:
                if cliente_upper not in valid_new_clients:
                    invalid_names.append(cliente_upper)
    
    if invalid_names:
        raise ModelRetry(
            f"ERRO CRÍTICO DE NOMENCLATURA: Nomes ERRADOS encontrados: {invalid_names}\n"
            f"Devem ser EXATAMENTE: 'Novo A' ou 'Novo B'."
        )
    
    # 2. Mapeia as liberações planejadas de cada cliente por bloco de destino
    liberacoes_por_bloco = {}  # { (cliente, bloco): quantidade }
    vagas_liberadas_globais = {}  # { cliente: quantidade }
    
    for a in proposta.acoes:
        if a.tipo == 'liberar':
            cli_key = str(a.cliente_a_liberar or a.cliente).strip().upper()
            vagas_liberadas_globais[cli_key] = vagas_liberadas_globais.get(cli_key, 0) + a.quantidade
            
            bloco_origem = a.novo_cliente.strip()
            if bloco_origem:
                liberacoes_por_bloco[(cli_key, bloco_origem)] = liberacoes_por_bloco.get((cli_key, bloco_origem), 0) + a.quantidade
    
    # 3. Validação de Balanço Geográfico por Bloco (Evita over-reduction e transbordo)
    destinos_por_cliente = {}

    for acao in proposta.acoes:
        if acao.tipo in ('realocar', 'alocar'):
            cliente_clean = re.sub(r'-(complemento|parte|excedente|residuo)', '', str(acao.cliente).strip(), flags=re.IGNORECASE).strip()
            match_novo = re.search(r'novo[-\s]*(?:cliente[-\s]*)?([a-zA-Z0-9]+)', cliente_clean, flags=re.IGNORECASE)
            cliente_clean = f"Novo {match_novo.group(1).upper()}" if match_novo else cliente_clean[:15]
            
            dest = acao.novo_cliente.strip()
            destinos_por_cliente.setdefault(cliente_clean, set()).add(dest)
            
            total_solicitado = acao.quantidade + acao.sala_lugares
            
            # Validação física: não pode alocar mais do que o tamanho do bloco suporta
            if dest in capacidades_totais and total_solicitado > capacidades_totais[dest]:
                raise ModelRetry(f"ERRO: {total_solicitado} posições excede a capacidade física total de {capacidades_totais[dest]} do bloco {dest}.")
            
            # Validação de Balanço Geográfico: vagas vazias + vagas liberadas planejadas naquele bloco
            if dest in vagas_vazias_originais:
                vazio_disponivel = vagas_vazias_originais[dest]
                
                cli_origem_key = str(acao.cliente_a_liberar or '').strip().upper()
                liberado_disponivel = liberacoes_por_bloco.get((cli_origem_key, dest), 0)
                
                total_vagas_geograficas = vazio_disponivel + liberado_disponivel
                
                if total_solicitado > total_vagas_geograficas:
                    raise ModelRetry(
                        f"ERRO DE BALANÇO GEOGRÁFICO: Você está tentando alocar {total_solicitado} PAs do {acao.cliente} no bloco {dest}, "
                        f"mas o bloco {dest} só possui {vazio_disponivel} vagas vazias originais e você só planejou liberar {liberado_disponivel} PAs do Cliente {cli_origem_key} neste bloco específico. "
                        f"Isso causará um transbordo de PAs ou um despejo forçado indevido que reduzirá o Cliente {cli_origem_key} além do limite! "
                        f"Ajuste as ações de 'liberar' e 'realocar' para que as quantidades batam exatamente dentro do mesmo bloco."
                    )
            
            elif "-liberado" in dest.lower():
                orig = dest.replace("-liberado", "").replace("-LIBERADO", "").strip().upper()
                vagas_disponivel = vagas_liberadas_globais.get(orig, 0)
                if total_solicitado > (vagas_disponivel + 40):
                    raise ModelRetry(f"ERRO: {total_solicitado} posições excede as {vagas_disponivel} liberadas de {orig}.")

    for c, dests in destinos_por_cliente.items():
        limite = 3 if c in ('1', '3', '4') else 2
        if len(dests) > limite:
            raise ModelRetry(f"ERRO DE FRAGMENTAÇÃO: Limite de {limite} blocos excedido para '{c}'.")
    
    return proposta
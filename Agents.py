from dataclasses import dataclass
from typing import List, Dict, Tuple
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
    cliente: str = Field(description="Nome do cliente afetado (ex: '1', 'Novo A', 'Novo B')")
    quantidade: int = Field(description="Quantidade de PAs físicas")
    novo_cliente: str = Field(default="", description="Ambiente destino ou origem (ex: 'vazio-3-A' ou 'vazio-6-A')")
    cliente_a_liberar: str = Field(default="", description="Se for realocar ou liberar em ambiente específico, indique o cliente de origem (ex: '1' ou '2')")
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
   - **NUNCA movimente ou realoque clientes estáveis** (como o Cliente '2', Cliente '3', Cliente '0', etc.) para outros blocos. Mover clientes estáveis gera custos de obra civil e risco de fragmentação secundária.
   - Sempre prefira balancear e ajustar as quantidades exatas de redução do Cliente '1' entre o sub-ambiente 'vazio-3-A' e o sub-ambiente 'vazio-6-A' para abrir espaço exato para o Novo A (124 PAs) e Novo B (165 PAs), mantendo os demais clientes 100% intocados em suas posições de origem.

3. USO DE LIBERAÇÕES CIRÚRGICAS POR AMBIENTE:
   - Divida as liberações do Cliente '1' indicando os sub-ambientes exatos e as quantidades exatas de redução (ex: liberar 137 no ambiente 'vazio-3-A' e liberar 133 no ambiente 'vazio-6-A'). Isso evita transbordos e garante o balanço de 270 exato.
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
   - As ações de "liberar" sempre serão executadas na Fase 1 (escrevendo a palavra 'vazio').
   - As ações de "realocar" sempre serão executadas na Fase 2 (posicionando apenas sobre espaço vazio ou 'vazio').

2. REDUÇÃO CIRÚRGICA DO CLIENTE PRINCIPAL '{nome_cliente}' POR AMBIENTE:
   - Divida as ações de 'liberar' indicando os ambientes exatos e as quantidades exatas de redução (ex: liberar no ambiente 'vazio-3-A' ou 'vazio-6-A').
   - O campo 'novo_cliente' na ação 'liberar' deve receber o ID do ambiente de origem correspondente (ex: 'vazio-3-A' ou 'vazio-6-A').
   - A soma de todas as ações de 'liberar' de '{nome_cliente}' deve totalizar exatamente 270 PAs.

3. PARA OS NOVOS CLIENTES ("Novo A" e "Novo B"):
   - Gere ações de 'realocar' indicando o ambiente de destino correto:
     - Use o ID do ambiente correspondente (ex: 'vazio-6-A' ou 'vazio-3-A').
     - O campo 'cliente_a_liberar' deve receber o cliente de origem ('1') se estiver usando o espaço reduzido dele naquele ambiente.
     - Nunca misture Novo A e Novo B no mesmo ambiente!

4. REGRA DE BALANÇO:
   - Para alocar um cliente em um ambiente, certifique-se de que a quantidade de vagas vazias originais daquele ambiente somada às vagas que você liberou nele seja MAIOR ou IGUAL à demanda do novo cliente. O script de alocação não faz despejos automáticos.

5. MATEMÁTICA:
   - Novo A: aloque apenas as 124 PAs operacionais (as salas de reunião devem ser ignoradas fisicamente na planilha).
   - Novo B: aloque apenas as 165 PAs operacionais (as salas de reunião devem ser ignoradas fisicamente na planilha).

RETORNE JSON estrito. Use EXATAMENTE "Novo A" ou "Novo B" nos respectivos campos de novos clientes."""


@posicionador.output_validator
def validate_proposta(ctx: RunContext[PlannerDeps], proposta: Proposta) -> Proposta:
    from collections import defaultdict
    
    # 1. Mapeia capacidades de ambientes, vagas vazias e clientes originais extraídos do blocos_info
    capacidades_totais = {}
    vagas_vazias_originais = {}
    clientes_originais = {}
    
    # Varre a string de blocos_info gerada dinamicamente pelo ScannerPremissas
    env_matches = re.finditer(
        r"Ambiente\s+([A-K])\s+\((vazio-\d+-[A-K])\):"
        r".*?Células totais mapeadas fisicamente neste ambiente:\s*(\d+)"
        r".*?Células sem clientes\s*/\s*em branco:\s*(\d+)"
        r"(.*?)(?=Ambiente\s+[A-K]|--------------------|━━━━━━━━━━━━━━━━|==|$)",
        ctx.deps.blocos_info, re.DOTALL | re.IGNORECASE
    )
    
    for m in env_matches:
        env_id = m.group(2).lower()  # ex: vazio-1-a
        cap = int(m.group(3))
        vazias = int(m.group(4))
        clientes_text = m.group(5)
        
        capacidades_totais[env_id] = cap
        vagas_vazias_originais[env_id] = vazias
        
        clientes_originais[env_id] = {}
        for cli_match in re.finditer(r"Cliente\s+'([^']+)':\s*quantidade:\s*(\d+)", clientes_text):
            cli_name = cli_match.group(1).upper()
            qty = int(cli_match.group(2))
            clientes_originais[env_id][cli_name] = qty
            
    # 2. Validação estrita de nomes de novos clientes
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
        
    # 3. Máquina de estados para calcular a ocupação final de cada ambiente
    estado_clientes = defaultdict(dict)
    for env_id, clis in clientes_originais.items():
        for cli, qty in clis.items():
            estado_clientes[cli][env_id] = qty
            
    liberacoes_por_env = {}  # { (cliente, ambiente): quantidade }
    
    # Processa as liberações primeiro (Fase 1)
    for acao in proposta.acoes:
        if acao.tipo == 'liberar':
            dest = str(acao.novo_cliente).strip().lower()  # ex: vazio-3-a
            cli_origem = str(acao.cliente_a_liberar or acao.cliente).strip().upper()
            liberacoes_por_env[(cli_origem, dest)] = liberacoes_por_env.get((cli_origem, dest), 0) + acao.quantidade
            
            if dest in estado_clientes[cli_origem]:
                estado_clientes[cli_origem][dest] -= acao.quantidade
                if estado_clientes[cli_origem][dest] <= 0:
                    del estado_clientes[cli_origem][dest]
            else:
                raise ModelRetry(
                    f"ERRO: Tentativa de liberar {acao.quantidade} do Cliente {cli_origem} no ambiente {dest}, "
                    f"mas esse cliente não existe originalmente neste ambiente."
                )
                
    # Processa as alocações/realocações (Fase 2)
    for acao in proposta.acoes:
        if acao.tipo in ('realocar', 'alocar'):
            dest = str(acao.novo_cliente).strip().lower()  # ex: vazio-3-a
            cli_name = str(acao.cliente).strip().upper()
            total_solicitado = acao.quantidade + acao.sala_lugares
            
            if dest not in capacidades_totais:
                raise ModelRetry(
                    f"ERRO DE GEOMETRIA: O ambiente '{dest}' indicado como destino não existe na planta."
                )
                
            if total_solicitado > capacidades_totais[dest]:
                raise ModelRetry(
                    f"ERRO DE EXCESSO FÍSICO: Alocação de {total_solicitado} PAs excede a capacidade total "
                    f"de {capacidades_totais[dest]} do ambiente {dest}."
                )
                
            vazio_disponivel = vagas_vazias_originais.get(dest, 0)
            cli_origem_key = str(acao.cliente_a_liberar or '').strip().upper()
            liberado_disponivel = liberacoes_por_env.get((cli_origem_key, dest), 0)
            
            total_vagas_disponiveis = vazio_disponivel + liberado_disponivel
            
            if total_solicitado > total_vagas_disponiveis:
                raise ModelRetry(
                    f"ERRO DE BALANÇO GEOGRÁFICO: Tentando alocar {total_solicitado} PAs do {acao.cliente} no ambiente {dest}, "
                    f"mas o ambiente {dest} possui apenas {vazio_disponivel} vagas vazias originais e você liberou {liberado_disponivel} PAs neste ambiente específico. "
                    f"Isso causará transbordo!"
                )
                
            estado_clientes[cli_name][dest] = estado_clientes[cli_name].get(dest, 0) + acao.quantidade
            
    # 4. Validação estrita da regra de não-fragmentação por ambiente
    for cli, envs in estado_clientes.items():
        # Exceção explícita prevista na premissa (Agora o 0 é cliente estável e validado no limite de 2)
        if cli in ('1', '3', '4'):
            continue
            
        # Filtra apenas os ambientes que terminaram com quantidade positiva de PAs deste cliente
        ambientes_ocupados = [env for env, qty in envs.items() if qty > 0]
        
        if len(ambientes_ocupados) > 2:
            raise ModelRetry(
                f"ERRO DE FRAGMENTAÇÃO POR AMBIENTE: O Cliente '{cli}' está dividido em {len(ambientes_ocupados)} ambientes "
                f"({', '.join(ambientes_ocupados)}), violando a premissa de não ficar em mais de 2 ambientes. "
                f"Os únicos clientes que podem exceder esse limite são: 1, 3 e 4. Consolide as equipes!"
            )
            
    return proposta
# Agents.py

from dataclasses import dataclass
from typing import List, Dict, Set, Tuple
from pydantic_ai import Agent, RunContext
import Prompts
from LLM import planner_model

# ══════════════════════════════════════════════════════════════════════════
# Estruturas de Dados de Entrada
# ══════════════════════════════════════════════════════════════════════════

@dataclass
class PosicionadorDeps:
    plant_info: str
    blocos_info: str
    premissas: str
    
@dataclass
class OrganizadorDeps:
    plant_info: str
    blocos_info: str
    premissas: str
    rascunho_layout: str     # Adicione esta linha
    ambientes_criados: str   # Adicione esta linha

# ══════════════════════════════════════════════════════════════════════════
# Inicialização dos Agentes (Fluxo Unificado de 2 Etapas)
# ══════════════════════════════════════════════════════════════════════════

posicionador = Agent(
    model=planner_model, 
    deps_type=PosicionadorDeps, 
    output_type=str
)

@posicionador.system_prompt
def posicionador_prompt(ctx: RunContext[PosicionadorDeps]) -> str:
    return Prompts.POSICIONADOR_TMPL.format(
        plant_info=ctx.deps.plant_info,
        premissas=ctx.deps.premissas,
        blocos_info=ctx.deps.blocos_info
    )

organizador = Agent(
    model=planner_model, 
    deps_type=OrganizadorDeps, 
    output_type=str
)

@organizador.system_prompt
def organizador_prompt(ctx: RunContext[OrganizadorDeps]) -> str:
    return Prompts.ORGANIZADOR_TMPL.format(
        premissas=ctx.deps.premissas,
        blocos_info=ctx.deps.blocos_info,
        ambientes_criados=ctx.deps.ambientes_criados,
        rascunho_layout=ctx.deps.rascunho_layout  # <--- ADICIONE ESTA LINHA
    )
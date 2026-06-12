"""
Agents.py
=========
Dois agentes:

  PlannerAgent (orquestrador)
    - Lê premissas + mapa + resumo da planta
    - Decide o que liberar e quando criar novos espaços
    - Para cada "criar espaço" chama a tool `criar_espaco`
      que internamente invoca o BlockDesignerAgent

  BlockDesignerAgent (especialista)
    - Recebe: nome do espaço, n_pas, n_sala, canvas disponível
    - Raciocina sobre o layout ideal (ilhas, corredores, catraca)
    - Devolve: BlocoLayout com ilhas retangulares em coords relativas
"""

from dataclasses import dataclass
from typing import List, Optional
from pydantic import BaseModel, Field
from pydantic_ai import Agent, RunContext

import Prompts
from LLM import planner_model


# ══════════════════════════════════════════════════════════════════════════
# Schemas compartilhados
# ══════════════════════════════════════════════════════════════════════════

class Ilha(BaseModel):
    """Região retangular dentro do bloco (coords relativas ao canvas)."""
    tipo:       str = Field(description="'PA', 'SALA' ou 'CATRACA'")
    row_offset: int = Field(description="Linha do topo da ilha. 0 = primeira linha dentro da parede. -1 = corredor externo (para CATRACA)")
    col_offset: int = Field(description="Coluna da esquerda da ilha")
    altura:     int = Field(default=1, description="Número de linhas")
    largura:    int = Field(default=1, description="Número de colunas")


class BlocoLayout(BaseModel):
    """Layout completo de um novo espaço em ilhas retangulares."""
    nome:          str        = Field(description="Nome do espaço, ex: 'ESP-A'")
    altura:        int        = Field(description="Altura total do bloco (linhas dentro da parede, sem contar o corredor externo)")
    largura:       int        = Field(description="Largura total do bloco em colunas")
    ilhas:         List[Ilha] = Field(description="Ilhas PA, SALA e CATRACA")
    justificativa: str        = Field(default="")


# ══════════════════════════════════════════════════════════════════════════
# BlockDesignerAgent — especialista em layout de blocos
# ══════════════════════════════════════════════════════════════════════════

@dataclass
class DesignerDeps:
    nome:      str   # nome do espaço (ex: "ESP-A")
    n_pas:     int   # número de PAs a posicionar
    n_sala:    int   # lugares na sala
    canvas:    str   # descrição do espaço disponível (linha/col de origem)


block_designer_agent: Agent[DesignerDeps, BlocoLayout] = Agent(
    model=planner_model,
    deps_type=DesignerDeps,
    output_type=BlocoLayout,
)


@block_designer_agent.system_prompt
def designer_system_prompt(ctx: RunContext[DesignerDeps]) -> str:
    return Prompts.BLOCK_DESIGNER_TMPL.format(
        nome=ctx.deps.nome,
        n_pas=ctx.deps.n_pas,
        n_sala=ctx.deps.n_sala,
        canvas=ctx.deps.canvas,
    )


# ══════════════════════════════════════════════════════════════════════════
# PlannerAgent — orquestrador
# ══════════════════════════════════════════════════════════════════════════

@dataclass
class PlannerDeps:
    plant_info:   str   # zonas, contagens
    mapa_2d:      str   # grid ASCII + paredes + catracas
    premissas:    str
    # canvas_map é injetado pelo main antes de executar o agente
    # {nome_espaco: {row_start, col_start, n_pas, n_sala}}
    canvas_map:   dict


class Acao(BaseModel):
    tipo:              str  = Field(description="'liberar' ou 'alocar'")
    setor:             str  = Field(description="Valor do cliente/setor na planta")
    quantidade:        int  = Field(description="Número de PAs")
    cliente:           str  = Field(default="")
    cliente_nome:      str  = Field(default="")
    cliente_a_liberar: str  = Field(default="")
    sala_lugares:      int  = Field(default=0)
    justificativa:     str  = Field(default="")
    layout:            Optional[BlocoLayout] = Field(
        default=None,
        description="Preenchido automaticamente pela tool criar_espaco"
    )


class Proposta(BaseModel):
    proposta:            int
    nome:                str
    descricao:           str
    acoes:               List[Acao]
    premissas_atendidas: List[str]
    catracas_novas:      int
    custo_obras:         str
    observacoes:         str


planner_agent: Agent[PlannerDeps, Proposta] = Agent(
    model=planner_model,
    deps_type=PlannerDeps,
    output_type=Proposta,
)


@planner_agent.system_prompt
def planner_system_prompt(ctx: RunContext[PlannerDeps]) -> str:
    return Prompts.PLANNER_TMPL.format(
        plant_info=ctx.deps.plant_info,
        mapa_2d=ctx.deps.mapa_2d,
        premissas=ctx.deps.premissas,
    )


@planner_agent.tool
async def criar_espaco(ctx: RunContext[PlannerDeps], nome: str, n_pas: int, n_sala: int) -> dict:
    """
    Cria o layout de um novo espaço apartado.

    Chame esta tool uma vez por espaço a criar, passando:
      nome  — identificador único do espaço (ex: 'ESP-A', 'ESP-B')
      n_pas — número exato de PAs conforme as premissas
      n_sala — número exato de lugares na sala conforme as premissas

    Retorna o layout completo (ilhas, corredores, catraca) que será
    aplicado automaticamente na planilha.
    """
    canvas_map = ctx.deps.canvas_map

    # Busca canvas por nome exato, depois por ordem (próximo não usado)
    canvas_info = canvas_map.get(nome)
    if canvas_info is None or canvas_info.get('_used'):
        for k, v in canvas_map.items():
            if not v.get('_used') and not v.get('_layout'):
                canvas_info = v
                break

    if canvas_info is None:
        return {'erro': f'Canvas não disponível para {nome}'}

    canvas_desc = (
        f"Espaço '{nome}': inicia na linha {canvas_info['row_start']}, "
        f"coluna {canvas_info['col_start']} (índice numérico). "
        f"Você tem espaço livre à direita e abaixo."
    )

    print(f"  [BlockDesigner] Projetando '{nome}': {n_pas} PAs + {n_sala} sala...")
    deps = DesignerDeps(nome=nome, n_pas=n_pas, n_sala=n_sala, canvas=canvas_desc)
    result = await block_designer_agent.run(
        f"Projete o layout do espaço '{nome}' com {n_pas} PAs e sala de {n_sala} lugares.",
        deps=deps,
    )
    layout = result.output
    print(f"  [BlockDesigner] '{nome}' projetado: {len(layout.ilhas)} ilhas, "
          f"{layout.altura}L × {layout.largura}C")
    for ilha in layout.ilhas:
        print(f"    {ilha.tipo:8} row={ilha.row_offset:>3} col={ilha.col_offset:>3} "
              f"{ilha.altura}L×{ilha.largura}C  ({ilha.altura * ilha.largura} células)")

    # Salva o layout e o nome usado no canvas_map para o execute recuperar
    canvas_info['_used']   = True
    canvas_info['_layout'] = layout
    canvas_info['_nome']   = nome   # nome que a LLM usou ao chamar a tool

    return layout.model_dump()

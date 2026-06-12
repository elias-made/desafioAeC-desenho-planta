# ── Orquestrador ────────────────────────────────────────────────────────

PLANNER_TMPL = """
Você é um especialista em space planning de call centers.
Interprete as premissas e orquestre as mudanças na planta.

== PLANTA ==
{plant_info}

== MAPA 2D ==
{mapa_2d}

== PREMISSAS ==
{premissas}

== INSTRUÇÕES ==

Para cada "Eliminar X posições cliente N":
  → Gere uma ação "liberar" com o cliente e a quantidade exata.

Para cada "Criar novo espaço apartado":
  → Chame a tool criar_espaco(nome, n_pas, n_sala) — uma chamada por espaço.
  → Use nomes sequenciais: ESP-A, ESP-B, ESP-C...

Respeite todas as restrições das premissas (salas proibidas, divisão de clientes, custo).

REGRAS:
- Números exatos conforme as premissas
- criar_espaco apenas uma vez por espaço
- layout null para ações "liberar"
"""


# ── Designer de Blocos ───────────────────────────────────────────────────

BLOCK_DESIGNER_TMPL = """
Você é um arquiteto de call centers. Projete o layout de UM espaço apartado.

== ESPAÇO ==
Nome:   {nome}
PAs:    {n_pas}
Sala:   {n_sala} lugar(es)
Canvas: {canvas}

== ELEMENTOS ==

PA: posição de atendimento. Agrupe em ilhas retangulares separadas por corredores.
SALA: área fechada dentro do espaço (reuniões/supervisão).
CATRACA: controle de acesso no corredor externo (row_offset = -1, fora da parede).
  Quantidade: ceil({n_pas} / 250). Posicione onde o fluxo de pessoas for maior. NUNCA DEVE ESTAR POSICIONADO NA CÉLULA AO LADO DE UMA SALA, O POSICIONAMENTO SEMPRE DEVE ESTAR RELACIONADO AO FLUXO DOS CORREDORES.

As paredes são desenhadas automaticamente no bounding box de PA+SALA — não as inclua.

== OTIMIZE O LAYOUT ==

Você tem liberdade para decidir a melhor forma de organizar {n_pas} PAs e a sala.
Pense como um arquiteto: forme um espaço retangular funcional, com circulação clara
e boa densidade. Algumas diretrizes de referência (não obrigatórias):

  - Ilhas de PA costumam ter 2 colunas de largura e 6-10 linhas de altura
  - Corredores entre ilhas devem ter pelo menos 2 células para circulação
  - A sala fica melhor no final ou lateral do bloco, acessível pelo corredor
  - A catraca deve estar na entrada principal, não encostada numa parede de sala
  - Os espaço mínimo dos corredores é uma linha ou coluna, e no máximo 2 linhas ou colunas, depende da direção do fluxo.

== RESTRIÇÕES (obrigatórias) ==

1. soma(altura × largura) das ilhas PA = exatamente {n_pas}
2. soma(altura × largura) das ilhas SALA = exatamente {n_sala}
3. Nenhuma ilha se sobrepõe a outra
4. CATRACA sempre com row_offset = -1
5. Retorne apenas o JSON, sem texto adicional

== SAÍDA ==

{{
  "nome": "{nome}",
  "altura": <linhas totais do bloco, excluindo o corredor externo>,
  "largura": <colunas totais>,
  "justificativa": "<suas decisões de layout em uma frase>",
  "ilhas": [
    {{"tipo": "PA",      "row_offset": <r>, "col_offset": <c>, "altura": <h>, "largura": <w>}},
    {{"tipo": "SALA",    "row_offset": <r>, "col_offset": <c>, "altura": <h>, "largura": <w>}},
    {{"tipo": "CATRACA", "row_offset": -1,  "col_offset": <c>, "altura": 1,   "largura": 1}}
  ]
}}
"""

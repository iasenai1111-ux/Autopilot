---
description: Monta ou melhora o painel no 3_painel.py a partir de um pedido em texto
---
Leia `projeto/modelo.json` (tabelas, colunas e medidas) e o arquivo `3_painel.py`.
Atenda ao pedido do aluno: ${input:pedido:O que o painel precisa responder? Ex.: quais obras estouraram o orçamento}

Regras:
- Use apenas os métodos de `ap.Painel`: segmentacao, cartao, barras, colunas, linhas, pizza, tabela.
- No máximo 2 filtros, 4 cartões e 4 gráficos (o painel tem que caber numa tela).
- Cada gráfico tem um `titulo` em forma de pergunta de negócio.
- Comente cada linha com `#` dizendo que pergunta aquele visual responde.
- No fim, diga ao aluno para rodar `python 3_painel.py` e reabrir `projeto/Obras.pbip` no Power BI Desktop.

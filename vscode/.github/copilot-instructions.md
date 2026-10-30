# Instruções para o GitHub Copilot — BI Autopilot (SENAI)

Este projeto é usado em sala de aula por alunos iniciantes de Power BI e análise de dados.
Responda sempre em português do Brasil, com frases curtas e sem jargão.

## Como o projeto funciona
1. `1_tratar.py` — o aluno limpa a planilha suja com pandas e salva `dados/base_tratada.xlsx`.
2. `2_modelar.py` — o aluno escreve os relacionamentos e as medidas DAX (num dicionário Python).
3. `3_painel.py` — o aluno lista os visuais do painel, uma linha por visual.
4. `autopilot.py` é o motor que transforma isso no projeto oficial do Power BI (PBIP: TMDL + PBIR) em `projeto/`.

## Regras
- Edite somente `1_tratar.py`, `2_modelar.py` e `3_painel.py`. Não edite `autopilot.py` nem os arquivos de `projeto/` (eles são gerados).
- Coluna se escreve `Tabela[Coluna]`. Medida se escreve só pelo nome.
- Medida nunca pode ter o mesmo nome de uma coluna.
- Use somente tabelas e colunas que existem em `projeto/modelo.json`.
- Visuais disponíveis em `3_painel.py`: `segmentacao(coluna)`, `cartao(medida)`, `barras(categoria, valor)`,
  `colunas(categoria, valor)`, `linhas(eixo, [valores])`, `pizza(categoria, valor)`, `tabela([campos])`.
  Todos aceitam `titulo="..."`, menos `segmentacao` e `cartao`.
- Toda sugestão vem com um comentário `#` curto explicando o que faz e por que ajuda o negócio.
- Nunca entregue a resposta de um exercício. Se o aluno pedir, dê uma pista e mostre um exemplo com outro tema.
- Depois de mudar um arquivo, diga ao aluno qual comando rodar (`python 2_modelar.py`, por exemplo) e o que conferir no Power BI.

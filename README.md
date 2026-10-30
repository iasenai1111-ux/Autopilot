# 🤖 BI Autopilot — da planilha suja ao painel do Power BI

> **O Python calcula. A IA explica. Você confere.**

Material do curso **Power BI e Análise de Dados com IA — SENAI**.
O aluno monta um tutor de IA com até três provedores, escreve as peças de um motor de limpeza de dados e gera um
**projeto oficial do Power BI (PBIP)** a partir da planilha suja de uma construtora.

## ▶ Começar pelo Colab (nada para instalar)

[![Abrir no Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/iasenai1111-ux/Autopilot/blob/main/notebooks/Monte_seu_Tutor_e_seu_Motor.ipynb)

1. Abra o notebook pelo botão acima e faça **Arquivo → Salvar uma cópia no Drive**.
2. Siga as partes 0 a 6: imagem explicando → lacunas `______` para completar → resultado esperado.
3. No fim, o Colab baixa o **Obras.zip**. Extraia em qualquer pasta, abra `projeto\Obras.pbip`
   no Power BI Desktop e clique em **Atualizar** (os dados já vão dentro do projeto).

O tutor tenta até **3 opções de IA**. Se uma falhar, demorar ou retornar resposta vazia, tenta a próxima.
Todos os planos têm limites e podem ficar indisponíveis. A célula de preparação instala as bibliotecas.
O notebook fixa o commit do motor para manter a compatibilidade da aula.

| Plano | Chave | Destaque |
|---|---|---|
| A · Colab AI | sem chave separada | depende da conta e dos limites do Colab |
| B · Groq | console.groq.com (`GROQ_API_KEY`) | opção com API compatível com o cliente OpenAI |
| C · Gemini | aistudio.google.com (`GEMINI_API_KEY`) | opcional, com limites e políticas conforme o plano |

As chaves ficam nos **Segredos 🔑** do Colab, nunca no código.
Confira as políticas atuais de [Colab](https://research.google.com/colaboratory/faq.html),
[Groq](https://console.groq.com/docs/your-data) e [Gemini](https://ai.google.dev/gemini-api/terms).
Envie somente os dados fictícios da aula. As regras orientam o tutor, mas não garantem acerto.

O tutor recebe contexto explícito e dá pistas em **Python 3/pandas**. Para medidas do Power BI,
usa **DAX** com o resumo do modelo anexado à pergunta. Ele não vê imagens nem células anteriores automaticamente.
O modo exercício preserva as lacunas: pedir uma função pronta deve gerar uma pista relacionada em Python.

## Material da aula

- [Apresentação](materiais/BI_Autopilot_Apresentacao_REVISADA.pptx)
- [Passo a passo](materiais/Passo_a_Passo_BI_Autopilot_REVISADO.docx)

O gabarito do professor é distribuído separadamente.

## 💻 Versão VS Code (como na empresa)

Na pasta [`vscode/`](vscode):
```
cd vscode
pip install -r requirements.txt
python 1_tratar.py      # limpa a planilha com pandas
python 2_modelar.py     # relacionamentos e medidas DAX
python 3_painel.py      # visuais do painel
```
Depois abra `projeto\Obras.pbip` no Power BI Desktop e clique em **Atualizar**.
Com o GitHub Copilot, use `/sugerir-medidas` e `/criar-painel` no chat.

## 📁 O que tem aqui

| Pasta | Conteúdo |
|---|---|
| `kit/` | `autopilot.py` (o motor) e `obras_bruto.xlsx` (a planilha suja). O notebook baixa daqui. |
| `notebooks/` | **Monte seu Tutor e seu Motor** — a aula no Colab |
| `vscode/` | O mesmo fluxo em 3 arquivos Python, para o VS Code |

O `autopilot.py` escreve o projeto no formato PBIP da Microsoft (TMDL para o modelo e PBIR para o painel).

## Contrato e validação

- Dinheiro: números finitos ou texto brasileiro como `R$ 1.234,56` e `-R$ 12,30`.
  Textos estrangeiros como `1234.56`, vazios e ausentes geram erro, sem conversão silenciosa para zero.
- Previsto e realizado devem bater com os totais do sistema antes de modelar.
- Reexecutar uma medida substitui sua definição. Peças reprovadas perdem a aprovação anterior.
- `Painel(modelo).baixar()` exporta com staging e preserva outros projetos. `Painel(nome)` usa `salvar()`.
- O timeout limita a espera, não cancela à força uma chamada já iniciada.

Na raiz, execute `python -m unittest discover -s tests -v`.
O GitHub Actions também executa os testes e as três etapas VS Code.
A abertura e atualização no Power BI Desktop continuam sendo a conferência final do aluno.

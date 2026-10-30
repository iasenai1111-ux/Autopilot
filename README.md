# 🤖 BI Autopilot — da planilha suja ao painel do Power BI

> **O Python calcula. A IA explica. Você confere.**

Material do curso **Power BI e Análise de Dados com IA — SENAI**.
O aluno monta um tutor de IA (Gemini), escreve as peças de um motor de limpeza de dados e gera um
**projeto oficial do Power BI (PBIP)** a partir da planilha suja de uma construtora.

## ▶ Começar pelo Colab (nada para instalar)

[![Abrir no Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/iasenai1111-ux/Autopilot/blob/main/notebooks/Monte_seu_Tutor_e_seu_Motor.ipynb)

1. Abra o notebook pelo botão acima e faça **Arquivo → Salvar uma cópia no Drive**.
2. Siga as partes 0 a 6: imagem explicando → lacunas `______` para completar → resultado esperado.
3. No fim, o Colab baixa o **Obras.zip**. Extraia em qualquer pasta, abra `projeto\Obras.pbip`
   no Power BI Desktop e clique em **Atualizar** (os dados já vão dentro do projeto).

O tutor tem **3 opções de IA**, e se uma lotar ou demorar, ele usa a próxima:

| Plano | Chave | Destaque |
|---|---|---|
| A · Colab AI | não precisa | já vem no Colab |
| B · Groq | grátis em console.groq.com (`GROQ_API_KEY`) | muito rápido, não treina com seus dados |
| C · Gemini | grátis em aistudio.google.com (`GEMINI_API_KEY`) | opcional — pode lotar (503) e usar as perguntas no plano grátis |

As chaves ficam nos **Segredos 🔑** do Colab, nunca no código.

## 💻 Versão VS Code (como na empresa)

Na pasta [`vscode/`](vscode):
```
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

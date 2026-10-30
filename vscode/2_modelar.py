# =============================================================
# ETAPA 2 — MODELAR: relacionamentos e medidas
# Relacionamento = como as tabelas conversam (a obra do lançamento é a obra do cadastro).
# Medida = uma conta em DAX que o painel vai mostrar.
#
# Rodar: python 2_modelar.py
# =============================================================
import autopilot as ap

modelo = ap.Modelo("Obras")   # lê as abas de dados/base_tratada.xlsx

# ---------- 1. RELACIONAMENTOS ----------
# de (lado MUITOS, a tabela de fatos)  ->  para (lado UM, o cadastro)
modelo.relacionar("Lancamentos[Cod_Obra]", "Obras[Cod_Obra]")
modelo.relacionar("Lancamentos[Data]", "Calendario[Data]")

# ---------- 2. MEDIDAS ----------
# Um dicionário: "Nome da medida": "conta em DAX"
# Formato automático: nome com % vira porcentagem, "Qtd ..." vira número, o resto vira R$
# Atenção: medida NÃO pode ter o mesmo nome de uma coluna (por isso "Total Previsto").
medidas = {
    "Total Previsto": "SUM(Lancamentos[Previsto])",
    "Total Realizado": "SUM(Lancamentos[Realizado])",
    "Desvio": "[Total Realizado] - [Total Previsto]",
    "Desvio %": "DIVIDE([Desvio], [Total Previsto])",
    "Qtd Obras": "DISTINCTCOUNT(Lancamentos[Cod_Obra])",
    "Orçamento": "SUM(Obras[Orcamento])",
    "% do Orçamento Usado": "DIVIDE([Total Realizado], [Orçamento])",
}
modelo.medidas(medidas, tabela="Lancamentos")

# ---------- 3. PEDIR SUGESTÕES À IA (opcional) ----------
# Tire o # da linha abaixo, rode, copie o pedido e cole no Copilot ou no Gemini.
# Cole as medidas sugeridas no dicionário acima. Depois VOCÊ confere no Power BI.
# modelo.pedido_para_ia()

# ---------- 4. SALVAR ----------
modelo.salvar()

# =============================================================
# ETAPA 1 — TRATAR A PLANILHA
# A planilha que sai do sistema vem "suja": título em cima, linha vazia,
# linha duplicada, dinheiro escrito como texto, categoria escrita de 3 jeitos.
# Aqui você limpa tudo isso. O Power BI só vai receber dados limpos.
#
# Rodar: python 1_tratar.py
# =============================================================
import pandas as pd
import autopilot as ap

ARQUIVO = "dados/obras_bruto.xlsx"

# ---------- 1. LER ----------
# skiprows = quantas linhas de título existem ANTES do cabeçalho
obras = pd.read_excel(ARQUIVO, sheet_name="Obras", skiprows=1)
lanc = pd.read_excel(ARQUIVO, sheet_name="Lançamentos", skiprows=3)

# ---------- 2. RENOMEAR ----------
# Um dicionário: "nome antigo": "nome novo" (curto, sem acento, sem espaço)
obras = obras.rename(columns={
    "Cód. Obra": "Cod_Obra",
    "Nome da Obra": "Obra",
    "Responsável": "Responsavel",
    "Início": "Inicio",
    "Orçamento Total": "Orcamento",
})
lanc = lanc.rename(columns={
    "Obra": "Cod_Obra",
    "Descrição": "Descricao",
    "Valor Previsto": "Previsto",
    "Valor Realizado": "Realizado",
})

# ---------- 3. TIRAR LINHAS VAZIAS E A LINHA DE TOTAL ----------
obras = obras.dropna(how="all")
lanc = lanc.dropna(how="all")
total_do_sistema = lanc[lanc["Data"] == "TOTAL GERAL"]   # guardamos para conferir no fim
lanc = lanc[lanc["Data"] != "TOTAL GERAL"]


# ---------- 4. DINHEIRO: TEXTO -> NÚMERO ----------
def para_numero(valor):
    """'R$ 1.234,56' vira 1234.56. Se já for número, devolve igual."""
    valor = ap.validar_numero_br(valor)
    if isinstance(valor, str):
        valor = valor.replace("R$", "").replace(".", "").replace(",", ".").strip()
    return float(valor)


obras["Orcamento"] = obras["Orcamento"].apply(para_numero)
lanc["Previsto"] = lanc["Previsto"].apply(para_numero)
lanc["Realizado"] = lanc["Realizado"].apply(para_numero)

# ---------- 5. DATAS ----------
# dayfirst=True: no Brasil o dia vem primeiro (05/03 = 5 de março)
obras["Inicio"] = pd.to_datetime(obras["Inicio"], dayfirst=True, format="mixed")
lanc["Data"] = pd.to_datetime(lanc["Data"], dayfirst=True, format="mixed")

# ---------- 6. PADRONIZAR TEXTOS ----------
obras["Cod_Obra"] = obras["Cod_Obra"].str.strip().str.upper()    # " ob-03" -> "OB-03"
obras["Obra"] = obras["Obra"].str.strip()
obras["Cidade"] = obras["Cidade"].str.strip().str.title()         # "são paulo " -> "São Paulo"
lanc["Cod_Obra"] = lanc["Cod_Obra"].str.strip().str.upper()

padrao_categoria = {                                              # jeito escrito -> jeito certo
    "material": "Material",
    "mão de obra": "Mão de Obra",
    "equipamentos": "Equipamentos",
    "serviços terceiros": "Serviços Terceiros",
}
lanc["Categoria"] = lanc["Categoria"].str.strip().str.lower().map(padrao_categoria)

# ---------- 7. TIRAR DUPLICADAS ----------
# Só depois de padronizar! Antes disso " ob-01" e "OB-01" pareciam diferentes.
obras = obras.drop_duplicates()
lanc = lanc.drop_duplicates()

# ---------- 8. VOCÊ CONFERE ----------
total_python = lanc["Previsto"].sum()
total_sistema = para_numero(total_do_sistema["Previsto"].iloc[0])
print(f"Total previsto — Python: {total_python:,.2f} | Sistema: {total_sistema:,.2f}")
print("Bateu! ✅" if round(total_python, 2) == round(total_sistema, 2) else "NÃO bateu ❌ — sobrou sujeira")
if round(total_python, 2) != round(total_sistema, 2):
    raise ValueError("Conferência falhou: a base não será salva.")
print("Categorias:", sorted(lanc["Categoria"].dropna().unique()))
print("Categorias que não foram reconhecidas:", lanc["Categoria"].isna().sum())
if lanc["Categoria"].isna().any():
    raise ValueError("Há categorias desconhecidas. Corrija o mapeamento antes de salvar.")
total_realizado_sistema = para_numero(total_do_sistema["Realizado"].iloc[0])
if round(lanc["Realizado"].sum(), 2) != round(total_realizado_sistema, 2):
    raise ValueError("O total realizado diverge do sistema. A base não será salva.")
print("Total realizado também conferido ✅")

# ---------- 9. SALVAR ----------
calendario = ap.calendario(lanc["Data"])
ap.salvar_base({"Obras": obras, "Lancamentos": lanc, "Calendario": calendario})

# =============================================================
# ETAPA 3 — PAINEL
# Cada linha é um visual. O motor organiza tudo na página sozinho.
# Coluna se escreve Tabela[Coluna]. Medida se escreve só o nome.
#
# Rodar: python 3_painel.py   (depois abra projeto/Obras.pbip no Power BI Desktop)
# =============================================================
import autopilot as ap

painel = ap.Painel("Obras", titulo="Controle de Custos das Obras")

# ---------- FILTROS (no topo, à direita) ----------
painel.segmentacao("Obras[Cidade]")
painel.segmentacao("Calendario[Mes]")

# ---------- CARTÕES (os números grandes) ----------
painel.cartao("Total Previsto")
painel.cartao("Total Realizado")
painel.cartao("Desvio")
painel.cartao("Desvio %")

# ---------- GRÁFICOS (em grade, 2 por linha) ----------
painel.barras("Obras[Obra]", "Desvio", titulo="Onde estamos perdendo dinheiro?")
painel.linhas("Calendario[AnoMes]", ["Total Previsto", "Total Realizado"], titulo="Previsto x Realizado por mês")
painel.colunas("Lancamentos[Categoria]", "Total Realizado", titulo="Gasto por categoria")
painel.tabela(["Obras[Obra]", "Obras[Responsavel]", "Total Previsto", "Total Realizado", "Desvio %"],
              titulo="Detalhe por obra")

painel.salvar()

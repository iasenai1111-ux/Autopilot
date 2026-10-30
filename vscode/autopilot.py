"""
BI Autopilot — o "motor" que escreve o projeto do Power BI.

Você NÃO precisa editar este arquivo.
Quem decide tudo é você, nos arquivos 1_tratar.py, 2_modelar.py e 3_painel.py.
Este motor só transforma as suas decisões nos arquivos que o Power BI Desktop abre
(formato oficial PBIP da Microsoft: TMDL para o modelo e PBIR para o painel).

    Você decide                 O motor escreve
    -------------------------   ------------------------------------------
    tabelas limpas (pandas)  -> dados/base_tratada.xlsx
    relacionamentos/medidas  -> projeto/<nome>.SemanticModel  (TMDL)
    visuais do painel        -> projeto/<nome>.Report         (PBIR)
"""

import base64
import json
import math
import os
import re
import shutil
import uuid
import zlib
from pathlib import Path

import pandas as pd

PASTA_DADOS = Path("dados")
PASTA_PROJETO = Path("projeto")
BASE_TRATADA = PASTA_DADOS / "base_tratada.xlsx"
PASTA_NO_WINDOWS = "C:\\BI_Autopilot\\dados\\"  # onde os dados ficam se o projeto veio do Colab

CORES = ["#164194", "#E84910", "#5B8DEF", "#F4A27F", "#0E2A5E", "#9AA5B8", "#2E7D32", "#C62828"]
MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
FORMATOS = {"R$": "\\R$ #,##0", "%": "0.0%", "0": "#,##0"}


# =====================================================================
# ETAPA 1 — ajudantes do tratamento
# =====================================================================

def calendario(datas):
    """Cria a tabela Calendario: um dia por linha, do primeiro ao último dia dos dados."""
    datas = pd.to_datetime(datas)
    inicio = datas.min().replace(day=1)
    fim = datas.max() + pd.offsets.MonthEnd(0)
    dias = pd.date_range(inicio, fim, freq="D")
    return pd.DataFrame({
        "Data": dias,
        "Ano": dias.year,
        "MesNum": dias.month,
        "Mes": [MESES[m - 1] for m in dias.month],
        "AnoMes": dias.strftime("%Y-%m"),
    })


def salvar_base(tabelas, arquivo=BASE_TRATADA):
    """Salva cada tabela limpa como uma aba do Excel (é dali que o Power BI lê)."""
    Path(arquivo).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(arquivo, engine="openpyxl") as excel:
        for nome, tabela in tabelas.items():
            tabela.to_excel(excel, sheet_name=nome, index=False)
    print(f"✅ Base tratada salva em {arquivo}")
    for nome, tabela in tabelas.items():
        print(f"   {nome}: {len(tabela)} linhas, colunas = {list(tabela.columns)}")


# =====================================================================
# AS PEÇAS DO MOTOR — o aluno escreve, o Python confere
# =====================================================================

PECAS = {}          # peças aprovadas no teste
LACUNA = "🟡 Ainda tem lacuna ______ nesta célula. Troque cada ______ pelo que a explicação pede e rode de novo."


class Lacuna(Exception):
    pass


def _sem_lacuna(*textos):
    for texto in textos:
        if "___" in str(texto):
            raise Lacuna(LACUNA)


# ---------- erros em português (só no Colab/Jupyter) ----------
_DICAS_ERRO = {
    "NameError": "O Python não conhece esse nome. Você rodou as células de cima? Escreveu igualzinho (maiúscula conta)?",
    "ImportError": "Essa ferramenta não existe com esse nome. Confira a escrita do import.",
    "ModuleNotFoundError": "Essa ferramenta não existe com esse nome. Confira a escrita do import.",
    "AttributeError": "Depois do ponto veio um nome que não existe. Confira a escrita (maiúscula conta).",
    "KeyError": "Esse nome de coluna não existe. Copie o nome exato que apareceu na saída.",
    "TypeError": "Algum valor está no tipo errado (texto no lugar de número, ou faltou algo entre os parênteses).",
    "SecretNotFoundError": "O cofre não tem um segredo com esse nome. Abra a chave 🔑 e confira o Name (e o Notebook access).",
    "NotebookAccessError": "Ative o Notebook access do segredo na chave 🔑 do menu da esquerda.",
}


def ajuda(silencioso=False):
    """Liga as mensagens de erro em português. Fica ligado até reiniciar o Colab."""
    try:
        shell = get_ipython()  # noqa: F821  (existe dentro do Colab/Jupyter)
    except NameError:
        return

    def tratar(shell, tipo, valor, tb, tb_offset=None):
        celula = (shell.user_ns.get("In") or [""])[-1]
        if isinstance(valor, Lacuna) or "___" in str(valor) or "___" in celula:
            print(LACUNA)
            return None
        print(f"🔴 {tipo.__name__}: {valor}")
        texto = str(valor)
        if tipo.__name__ == "SecretNotFoundError" and len(texto) > 60 and "_API_KEY" not in texto:
            print("💡 Parece que você colou a CHAVE no lugar do NOME. Na célula vai só o nome (ex.: \"GROQ_API_KEY\").")
            print("   A chave fica no cofre 🔑, no campo Value. E troque essa chave por uma nova no site do fornecedor.")
        elif tipo.__name__ == "SecretNotFoundError":
            print("💡 O cofre não tem um segredo com esse nome. Abra a chave 🔑 e confira o Name, com _ (sublinhado), e o Notebook access.")
        elif "401" in texto or "invalid_api_key" in texto.lower() or "API key not valid" in texto or "PERMISSION_DENIED" in texto:
            print("💡 A chave não foi aceita. Crie uma nova no site do fornecedor e cole no cofre 🔑 (campo Value).")
        elif "503" in texto or "UNAVAILABLE" in texto or "high demand" in texto or "overloaded" in texto.lower():
            print("💡 Esse fornecedor de IA está sobrecarregado agora (não é erro seu). Use o tutor() — ele tenta o próximo plano.")
        elif "timed out" in texto.lower() or "timeout" in texto.lower():
            print("💡 A IA demorou demais para responder. Use o tutor() — ele tenta o próximo plano.")
        elif "429" in texto or "RESOURCE_EXHAUSTED" in texto or "rate limit" in texto.lower():
            print("💡 Limite de perguntas por minuto. Espere 1 minuto, ou use o tutor() para tentar outro plano.")
        elif "404" in texto and "model" in texto.lower():
            print("💡 Esse MODELO não existe mais. Peça ao professor o nome novo e troque na variável do modelo.")
        elif tipo.__name__ == "NameError" and "PLANOS" in celula and "'tutor_" in texto:
            print("💡 Na lista PLANOS, deixe só os planos que você montou (ex.: [tutor_colab, tutor_groq]).")
        elif tipo.__name__ == "NameError" and any(f"'{n}'" in texto for n in ("cliente", "tutor", "genai", "userdata", "chave", "REGRAS",
                                                                               "tutor_colab", "tutor_groq", "tutor_gemini", "OpenAI", "ai")):
            print("💡 O tutor ainda não está montado. Volte na Parte 1 e complete os Passos 1 a 5, na ordem.")
        elif tipo.__name__ == "NameError" and any(f"'{n}'" in texto for n in ("obras", "lanc", "modelo", "painel")):
            print("💡 Falta rodar uma célula de cima (ler, limpar ou modelar). Rode as células anteriores, na ordem.")
        else:
            print("💡 " + _DICAS_ERRO.get(tipo.__name__, "Leia a mensagem acima e confira a linha que você mudou."))
        return None

    shell.set_custom_exc((Exception,), tratar)
    if not silencioso:
        print("🛟 Ajuda ligada: erros vão aparecer explicados em português.")


def mostrar_chave(chave):
    """Confere a chave sem mostrar ela inteira."""
    _sem_lacuna(chave)
    if not chave or len(str(chave)) < 20:
        print("❌ A chave veio vazia ou curta demais. Confira o Value no cofre 🔑."); return
    print(f"✅ Chave encontrada: {chave[:4]}{'•' * 12}{chave[-3:]}  (nunca mostre a chave inteira)")
    return


_RECEITAS = {
    "tutor_colab": (("generate_text", "REGRAS"), "Plano A (Colab AI)"),
    "tutor_groq": (("chat", "completions", "REGRAS", "content", "MODELO_GROQ"), "Plano B (Groq)"),
    "tutor_gemini": (("generate_content", "REGRAS", "text", "MODELO_GEMINI"), "Plano C (Gemini)"),
    "tutor": (("PLANOS",), "O tutor que nunca para"),
}


def checar_tutor(funcao):
    """Confere uma receita de tutor antes de usar (sem gastar pergunta da IA)."""
    codigo = funcao.__code__
    nomes = set(codigo.co_names) | set(codigo.co_varnames)
    for interno in codigo.co_consts:          # funções dentro da função
        if hasattr(interno, "co_names"):
            nomes |= set(interno.co_names)
    if any("___" in n for n in nomes) or any("___" in str(c) for c in codigo.co_consts):
        print(LACUNA); return
    precisa, nome_bonito = _RECEITAS.get(funcao.__name__, ((), "Receita"))
    faltando = [n for n in precisa if n not in nomes]
    if faltando:
        print(f"❌ A receita {funcao.__name__} ainda não usa: {faltando}"); return
    print(f"✅ {nome_bonito} pronto. Teste com: {funcao.__name__}(\"sua pergunta\")")


def com_limite(plano, pergunta, segundos=30):
    """Roda um plano de IA com tempo máximo. Demorou demais? Dá erro, e o tutor() tenta o próximo."""
    import concurrent.futures
    executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
    tarefa = executor.submit(plano, pergunta)
    try:
        return tarefa.result(timeout=segundos)
    except concurrent.futures.TimeoutError:
        raise TimeoutError(f"demorou mais de {segundos} segundos") from None
    finally:
        executor.shutdown(wait=False)


def motivo(erro):
    """Explica em poucas palavras por que um plano de IA falhou."""
    texto, tipo = str(erro), type(erro).__name__
    if "___" in texto:
        return "ainda tem lacuna ______"
    if tipo == "TimeoutError" or "timed out" in texto.lower():
        return "demorou demais"
    if "503" in texto or "UNAVAILABLE" in texto or "overloaded" in texto.lower() or "high demand" in texto:
        return "está lotado agora"
    if "429" in texto or "RESOURCE_EXHAUSTED" in texto or "rate limit" in texto.lower():
        return "limite de perguntas por minuto"
    if "401" in texto or "invalid_api_key" in texto.lower() or "API key not valid" in texto:
        return "a chave foi recusada"
    if tipo == "SecretNotFoundError":
        return "a chave não está no cofre"
    if tipo == "NameError":
        return "esse plano ainda não foi montado"
    return f"{tipo}: {texto[:80]}"


def checar_regras(regras):
    """Confere se as regras do tutor estão completas."""
    _sem_lacuna(regras)
    faltando = [p for p in ("NUNCA", "passos") if p.lower() not in regras.lower()]
    if faltando:
        print(f"❌ Faltou nas regras: {faltando}. Sem isso, a IA pode entregar a resposta pronta."); return
    print(f"✅ Regras prontas ({len(regras.splitlines())} linhas). A IA vai obedecer a elas em toda pergunta.")
    return

_CASOS = {
    "para_numero": [("R$ 1.234,56", 1234.56), ("R$ 950.000,00", 950000.0), ("R$ 12,30", 12.3), (1500.5, 1500.5)],
    "data_br": [("05/03/2026", "2026-03-05"), ("28/09/2026", "2026-09-28"), (pd.Timestamp("2026-01-10"), "2026-01-10")],
    "padronizar": [(" ob-01", "OB-01"), ("OB-02 ", "OB-02"), ("ob-03", "OB-03")],
    "conferir_total": [(([10.5, 20.25], 30.75), True), (([10.4, 20.0], 30.0), False), (([0.1, 0.2], 0.3), True)],
}
_DICAS = {
    "para_numero": "No Brasil o ponto separa o milhar e a vírgula separa os centavos. No Python é o contrário.",
    "data_br": "Em 05/03/2026 o dia vem primeiro. Procure o parâmetro que avisa isso ao pandas.",
    "padronizar": "Dois passos: tirar os espaços das pontas e deixar tudo em maiúsculas.",
    "conferir_total": "0.1 + 0.2 não dá exatamente 0.3 no computador. Arredonde para 2 casas antes de comparar.",
}


def testar(funcao):
    """Confere a sua peça com casos de teste. Se passar, ela é encaixada no motor."""
    nome = funcao.__name__
    codigo = funcao.__code__
    if any("___" in str(x) for x in (*codigo.co_consts, *codigo.co_names)):
        print(LACUNA); return
    if nome in ("ler", "limpar"):
        return _testar_motor(nome, funcao)
    if nome not in _CASOS:
        print(f"Não conheço a peça '{nome}'. Peças: {list(_CASOS)}"); return
    for entrada, esperado in _CASOS[nome]:
        argumentos = entrada if nome == "conferir_total" else (entrada,)
        try:
            saida = funcao(*argumentos)
            if nome == "data_br":
                saida = pd.Timestamp(saida).strftime("%Y-%m-%d")
            certo = abs(saida - esperado) < 1e-9 if nome == "para_numero" else saida == esperado
        except Exception as erro:
            if "___" in str(erro):
                print(LACUNA); return
            print(f"❌ {nome}({entrada!r}) deu erro: {erro}")
            print(f"💡 Dica: {_DICAS[nome]}"); return
        if not certo:
            print(f"❌ {nome}({entrada!r}) devolveu {saida!r}, mas o esperado era {esperado!r}")
            print(f"💡 Dica: {_DICAS[nome]}"); return
    PECAS[nome] = funcao
    print(f"✅ Peça '{nome}' aprovada e encaixada no motor ({len(PECAS)}/{len(_CASOS)})")
    return


def _testar_motor(nome, funcao):
    """Testa as funções do motor que o aluno montou."""
    import builtins
    if nome == "ler":
        try:
            tabela = funcao("obras_bruto.xlsx", "Obras", 1)
        except Exception as erro:
            print(LACUNA if "___" in str(erro) else f"❌ ler() deu erro: {erro}"); return
        if not isinstance(tabela, pd.DataFrame):
            print("❌ ler() não devolveu a tabela. Faltou o return?"); return
        if "Cód. Obra" not in tabela.columns:
            print("❌ As colunas vieram erradas. Confira se o skiprows recebe o 'pular'."); return
        print("✅ ler() aprovada: lê a aba certa e pula o título")
        return
    faltam = [p for p in ("para_numero", "data_br", "padronizar") if p not in PECAS]
    if faltam:
        print(f"❌ Antes do limpar(), aprove as peças: {faltam}"); return
    suja = pd.DataFrame({
        "Cod": [" ob-01", "OB-01 ", None, "ob-02", "TOTAL"],
        "Valor": ["R$ 1.000,50", "R$ 1.000,50", None, "R$ 20,00", "R$ 1.020,50"],
        "Dia": ["05/03/2026", "05/03/2026", None, "06/03/2026", None],
        "Cidade": ["são paulo", "são paulo", None, "CAMPINAS", None],
    })
    print("🧪 Testando com uma tabela suja de 5 linhas (1 vazia, 1 duplicada, 1 TOTAL)...")
    falar = builtins.print
    try:
        builtins.print = lambda *a, **k: None   # silencia os prints do aluno durante o teste
        limpa = funcao(suja.copy(), codigo="Cod", dinheiro=["Valor"], datas=["Dia"], textos=["Cidade"], tirar_linha="TOTAL")
    except Exception as erro:
        builtins.print = falar
        print(LACUNA if "___" in str(erro) else f"❌ limpar() deu erro: {erro}"); return
    finally:
        builtins.print = falar
    if not isinstance(limpa, pd.DataFrame):
        print("❌ limpar() não devolveu a tabela. Faltou o return?"); return
    limpa = limpa.reset_index(drop=True)
    checagens = [
        (len(limpa) == 2, f"Sobraram {len(limpa)} linhas, o certo é 2. Confira os passos 1 (vazias), 2 (TOTAL) e 7 (duplicadas)."),
        (list(limpa["Cod"]) == ["OB-01", "OB-02"], "O código não ficou padronizado. Confira o passo 5."),
        (list(pd.to_numeric(limpa["Valor"], errors="coerce").round(2)) == [1000.5, 20.0], "O dinheiro não virou número. Confira o passo 3."),
        (list(pd.to_datetime(limpa["Dia"]).dt.strftime("%Y-%m-%d")) == ["2026-03-05", "2026-03-06"], "As datas não ficaram certas. Confira o passo 4."),
        (list(limpa["Cidade"]) == ["São Paulo", "Campinas"], "Os textos não ficaram com a primeira letra maiúscula. Confira o passo 6."),
    ]
    for ok, dica in checagens:
        if not ok:
            print(f"❌ {dica}"); return
    print("✅ limpar() aprovada: tira vazias, TOTAL e duplicadas e usa as suas 3 peças")


def _peca(nome):
    if nome not in PECAS:
        raise RuntimeError(f"Falta a peça '{nome}'. Volte, escreva a função e rode ap.testar({nome}).")
    return PECAS[nome]


def ler(arquivo, aba, pular=0):
    """Lê uma aba do Excel pulando as linhas de título."""
    _sem_lacuna(arquivo, aba)
    df = pd.read_excel(arquivo, sheet_name=aba, skiprows=pular)
    print(f"📄 {aba}: {len(df)} linhas")
    print(f"   Colunas: {list(df.columns)}")
    with pd.option_context("display.width", 160, "display.max_columns", 20):
        print("   Começo:"); print(df.head(3).to_string())
        print("   Fim:"); print(df.tail(2).to_string())
    return df


_TOTAIS = {}


def limpar(df, codigo=None, dinheiro=(), datas=(), textos=(), tirar_linha=None):
    """Usa as SUAS peças para limpar a tabela e conta o que foi feito."""
    pedidas = ([codigo] if codigo else []) + list(dinheiro) + list(datas) + list(textos)
    _sem_lacuna(*pedidas, tirar_linha or "")
    for coluna in pedidas:
        if coluna not in df.columns:
            raise KeyError(f"A coluna '{coluna}' não existe. Colunas desta tabela: {list(df.columns)}")
    antes = len(df)
    df = df.dropna(how="all").copy()
    vazias = antes - len(df)
    if tirar_linha:
        marcada = df.apply(lambda linha: linha.astype(str).str.strip().eq(tirar_linha).any(), axis=1)
        for coluna in dinheiro:
            if marcada.any():
                _TOTAIS[coluna] = _peca("para_numero")(df.loc[marcada, coluna].iloc[0])
        df = df[~marcada]
    for coluna in dinheiro:
        df[coluna] = df[coluna].apply(_peca("para_numero"))
    for coluna in datas:
        df[coluna] = pd.to_datetime(df[coluna].apply(_peca("data_br")))
    if codigo:
        df[codigo] = df[codigo].apply(_peca("padronizar"))
    for coluna in textos:
        df[coluna] = df[coluna].astype(str).str.strip().str.title().str.replace(
            r" (De|Da|Do|Das|Dos|E) ", lambda m: " " + m.group(1).lower() + " ", regex=True)
    sem_duplicadas = df.drop_duplicates()
    extra = f" | linha '{tirar_linha}' tirada" if tirar_linha else ""
    print(f"🧹 {len(sem_duplicadas)} linhas limpas | vazias tiradas: {vazias} | "
          f"duplicadas tiradas: {len(df) - len(sem_duplicadas)}{extra}")
    return sem_duplicadas.reset_index(drop=True)


def conferir(df, coluna):
    """Compara a soma limpa com o total que o sistema mostrou (a regra de ouro: você confere)."""
    _sem_lacuna(coluna)
    if coluna not in _TOTAIS:
        print(f"Não guardei o total do sistema para '{coluna}'. Use tirar_linha= no limpar()."); return
    ok = _peca("conferir_total")(df[coluna].tolist(), _TOTAIS[coluna])
    print(f"Soma limpa: {df[coluna].sum():,.2f} | Total do sistema: {_TOTAIS[coluna]:,.2f} -> "
          + ("bateu ✅" if ok else "NÃO bateu ❌ (sobrou sujeira?)"))
    return


# =====================================================================
# ETAPA 2 — o modelo (tabelas, relacionamentos e medidas) em TMDL
# =====================================================================

def _nome(texto):
    """Nomes com espaço, acento ou símbolo vão entre aspas simples no TMDL."""
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", texto):
        return texto
    return "'" + texto.replace("'", "''") + "'"


def _separar(campo):
    """'Obras[Cidade]' -> ('Obras', 'Cidade')"""
    achou = re.fullmatch(r"\s*'?([^'\[\]]+?)'?\[([^\]]+)\]\s*", campo)
    if not achou:
        raise ValueError(f"Escreva o campo no formato Tabela[Coluna]. Recebi: {campo!r}")
    return achou.group(1), achou.group(2)


def _tipo(serie):
    """Tipo do pandas -> (tipo no Power BI, tipo no Power Query)"""
    if pd.api.types.is_bool_dtype(serie):
        return "boolean", "type logical"
    if pd.api.types.is_integer_dtype(serie):
        return "int64", "Int64.Type"
    if pd.api.types.is_numeric_dtype(serie):
        return "double", "type number"
    if pd.api.types.is_datetime64_any_dtype(serie):
        return "dateTime", "type date"
    return "string", "type text"


def _escrever(caminho, texto):
    caminho = Path(caminho)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    caminho.write_text(texto, encoding="utf-8")  # UTF-8 sem BOM, como o Power BI pede


def _json(caminho, dados):
    _escrever(caminho, json.dumps(dados, ensure_ascii=False, indent=2))


def _id(*partes):
    """Identificador fixo (o mesmo a cada execução) para o Git não acusar mudança à toa."""
    return str(uuid.uuid5(uuid.NAMESPACE_URL, "bi-autopilot/" + "/".join(partes)))


class Modelo:
    """O modelo semântico: tabelas + relacionamentos + medidas."""

    def __init__(self, nome, base=BASE_TRATADA, **tabelas):
        """Modelo("Obras", obras=obras, lancamentos=lanc)  ou  Modelo("Obras") para ler dados/base_tratada.xlsx"""
        self.nome = nome
        if tabelas:
            self.tabelas = {k[:1].upper() + k[1:]: df.reset_index(drop=True) for k, df in tabelas.items()}
            for df in self.tabelas.values():   # datas que vieram como "objeto" viram data de verdade
                for coluna in df.columns:
                    valores = df[coluna].dropna()
                    if df[coluna].dtype == object and len(valores) and valores.map(lambda v: isinstance(v, pd.Timestamp)).all():
                        df[coluna] = pd.to_datetime(df[coluna])
        else:
            if not Path(base).exists():
                raise FileNotFoundError(f"Não achei {base}. Rode antes o 1_tratar.py.")
            self.tabelas = pd.read_excel(base, sheet_name=None)
        self.relacoes = []
        self.lista_medidas = []  # (tabela, nome, dax, formato)
        print(f"📦 Modelo '{nome}' com as tabelas: {', '.join(self.tabelas)}")

    # ---------- relacionamentos ----------
    def relacionar(self, de, para):
        """de = lado MUITOS (fato), para = lado UM (dimensão). Ex.: 'Lancamentos[Cod_Obra]' -> 'Obras[Cod_Obra]'"""
        _sem_lacuna(de, para)
        t1, c1 = _separar(de)
        t2, c2 = _separar(para)
        for t, c in ((t1, c1), (t2, c2)):
            if t not in self.tabelas or c not in self.tabelas[t].columns:
                raise ValueError(f"A coluna {t}[{c}] não existe. Confira o nome no 1_tratar.py.")
        if self.tabelas[t2][c2].duplicated().any():
            if not self.tabelas[t1][c1].duplicated().any():
                raise ValueError(f"Parece que a ordem está invertida. Escreva primeiro o lado MUITOS: "
                                 f"modelo.relacionar('{t2}[{c2}]', '{t1}[{c1}]')")
            repetidos = self.tabelas[t2][c2][self.tabelas[t2][c2].duplicated()].unique()[:5]
            raise ValueError(f"O lado UM ({t2}[{c2}]) tem valores repetidos: {list(repetidos)}. "
                             "Remova as duplicadas no tratamento.")
        sobrando = set(self.tabelas[t1][c1].dropna()) - set(self.tabelas[t2][c2])
        if sobrando:
            print(f"⚠️  {len(sobrando)} valor(es) de {t1}[{c1}] não existem em {t2}[{c2}]: {list(sobrando)[:5]}")
        self.relacoes.append((t1, c1, t2, c2))
        print(f"🔗 {t1}[{c1}]  →  {t2}[{c2}]")

    # ---------- calendário ----------
    def calendario(self, campo):
        """Cria a tabela Calendario a partir de uma coluna de data e já liga as duas."""
        _sem_lacuna(campo)
        t, c = _separar(campo)
        if t not in self.tabelas or c not in self.tabelas[t].columns:
            raise ValueError(f"A coluna {campo} não existe. Confira o nome.")
        self.tabelas["Calendario"] = calendario(self.tabelas[t][c])
        print(f"📅 Calendario criado: {len(self.tabelas['Calendario'])} dias")
        self.relacionar(campo, "Calendario[Data]")

    # ---------- medidas ----------
    def medidas(self, dicionario, tabela=None):
        """dicionario = {"Nome da medida": "DAX"} ou {"Nome": ("DAX", "R$" | "%" | "0")}
        Sem dizer a tabela, as medidas vão para a maior tabela (a de fatos)."""
        if tabela is None:
            tabela = max((t for t in self.tabelas if t != "Calendario"), key=lambda t: len(self.tabelas[t]))
        if tabela not in self.tabelas:
            raise ValueError(f"A tabela {tabela} não existe.")
        for nome, valor in dicionario.items():
            _sem_lacuna(nome, valor)
            dax, formato = valor if isinstance(valor, tuple) else (valor, self._formato_padrao(nome))
            for t, df in self.tabelas.items():
                if nome in df.columns:
                    raise ValueError(f"A medida '{nome}' tem o mesmo nome da coluna {t}[{nome}]. "
                                     f"Use outro nome, por exemplo 'Total {nome}'.")
            for t, c in re.findall(r"'?([^\s\[\]'(),]+)'?\[([^\]]+)\]", dax):
                if t in self.tabelas and c not in self.tabelas[t].columns:
                    raise ValueError(f"Na medida '{nome}': a coluna {t}[{c}] não existe.")
            self.lista_medidas.append((tabela, nome, " ".join(dax.split()), formato))
            print(f"📐 {nome} = {dax}")

    @staticmethod
    def _formato_padrao(nome):
        if "%" in nome:
            return "%"
        if nome.lower().startswith(("qtd", "quantidade", "nº", "num")):
            return "0"
        return "R$"

    def resumo(self):
        """Texto curto com tabelas, colunas e medidas (bom para mandar ao tutor)."""
        linhas = [f"{t}: {', '.join(df.columns)}" for t, df in self.tabelas.items()]
        linhas.append("Medidas: " + (", ".join(f"{n} = {d}" for _, n, d, _ in self.lista_medidas) or "nenhuma"))
        return "\n".join(linhas)

    def pedido_para_ia(self):
        """Imprime um pedido pronto para colar no Copilot, Gemini ou ChatGPT."""
        linhas = [f"- {t}: {', '.join(df.columns)}" for t, df in self.tabelas.items()]
        ja_tem = ", ".join(n for _, n, _, _ in self.lista_medidas) or "nenhuma"
        print("\n".join([
            "----- COPIE A PARTIR DAQUI -----",
            f"Sou analista de dados em um projeto de Power BI chamado '{self.nome}'.",
            "Tabelas e colunas do modelo:",
            *linhas,
            f"Medidas que já existem: {ja_tem}.",
            "Sugira 5 medidas DAX úteis para a gestão deste negócio.",
            "Responda SOMENTE em linhas de dicionário Python, assim:",
            '    "Nome da medida": "EXPRESSÃO DAX",',
            "Use apenas as tabelas e colunas acima. Explique cada medida em um comentário # curto.",
            "----- ATÉ AQUI -----",
        ]))

    # ---------- escrever os arquivos ----------
    def salvar(self, dados_embutidos=None):
        """dados_embutidos=False: o Power BI lê dados/base_tratada.xlsx (como na empresa).
        dados_embutidos=True: os dados vão DENTRO do projeto (igual ao "Inserir dados"), abre em qualquer pasta.
        Sem dizer nada: no Windows (VS Code) lê o arquivo; no Colab embute."""
        self.embutir = (os.name != "nt") if dados_embutidos is None else dados_embutidos
        pasta = str(PASTA_DADOS.resolve()) + "\\" if os.name == "nt" else PASTA_NO_WINDOWS
        raiz = PASTA_PROJETO / f"{self.nome}.SemanticModel"
        if (raiz / "definition").exists():
            shutil.rmtree(raiz / "definition")

        _json(raiz / "definition.pbism", {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
            "version": "4.2", "settings": {}})
        _json(raiz / ".platform", _plataforma("SemanticModel", self.nome))
        _escrever(raiz / "definition/database.tmdl", f"database {_nome(self.nome)}\n\tcompatibilityLevel: 1600\n")
        _escrever(raiz / "definition/model.tmdl", "\n".join([
            "model Model",
            "\tculture: pt-BR",
            "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
            "\tdiscourageImplicitMeasures",
            "\tsourceQueryCulture: pt-BR",
            "",
            "annotation __PBI_TimeIntelligenceEnabled = 0",
            "",
            *[f"ref table {_nome(t)}" for t in self.tabelas],
            "",
        ]))
        if not self.embutir:
            _escrever(raiz / "definition/expressions.tmdl",
                      f'expression PastaDados = "{pasta}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n'
                      "\n\tannotation PBI_ResultType = Text\n")
        for tabela, df in self.tabelas.items():
            _escrever(raiz / f"definition/tables/{tabela}.tmdl", self._tabela_tmdl(tabela, df))
        if self.relacoes:
            _escrever(raiz / "definition/relationships.tmdl", "\n".join(
                f"relationship {_id(self.nome, t1, c1, t2, c2)}\n"
                f"\tfromColumn: {_nome(t1)}.{_nome(c1)}\n\ttoColumn: {_nome(t2)}.{_nome(c2)}\n"
                for t1, c1, t2, c2 in self.relacoes))

        # resumo que o 3_painel.py usa para achar cada campo
        _json(PASTA_PROJETO / "modelo.json", {
            "nome": self.nome,
            "tabelas": {t: list(df.columns) for t, df in self.tabelas.items()},
            "medidas": {n: t for t, n, _, _ in self.lista_medidas},
        })
        _escrever(PASTA_PROJETO / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n")
        if not (PASTA_PROJETO / f"{self.nome}.Report").exists():
            Painel(self.nome).salvar(silencioso=True)  # página em branco, para já abrir no Power BI
        print(f"\n✅ Modelo salvo em {raiz}")
        if self.embutir:
            print("   Dados embutidos no projeto: abre em qualquer pasta, sem configurar caminho.")
        else:
            print(f"   O Power BI vai buscar os dados em: {pasta}")

    def _tabela_tmdl(self, tabela, df):
        eh_calendario = tabela.lower() == "calendario"
        linhas = [f"table {_nome(tabela)}"]
        if eh_calendario:
            linhas.append("\tdataCategory: Time")
        tipos_m = []
        for t, nome, dax, formato in self.lista_medidas:
            if t == tabela:
                linhas += ["", f"\tmeasure {_nome(nome)} = {dax}",
                           f"\t\tformatString: {FORMATOS.get(formato, formato)}"]
        for coluna in df.columns:
            tipo, tipo_m = _tipo(df[coluna])
            tipos_m.append(f'{{"{coluna}", {tipo_m}}}')
            linhas += ["", f"\tcolumn {_nome(coluna)}", f"\t\tdataType: {tipo}"]
            if eh_calendario and coluna == "Data":
                linhas.append("\t\tisKey")
            if tipo == "dateTime":
                linhas.append("\t\tformatString: dd/mm/yyyy")
            linhas += ["\t\tsummarizeBy: none", f"\t\tsourceColumn: {coluna}"]
            if eh_calendario and coluna == "Mes":
                linhas.append("\t\tsortByColumn: MesNum")
        if self.embutir:
            m = _m_embutido(df, tipos_m)
        else:
            m = [
                "let",
                '    Fonte = Excel.Workbook(File.Contents(PastaDados & "base_tratada.xlsx"), null, true),',
                f'    Aba = Fonte{{[Item="{tabela}", Kind="Sheet"]}}[Data],',
                "    Cabecalho = Table.PromoteHeaders(Aba, [PromoteAllScalars=true]),",
                f"    Tipos = Table.TransformColumnTypes(Cabecalho, {{{', '.join(tipos_m)}}})",
                "in",
                "    Tipos",
            ]
        linhas += ["", f"\tpartition {_nome(tabela)} = m", "\t\tmode: import", "\t\tsource ="]
        linhas += ["\t\t\t\t" + l for l in m]
        linhas += ["", "\tannotation PBI_ResultType = Table", ""]
        return "\n".join(linhas)


def _m_embutido(df, tipos_m):
    """Mesmo formato que o Power BI usa no "Inserir dados": a tabela vira texto compactado dentro do projeto."""
    def texto(v):
        if v is None or (isinstance(v, float) and math.isnan(v)) or v is pd.NaT:
            return None
        if isinstance(v, pd.Timestamp):
            return v.strftime("%Y-%m-%d")
        return str(v)
    linhas = [[texto(v) for v in linha] for linha in df.astype(object).itertuples(index=False)]
    compactador = zlib.compressobj(9, zlib.DEFLATED, -15)  # deflate "puro", o que o Power BI espera
    bruto = compactador.compress(json.dumps(linhas, ensure_ascii=False).encode("utf-8")) + compactador.flush()
    b64 = base64.b64encode(bruto).decode("ascii")
    colunas = ", ".join(f'#"{c}" = _t' for c in df.columns)
    return [
        "let",
        f'    Fonte = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("{b64}", BinaryEncoding.Base64), Compression.Deflate)), '
        f"let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [{colunas}]),",
        f'    Tipos = Table.TransformColumnTypes(Fonte, {{{", ".join(tipos_m)}}}, "en-US")',
        "in",
        "    Tipos",
    ]


def _plataforma(tipo, nome):
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": tipo, "displayName": nome},
        "config": {"version": "2.0", "logicalId": _id(tipo, nome)},
    }


# =====================================================================
# ETAPA 3 — o painel (página e visuais) em PBIR
# =====================================================================

SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"
LARGURA, ALTURA, MARGEM, ESPACO = 1280, 720, 24, 16


def _literal(valor):
    return {"expr": {"Literal": {"Value": valor}}}


def _texto(valor):
    return _literal("'" + str(valor).replace("'", "''") + "'")


class Painel:
    """Uma página de painel. Cada método acrescenta um visual."""

    def __init__(self, nome, titulo=None):
        """Painel(modelo, titulo="...")  ou  Painel("Obras") depois do 2_modelar.py"""
        self.modelo = nome if isinstance(nome, Modelo) else None
        self.nome = self.modelo.nome if self.modelo else nome
        self.titulo = titulo
        self.filtros, self.cartoes, self.graficos = [], [], []
        if self.modelo:
            self.resumo = {"tabelas": {t: list(df.columns) for t, df in self.modelo.tabelas.items()},
                           "medidas": {n: t for t, n, _, _ in self.modelo.lista_medidas}}
        else:
            resumo = PASTA_PROJETO / "modelo.json"
            self.resumo = json.loads(resumo.read_text(encoding="utf-8")) if resumo.exists() else None

    # ---------- achar os campos ----------
    def _campo(self, texto):
        """'Realizado' (medida) ou 'Obras[Cidade]' (coluna) -> referência do Power BI"""
        _sem_lacuna(texto)
        if self.resumo is None:
            raise FileNotFoundError("Rode antes o 2_modelar.py.")
        if "[" in texto:
            tabela, coluna = _separar(texto)
            if coluna not in self.resumo["tabelas"].get(tabela, []):
                raise ValueError(f"A coluna {texto} não existe no modelo.")
            ref = {"Column": {"Expression": {"SourceRef": {"Entity": tabela}}, "Property": coluna}}
            return ref, tabela, coluna
        if texto not in self.resumo["medidas"]:
            raise ValueError(f"A medida '{texto}' não existe. Medidas do modelo: {list(self.resumo['medidas'])}")
        tabela = self.resumo["medidas"][texto]
        ref = {"Measure": {"Expression": {"SourceRef": {"Entity": tabela}}, "Property": texto}}
        return ref, tabela, texto

    def _projecao(self, texto, categoria=False):
        ref, tabela, nome = self._campo(texto)
        p = {"field": ref, "queryRef": f"{tabela}.{nome}", "nativeQueryRef": nome}
        if categoria:
            p["active"] = True
        return p

    # ---------- os visuais que o aluno pode pedir ----------
    def segmentacao(self, coluna):
        self.filtros.append(("slicer", {"Values": [coluna]}, None))

    def cartao(self, medida):
        self.cartoes.append(("card", {"Values": [medida]}, medida))

    def barras(self, categoria, valor, titulo=None):
        self.graficos.append(("clusteredBarChart", {"Category": [categoria], "Y": [valor]},
                              titulo or f"{valor} por {_separar(categoria)[1]}"))

    def colunas(self, categoria, valor, titulo=None):
        self.graficos.append(("clusteredColumnChart", {"Category": [categoria], "Y": [valor]},
                              titulo or f"{valor} por {_separar(categoria)[1]}"))

    def linhas(self, eixo, valores, titulo=None):
        valores = [valores] if isinstance(valores, str) else valores
        self.graficos.append(("lineChart", {"Category": [eixo], "Y": valores},
                              titulo or f"{' x '.join(valores)} por {_separar(eixo)[1]}"))

    def pizza(self, categoria, valor, titulo=None):
        self.graficos.append(("donutChart", {"Category": [categoria], "Y": [valor]},
                              titulo or f"{valor} por {_separar(categoria)[1]}"))

    def tabela(self, campos, titulo="Detalhe"):
        self.graficos.append(("tableEx", {"Values": campos}, titulo))

    # ---------- montar os arquivos ----------
    def _visual(self, numero, tipo, papeis, titulo, x, y, w, h):
        estado = {papel: {"projections": [self._projecao(c, categoria=(papel == "Category")) for c in campos]}
                  for papel, campos in papeis.items()}
        visual = {"visualType": tipo, "query": {"queryState": estado}, "drillFilterOtherVisuals": True}
        if tipo in ("clusteredBarChart", "clusteredColumnChart", "donutChart"):
            visual["query"]["sortDefinition"] = {
                "sort": [{"field": self._campo(papeis["Y"][0])[0], "direction": "Descending"}], "isDefaultSort": True}
        if tipo == "lineChart":
            visual["query"]["sortDefinition"] = {
                "sort": [{"field": self._campo(papeis["Category"][0])[0], "direction": "Ascending"}], "isDefaultSort": True}
        if tipo == "slicer":
            visual["objects"] = {"data": [{"properties": {"mode": _texto("Dropdown")}}]}
        if titulo and tipo not in ("slicer", "card"):
            visual["visualContainerObjects"] = {
                "title": [{"properties": {"show": _literal("true"), "text": _texto(titulo)}}]}
        return {
            "$schema": f"{SCHEMA}/visualContainer/2.7.0/schema.json",
            "name": f"visual{numero:02d}",
            "position": {"x": x, "y": y, "z": numero * 1000, "width": w, "height": h, "tabOrder": numero},
            "visual": visual,
        }

    def _caixa_titulo(self, texto):
        return {
            "$schema": f"{SCHEMA}/visualContainer/2.7.0/schema.json",
            "name": "visual00",
            "position": {"x": MARGEM, "y": 12, "z": 0, "width": 640, "height": 52, "tabOrder": 0},
            "visual": {"visualType": "textbox", "objects": {"general": [{"properties": {"paragraphs": [
                {"textRuns": [{"value": texto, "textStyle": {"fontSize": "20pt", "fontWeight": "bold", "color": CORES[0]}}]}
            ]}}]}, "drillFilterOtherVisuals": True},
        }

    def salvar(self, silencioso=False):
        visuais = self._layout()  # monta tudo antes: se um campo estiver errado, nada é gravado
        raiz = PASTA_PROJETO / f"{self.nome}.Report"
        if (raiz / "definition").exists():
            shutil.rmtree(raiz / "definition")
        _json(PASTA_PROJETO / f"{self.nome}.pbip", {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
            "version": "1.0", "artifacts": [{"report": {"path": f"{self.nome}.Report"}}],
            "settings": {"enableAutoRecovery": True}})
        _json(raiz / ".platform", _plataforma("Report", self.nome))
        _json(raiz / "definition.pbir", {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
            "version": "4.0", "datasetReference": {"byPath": {"path": f"../{self.nome}.SemanticModel"}}})
        _json(raiz / "definition/version.json", {"$schema": f"{SCHEMA}/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})
        _json(raiz / "StaticResources/RegisteredResources/tema_bi_autopilot.json", {
            "name": "BI Autopilot", "dataColors": CORES, "background": "#FFFFFF", "foreground": "#14213D",
            "tableAccent": CORES[0], "good": "#2E7D32", "neutral": "#9AA5B8", "bad": "#C62828"})
        _json(raiz / "definition/report.json", {
            "$schema": f"{SCHEMA}/report/3.0.0/schema.json",
            "themeCollection": {"customTheme": {
                "name": "tema_bi_autopilot.json", "type": "RegisteredResources",
                "reportVersionAtImport": {"visual": "2.1.0", "report": "2.1.0", "page": "2.0.0"}}},
            "resourcePackages": [{"name": "RegisteredResources", "type": "RegisteredResources", "items": [
                {"name": "tema_bi_autopilot.json", "path": "tema_bi_autopilot.json", "type": "CustomTheme"}]}],
            "settings": {"useStylableVisualContainerHeader": True, "defaultDrillFilterOtherVisuals": True,
                         "useEnhancedTooltips": True},
        })
        _json(raiz / "definition/pages/pages.json", {
            "$schema": f"{SCHEMA}/pagesMetadata/1.0.0/schema.json",
            "pageOrder": ["pagina1"], "activePageName": "pagina1"})
        _json(raiz / "definition/pages/pagina1/page.json", {
            "$schema": f"{SCHEMA}/page/2.0.0/schema.json",
            "name": "pagina1", "displayName": "Painel", "displayOption": "FitToPage",
            "height": ALTURA, "width": LARGURA,
            "objects": {"background": [{"properties": {
                "color": {"solid": {"color": _texto("#F6F7FB")}}, "transparency": _literal("0D")}}]},
        })
        for v in visuais:
            _json(raiz / f"definition/pages/pagina1/visuals/{v['name']}/visual.json", v)
        if not silencioso:
            total = len(self.filtros) + len(self.cartoes) + len(self.graficos)
            print(f"✅ Painel salvo em {raiz} com {total} visuais")
            print(f"   Abra no Power BI Desktop: {PASTA_PROJETO / (self.nome + '.pbip')}")

    def baixar(self):
        """Grava o modelo e o painel, junta tudo num .zip e (no Colab) baixa."""
        self._layout()  # confere os campos antes de gravar qualquer coisa
        if PASTA_PROJETO.exists():
            shutil.rmtree(PASTA_PROJETO)
        self.modelo.salvar(dados_embutidos=True)
        self.salvar()
        arquivo = shutil.make_archive(self.nome, "zip", ".", PASTA_PROJETO.name)
        print(f"📦 {arquivo}: extraia em qualquer pasta e abra projeto\\{self.nome}.pbip > Atualizar")
        try:
            from google.colab import files
            files.download(arquivo)
        except ImportError:
            pass

    def _layout(self):
        """Organiza sozinho: título e filtros no topo, cartões numa linha, gráficos em grade."""
        visuais, n = [], 1
        if self.titulo:
            visuais.append(self._caixa_titulo(self.titulo))
        largura_filtro = 240
        for i, (tipo, papeis, titulo) in enumerate(self.filtros):
            x = LARGURA - MARGEM - (len(self.filtros) - i) * (largura_filtro + ESPACO) + ESPACO
            visuais.append(self._visual(n, tipo, papeis, titulo, x, 8, largura_filtro, 60)); n += 1
        y = 80
        if self.cartoes:
            w = (LARGURA - 2 * MARGEM - (len(self.cartoes) - 1) * ESPACO) / len(self.cartoes)
            for i, (tipo, papeis, titulo) in enumerate(self.cartoes):
                visuais.append(self._visual(n, tipo, papeis, titulo, MARGEM + i * (w + ESPACO), y, w, 100)); n += 1
            y += 100 + ESPACO
        if self.graficos:
            colunas = 1 if len(self.graficos) == 1 else 2
            linhas = math.ceil(len(self.graficos) / colunas)
            w = (LARGURA - 2 * MARGEM - (colunas - 1) * ESPACO) / colunas
            h = (ALTURA - MARGEM - y - (linhas - 1) * ESPACO) / linhas
            for i, (tipo, papeis, titulo) in enumerate(self.graficos):
                lin, col = divmod(i, colunas)
                visuais.append(self._visual(n, tipo, papeis, titulo,
                                            MARGEM + col * (w + ESPACO), y + lin * (h + ESPACO), w, h)); n += 1
        return visuais


# =====================================================================
# Para o Colab: empacotar tudo num .zip para baixar
# =====================================================================

def empacotar(nome_zip="BI_Autopilot"):
    """Junta dados/ e projeto/ em BI_Autopilot.zip.
    No Windows: botão direito > Extrair tudo > destino C:\\BI_Autopilot
    (o zip não tem pasta extra dentro, então fica C:\\BI_Autopilot\\dados e C:\\BI_Autopilot\\projeto)."""
    temporaria = Path("_pacote")
    if temporaria.exists():
        shutil.rmtree(temporaria)
    shutil.copytree(PASTA_DADOS, temporaria / "dados")
    shutil.copytree(PASTA_PROJETO, temporaria / "projeto")
    arquivo = shutil.make_archive(nome_zip, "zip", temporaria)
    shutil.rmtree(temporaria)
    print(f"✅ {arquivo} pronto.")
    print("   No Windows: botão direito > Extrair tudo > destino C:\\BI_Autopilot")
    print("   Confira: tem que existir C:\\BI_Autopilot\\dados\\base_tratada.xlsx")
    return arquivo


ajuda(silencioso=True)  # já liga a ajuda em português assim que o motor é importado

"""Etapa 1 - ingestao: carrega o que o escritorio recebe (dados/brutos) no DuckDB (dados/contabil.duckdb).

  cadastro_clientes.xlsx      carteira do escritorio
  balancetes/<cliente>/*.xlsx um balancete por mes, exportado do sistema contabil de cada cliente
  folha_resumo.csv            funcionarios e total da folha por cliente e mes

Os balancetes chegam em dois layouts, de dois sistemas contabeis:

  layout A  Conta | Descricao | Saldo anterior | Debito | Credito | Saldo atual     (conta 1.1.01, saldo com sinal)
  layout B  Codigo | Nome da conta | Sld. Ant. | D/C | Debitos | Creditos | Sld. Final | D/C   (conta 1101, saldo + D/C)

O importador reconhece o layout pelo cabecalho e entrega tudo num formato so:
conta com pontos e saldo com sinal (devedor positivo, credor negativo).
"""
import os
import re
from pathlib import Path

import duckdb
import openpyxl
import pandas as pd

RAIZ = Path(__file__).resolve().parents[1]
DADOS = Path(os.environ.get("CONTABIL_DADOS") or RAIZ / "dados").resolve()
BRUTOS = DADOS / "brutos"
CONFIG = RAIZ / "config"
BANCO = DADOS / "contabil.duckdb"


def conta_padrao(codigo):
    d = re.sub(r"\D", "", str(codigo))
    return f"{d[0]}.{d[1]}.{d[2:]}" if len(d) >= 4 else None


def ler_balancete(arquivo):
    wb = openpyxl.load_workbook(arquivo, read_only=True, data_only=True)
    try:
        linhas = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    finally:
        wb.close()
    txt = lambda r: [str(x).strip() if x is not None else "" for x in r]
    h = next((i for i, r in enumerate(linhas[:15]) if {"Conta", "Código"} & set(txt(r))), None)
    if h is None:
        raise ValueError(f"{arquivo.name}: cabecalho de balancete nao reconhecido")
    cab = txt(linhas[h])
    out = []
    for r in linhas[h + 1:]:
        if not r or r[0] is None:
            continue
        if "Conta" in cab:                                   # layout A
            ant, deb, cre, fim = (float(r[cab.index(c)] or 0) for c in ("Saldo anterior", "Débito", "Crédito", "Saldo atual"))
            layout = "A"
        else:                                                # layout B: sinal vem das colunas D/C (a 1a e do saldo anterior, a 2a do final)
            i_ant, i_fim = cab.index("Sld. Ant."), cab.index("Sld. Final")
            ant = float(r[i_ant] or 0) * (1 if str(r[i_ant + 1]).strip().upper() == "D" else -1)
            fim = float(r[i_fim] or 0) * (1 if str(r[i_fim + 1]).strip().upper() == "D" else -1)
            deb, cre = float(r[cab.index("Débitos")] or 0), float(r[cab.index("Créditos")] or 0)
            layout = "B"
        out.append((conta_padrao(r[0]), ant, deb, cre, fim, layout))
    return out


def main():
    con = duckdb.connect(str(BANCO))
    linhas = []
    for arq in sorted((BRUTOS / "balancetes").glob("*/balancete_*.xlsx")):
        comp = re.search(r"_(\d{4}-\d{2})", arq.name).group(1)
        for conta, ant, deb, cre, fim, layout in ler_balancete(arq):
            linhas.append((arq.parent.name, comp + "-01", conta, ant, deb, cre, fim, layout, arq.name))
    df = pd.DataFrame(linhas, columns=["CODIGO_CLIENTE", "COMPETENCIA", "CONTA", "SALDO_ANTERIOR", "DEBITO", "CREDITO", "SALDO_FINAL", "LAYOUT", "ARQUIVO"])
    df["COMPETENCIA"] = pd.to_datetime(df["COMPETENCIA"])
    con.register("df_tmp", df)
    con.execute("CREATE OR REPLACE TABLE balancete_bruto AS SELECT * REPLACE (CAST(COMPETENCIA AS DATE) AS COMPETENCIA) FROM df_tmp")
    con.unregister("df_tmp")

    cad = pd.read_excel(BRUTOS / "cadastro_clientes.xlsx", sheet_name="Clientes", header=2)
    cad.columns = ["CODIGO_CLIENTE", "RAZAO_SOCIAL", "CNPJ", "ATIVIDADE", "CIDADE", "REGIME", "ANEXO_SIMPLES", "SISTEMA", "RESPONSAVEL", "CLIENTE_DESDE", "HONORARIO"]
    con.register("df_tmp", cad)
    con.execute("CREATE OR REPLACE TABLE cadastro_clientes AS SELECT * REPLACE (CAST(CLIENTE_DESDE AS DATE) AS CLIENTE_DESDE, CAST(ANEXO_SIMPLES AS VARCHAR) AS ANEXO_SIMPLES) FROM df_tmp")
    con.unregister("df_tmp")

    con.execute(f"""CREATE OR REPLACE TABLE folha_resumo AS
        SELECT CODIGO_CLIENTE, CAST(COMPETENCIA || '-01' AS DATE) AS COMPETENCIA, FUNCIONARIOS, TOTAL_FOLHA
        FROM read_csv('{(BRUTOS / "folha_resumo.csv").as_posix()}', delim=';', header=true, decimal_separator=',',
                      types={{'COMPETENCIA': 'VARCHAR', 'FUNCIONARIOS': 'BIGINT', 'TOTAL_FOLHA': 'DOUBLE'}})""")
    for nome in ("plano_de_contas", "simples_nacional"):
        con.execute(f"CREATE OR REPLACE TABLE cfg_{nome} AS SELECT * FROM read_csv('{(CONFIG / (nome + '.csv')).as_posix()}', delim=';', header=true{', all_varchar=true' if nome == 'plano_de_contas' else ''})")   # conta 1.1.01 nao e data
    for t in ("balancete_bruto", "cadastro_clientes", "folha_resumo"):
        print(f"{t}: {con.execute(f'SELECT COUNT(*) FROM {t}').fetchone()[0]:,} linhas")
    print("balancetes por layout:", con.execute("SELECT LAYOUT, COUNT(DISTINCT ARQUIVO || CODIGO_CLIENTE) FROM balancete_bruto GROUP BY 1 ORDER BY 1").fetchall())
    con.close()


if __name__ == "__main__":
    main()

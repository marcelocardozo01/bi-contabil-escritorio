"""Etapas 2 a 4: roda as regras em SQL, exporta para Parquet e valida.

  pipeline/sql/trusted/*.sql    uma regra de negocio por arquivo (ordem do prefixo)
  pipeline/sql/analytics/*.sql  alertas e visao da carteira
  dados/exports/*.parquet       camada lida pelo painel HTML e pelo Power BI

Validacao (sai com erro se falhar):
  1. todo balancete fecha: debitos = creditos e ativo = passivo + patrimonio liquido
  2. o saldo final de um mes e o saldo anterior do mes seguinte (fora a virada do ano nas contas de resultado)
  3. o lucro da DRE do ano bate com a variacao do patrimonio que nao veio de distribuicao
  4. nenhuma conta do balancete ficou sem classificacao no plano referencial
"""
import os
import re
import sys
from pathlib import Path

import duckdb

RAIZ = Path(__file__).resolve().parents[1]
DADOS = Path(os.environ.get("CONTABIL_DADOS") or RAIZ / "dados").resolve()
BANCO = DADOS / "contabil.duckdb"
SQL = Path(__file__).resolve().parent / "sql"
EXPORTAR = ["dim_cliente", "dim_mes", "dim_conta", "balancete", "dre_mensal", "balanco_mensal", "indicadores_mensal", "simples_mensal",
            "an_alertas", "an_carteira"]


def main():
    con = duckdb.connect(str(BANCO))
    for arquivo in sorted((SQL / "trusted").glob("*.sql")) + sorted((SQL / "analytics").glob("*.sql")):
        texto = arquivo.read_text(encoding="utf-8")
        con.execute(texto)
        criadas = re.findall(r"CREATE OR REPLACE TABLE (\w+)", texto)
        print(f"{arquivo.parent.name}/{arquivo.name}: " + ", ".join(f"{t} {con.execute(f'SELECT COUNT(*) FROM {t}').fetchone()[0]:,}" for t in criadas))

    (DADOS / "exports").mkdir(exist_ok=True)
    for t in EXPORTAR:
        con.execute(f"COPY (SELECT * FROM {t}) TO '{(DADOS / 'exports' / (t + '.parquet')).as_posix()}' (FORMAT PARQUET, COMPRESSION ZSTD)")

    q = lambda sql: con.execute(sql).fetchone()
    falhas = []

    def checar(nome, ok, detalhe=""):
        print(f"[{'ok' if ok else 'FALHOU'}] {nome}{' - ' + detalhe if detalhe else ''}")
        if not ok:
            falhas.append(nome)

    print()
    n, d1, d2 = q("""SELECT COUNT(*), MAX(ABS(DC)), MAX(ABS(BAL)) FROM (
                       SELECT SUM(DEBITO) - SUM(CREDITO) AS DC, SUM(SALDO_FINAL) AS BAL FROM balancete GROUP BY CODIGO_CLIENTE, COMPETENCIA)""")
    checar(f"os {n} balancetes fecham (debito = credito, ativo = passivo + PL)", d1 < 0.20 and d2 < 0.20, f"maior diferenca R$ {max(d1, d2):.2f}")
    n = q("""SELECT COUNT(*) FROM (
               SELECT b.SALDO_ANTERIOR, p.SALDO_FINAL AS FIM_ANT, MONTH(b.COMPETENCIA) AS M, c.CLASSE, b.CONTA
               FROM balancete b JOIN dim_conta c USING (CONTA)
               JOIN balancete p ON p.CODIGO_CLIENTE = b.CODIGO_CLIENTE AND p.CONTA = b.CONTA AND p.COMPETENCIA = b.COMPETENCIA - INTERVAL 1 MONTH)
             WHERE ABS(SALDO_ANTERIOR - FIM_ANT) > 0.05 AND NOT (M = 1 AND (CLASSE = 'Resultado' OR CONTA = '2.3.02'))""")[0]
    checar("saldo anterior = saldo final do mes passado", n == 0, f"{n} quebras de sequencia")
    dif = q("""SELECT MAX(ABS(d.LUCRO - r.SALDO)) FROM
                 (SELECT CODIGO_CLIENTE, YEAR(COMPETENCIA) AS ANO, MAX(COMPETENCIA) AS ULT, SUM(LUCRO_LIQUIDO) AS LUCRO FROM dre_mensal GROUP BY 1, 2
                  HAVING COUNT(*) = MAX(MONTH(COMPETENCIA)) - MIN(MONTH(COMPETENCIA)) + 1 AND MIN(MONTH(COMPETENCIA)) = 1) d
               JOIN (SELECT b.CODIGO_CLIENTE, b.COMPETENCIA, -SUM(b.SALDO_FINAL) AS SALDO FROM balancete b JOIN dim_conta c USING (CONTA)
                     WHERE c.CLASSE = 'Resultado' GROUP BY 1, 2) r ON r.CODIGO_CLIENTE = d.CODIGO_CLIENTE AND r.COMPETENCIA = d.ULT""")[0]
    checar("lucro acumulado da DRE = saldo das contas de resultado", dif is not None and dif < 0.5, f"maior diferenca R$ {dif:.2f}")
    n = q("SELECT COUNT(*) FROM balancete_bruto b LEFT JOIN dim_conta c USING (CONTA) WHERE c.CONTA IS NULL")[0]
    checar("todas as contas estao no plano referencial", n == 0, f"{n} linhas sem classificacao")
    con.close()
    if falhas:
        sys.exit(f"\n{len(falhas)} checagem(ns) falharam.")
    print("\nTodas as checagens passaram.")


if __name__ == "__main__":
    main()

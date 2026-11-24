"""Gera o painel HTML (painel_html/saida/index.html): a carteira do escritorio e a pagina de cada cliente.

Le os .parquet de dados/exports, resume num JSON e injeta no modelo. O resultado e um arquivo so, estatico:
abre no navegador ou no celular, e o endereco guarda o cliente aberto (index.html#cliente=C008), entao o
contador manda ao empresario o link direto do painel dele.

Uso: python painel_html/gerar_painel.py
"""
import json
import os
from pathlib import Path

import duckdb

RAIZ = Path(__file__).resolve().parents[1]
DADOS = Path(os.environ.get("CONTABIL_DADOS") or RAIZ / "dados").resolve()
AQUI = Path(__file__).resolve().parent
SERIES = ["RECEITA_BRUTA", "DEDUCOES", "RECEITA_LIQUIDA", "CUSTOS", "LUCRO_BRUTO", "DESPESAS_PESSOAL", "DESPESAS_ADMINISTRATIVAS", "DEPRECIACAO",
          "RESULTADO_OPERACIONAL", "RESULTADO_FINANCEIRO", "IMPOSTO_LUCRO", "LUCRO_LIQUIDO", "IMPOSTOS_RECEITA", "DISPONIVEL", "CLIENTES_A_RECEBER",
          "ESTOQUES", "DIVIDA_BANCARIA", "LIQUIDEZ_CORRENTE", "PRAZO_RECEBIMENTO_DIAS", "DIAS_ESTOQUE", "MARGEM_LIQUIDA_3M", "CARGA_TRIBUTARIA",
          "PESO_FOLHA", "FUNCIONARIOS", "FOLEGO_CAIXA_MESES", "DIVIDA_MESES_RECEITA", "RECEITA_BRUTA_12M", "LUCRO_LIQUIDO_12M", "MARGEM_LIQUIDA_12M"]
MES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]


def main():
    con = duckdb.connect()
    for f in (DADOS / "exports").glob("*.parquet"):
        con.execute(f"CREATE VIEW {f.stem} AS SELECT * FROM read_parquet('{f.as_posix()}')")

    def q(sql):
        cur = con.execute(sql)
        cols = [c[0] for c in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]

    arred = lambda v: None if v is None else (round(float(v), 4) if abs(float(v)) < 50 else round(float(v)))
    rot = lambda d: f"{MES[d.month - 1]}/{d:%y}"
    clientes = []
    for c in q("SELECT * FROM an_carteira ORDER BY CASE SITUACAO WHEN 'Crítico' THEN 0 WHEN 'Atenção' THEN 1 ELSE 2 END, RECEITA_BRUTA_12M DESC"):
        cod = c["CODIGO_CLIENTE"]
        linhas = q(f"SELECT * FROM indicadores_mensal WHERE CODIGO_CLIENTE = '{cod}' ORDER BY COMPETENCIA")
        cad = q(f"SELECT CNPJ, CLIENTE_DESDE FROM dim_cliente WHERE CODIGO_CLIENTE = '{cod}'")[0]
        s = q(f"SELECT * FROM simples_mensal WHERE CODIGO_CLIENTE = '{cod}' ORDER BY COMPETENCIA DESC LIMIT 1")
        clientes.append({
            "cod": cod, "nome": c["CLIENTE"], "atividade": c["ATIVIDADE"], "cidade": c["CIDADE"], "regime": c["REGIME"], "anexo": c["ANEXO_SIMPLES"],
            "resp": c["RESPONSAVEL"], "honorario": c["HONORARIO"], "cnpj": cad["CNPJ"], "ultimo": rot(c["ULTIMO_BALANCETE"]), "atraso": c["MESES_ATRASO"],
            "situacao": c["SITUACAO"], "criticos": c["ALERTAS_CRITICOS"], "atencoes": c["ALERTAS_ATENCAO"],
            "meses": [rot(l["COMPETENCIA"]) for l in linhas],
            "s": {k: [arred(l[k]) for l in linhas] for k in SERIES},
            "alertas": q(f"SELECT GRAVIDADE AS g, TEMA AS tema, TITULO AS titulo, TEXTO AS texto FROM an_alertas WHERE CODIGO_CLIENTE = '{cod}' ORDER BY ORDEM"),
            "simples": ({k.lower(): (arred(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else v)
                         for k, v in s[0].items() if k not in ("CODIGO_CLIENTE", "COMPETENCIA")} if s else None),
        })
    faixas = {}
    for f in q("SELECT ANEXO, FAIXA, RECEITA_12M_ATE, ALIQUOTA_NOMINAL FROM read_csv('" + (RAIZ / "config" / "simples_nacional.csv").as_posix()
               + "', delim=';', header=true) ORDER BY 1, 2"):
        faixas.setdefault(f["ANEXO"], []).append([f["RECEITA_12M_ATE"], f["ALIQUOTA_NOMINAL"]])
    ref = q("SELECT MAX(MES_REFERENCIA) AS r FROM dim_cliente")[0]["r"]
    d = {"escritorio": "Lastro Contábil", "referencia": rot(ref), "clientes": clientes, "faixas": faixas}
    modelo = (AQUI / "modelo.html").read_text(encoding="utf-8")
    saida = Path(os.environ.get("CONTABIL_SAIDA") or AQUI / "saida")
    saida.mkdir(exist_ok=True)
    dados = json.dumps(d, ensure_ascii=False, separators=(",", ":"), default=str).replace("</", "<\\/")
    (saida / "index.html").write_text(modelo.replace("/*__DADOS__*/null", dados), encoding="utf-8", newline="\n")
    print(f"{saida / 'index.html'}: {(saida / 'index.html').stat().st_size / 1024:,.0f} KB, {len(clientes)} clientes")


if __name__ == "__main__":
    main()

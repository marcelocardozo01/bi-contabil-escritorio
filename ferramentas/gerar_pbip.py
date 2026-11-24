"""Gera o relatorio Power BI (formato PBIP) do painel contabil a partir desta definicao em Python.

O PBIP e texto: modelo em TMDL e relatorio em JSON. Modelo, medidas e paginas ficam descritos aqui e sao
escritos em powerbi/Lastro_Contabil.*. Mudou uma medida ou um bloco? Edita este arquivo, roda de novo e o
diff do Git mostra o que mudou.

O desenho usa HTML e CSS dentro do Power BI: medidas DAX devolvem o HTML de cada bloco (cabecalho, cartoes,
tabela da carteira, pauta, alertas, "cada R$ 100", DRE, Simples) e o visual "HTML Content", baixado da
AppSource ao abrir, desenha o resultado. Series no tempo e segmentacoes sao visuais nativos, no mesmo tema.

Uso: python ferramentas/gerar_pbip.py
"""
import hashlib
import json
import re
import shutil
import uuid
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
PBI = RAIZ / "powerbi"
NOME = "Lastro_Contabil"
PASTA_PADRAO = "C:\\lastro-contabil\\dados\\exports\\"
VISUAL_HTML = "htmlContent443BE3AD55E043BF878BED274D3A6865"
NS = uuid.UUID("9a1c7e52-3b64-4f0d-8e21-6c5d4b3a2f10")
tag = lambda *p: str(uuid.uuid5(NS, "/".join(p)))

# Paleta (a mesma do painel HTML)
C = dict(fundo="#061A1C", painel="#0C2A2C", painel2="#123A3B", linha="#1B4A49", tinta="#F4FBFA", tinta2="#A9C6C2", mudo="#7A9A96",
         menta="#34E0BE", menta2="#1F6F66", coral="#FF6B5E", ambar="#F5C04A", lilas="#9B8CF2", laranja="#FF8A65", rosa="#F78FB3")
FONTE = f"font-family:Segoe UI,sans-serif;color:{C['tinta']};line-height:1.35"
VIDRO = f"background:linear-gradient(180deg,{C['painel2']},{C['painel']});border:1px solid {C['linha']};border-radius:16px;box-sizing:border-box"

T = {"date": ("dateTime", "type date"), "int": ("int64", "Int64.Type"), "num": ("double", "type number"), "text": ("string", "type text")}
N = "num"
TABELAS = {
    "dCliente": ("dim_cliente", {"CODIGO_CLIENTE": "text", "CLIENTE": "text", "ATIVIDADE": "text", "CIDADE": "text", "REGIME": "text",
                                 "ANEXO_SIMPLES": "text", "RESPONSAVEL": "text", "HONORARIO": N, "ULTIMO_BALANCETE": "date", "MESES_ATRASO": "int"}, [], {}),
    "dMes": ("dim_mes", {"COMPETENCIA": "date", "ANO_MES": "text", "MES_ROTULO": "text"}, ["ANO_MES"], {"MES_ROTULO": "ANO_MES"}),
    "fIndicador": ("indicadores_mensal", {"CODIGO_CLIENTE": "text", "COMPETENCIA": "date", "RECEITA_BRUTA": N, "DEDUCOES": N, "IMPOSTOS_RECEITA": N,
                                          "RECEITA_LIQUIDA": N, "CUSTOS": N, "LUCRO_BRUTO": N, "DESPESAS_PESSOAL": N, "DESPESAS_ADMINISTRATIVAS": N,
                                          "DEPRECIACAO": N, "RESULTADO_OPERACIONAL": N, "RESULTADO_FINANCEIRO": N, "IMPOSTO_LUCRO": N, "LUCRO_LIQUIDO": N,
                                          "DISPONIVEL": N, "CLIENTES_A_RECEBER": N, "ESTOQUES": N, "ATIVO_CIRCULANTE": N, "PASSIVO_CIRCULANTE": N,
                                          "DIVIDA_BANCARIA": N}, "*", {}),
    "fCarteira": ("an_carteira", {"CODIGO_CLIENTE": "text", "SITUACAO": "text", "ALERTAS_CRITICOS": "int", "ALERTAS_ATENCAO": "int",
                                  "RECEITA_BRUTA_12M": N, "MARGEM_LIQUIDA_12M": N, "LIQUIDEZ_CORRENTE": N}, ["CODIGO_CLIENTE", "RECEITA_BRUTA_12M", "MARGEM_LIQUIDA_12M", "LIQUIDEZ_CORRENTE", "ALERTAS_CRITICOS", "ALERTAS_ATENCAO"], {}),
    "fAlerta": ("an_alertas", {"CODIGO_CLIENTE": "text", "GRAVIDADE": "text", "TEMA": "text", "ORDEM": "int", "TITULO": "text", "TEXTO": "text"},
                ["CODIGO_CLIENTE", "ORDEM"], {}),
    "fSimples": ("simples_mensal", {"CODIGO_CLIENTE": "text", "COMPETENCIA": "date", "ANEXO": "text", "RBT12": N, "FAIXA": "int", "FOLGA_FAIXA": N,
                                    "USO_FAIXA": N, "ALIQUOTA_EFETIVA": N, "LIMITE_SIMPLES": N, "USO_LIMITE_SIMPLES": N}, "*", {}),
}
RELACOES = [("fIndicador.CODIGO_CLIENTE", "dCliente.CODIGO_CLIENTE"), ("fIndicador.COMPETENCIA", "dMes.COMPETENCIA"),
            ("fCarteira.CODIGO_CLIENTE", "dCliente.CODIGO_CLIENTE"), ("fAlerta.CODIGO_CLIENTE", "dCliente.CODIGO_CLIENTE"),
            ("fSimples.CODIGO_CLIENTE", "dCliente.CODIGO_CLIENTE"), ("fSimples.COMPETENCIA", "dMes.COMPETENCIA")]

REAIS, PCT = r"\R\$ #,##0", "0.0%"
soma = lambda c: f"SUM(fIndicador[{c}])"
ULT = lambda expr: ["VAR m = MAX(fIndicador[COMPETENCIA])", f"RETURN CALCULATE({expr}, FILTER(ALL(dMes), dMes[COMPETENCIA] = m))"]
MEDIDAS = {"fIndicador": [
    ("Faturamento", REAIS, soma("RECEITA_BRUTA")), ("Impostos e devoluções", REAIS, soma("DEDUCOES")), ("Receita líquida", REAIS, soma("RECEITA_LIQUIDA")),
    ("Custo total", REAIS, soma("CUSTOS")), ("Lucro bruto", REAIS, soma("LUCRO_BRUTO")), ("Despesas com pessoal", REAIS, soma("DESPESAS_PESSOAL")),
    ("Despesas administrativas", REAIS, soma("DESPESAS_ADMINISTRATIVAS")), ("Despesa de depreciação", REAIS, soma("DEPRECIACAO")),
    ("Resultado operacional", REAIS, soma("RESULTADO_OPERACIONAL")), ("Resultado financeiro", REAIS, soma("RESULTADO_FINANCEIRO")),
    ("IRPJ e CSLL", REAIS, soma("IMPOSTO_LUCRO")), ("Lucro líquido", REAIS, soma("LUCRO_LIQUIDO")),
    ("Margem líquida", PCT, "DIVIDE([Lucro líquido], [Receita líquida])"), ("Margem bruta", PCT, "DIVIDE([Lucro bruto], [Receita líquida])"),
    ("Impostos", REAIS, "-(SUM(fIndicador[IMPOSTOS_RECEITA]) + SUM(fIndicador[IMPOSTO_LUCRO]))"),
    ("Carga tributária", PCT, "DIVIDE([Impostos], [Faturamento])"),
    ("Disponível", REAIS, ULT(soma("DISPONIVEL"))), ("Dívida bancária", REAIS, ULT(soma("DIVIDA_BANCARIA"))),
    ("Liquidez corrente", "0.00", ["VAR m = MAX(fIndicador[COMPETENCIA])", "VAR f = FILTER(ALL(dMes), dMes[COMPETENCIA] = m)",
                                   "RETURN DIVIDE(CALCULATE(SUM(fIndicador[ATIVO_CIRCULANTE]), f), CALCULATE(SUM(fIndicador[PASSIVO_CIRCULANTE]), f))"]),
]}


# ----------------------------------------------------------------------------
# Moldes das medidas HTML
# ----------------------------------------------------------------------------
def s(txt):
    return '"' + txt.replace('"', '""') + '"'


def j(*partes):
    return " & ".join(partes)


BRL = lambda v: f'"R$ " & FORMAT({v}, "#,##0")'
JANELAS = ["VAR m = MAX(fIndicador[COMPETENCIA])",
           "VAR fM = FILTER(ALL(dMes), dMes[COMPETENCIA] = m)",
           "VAR fAnt = FILTER(ALL(dMes), dMes[COMPETENCIA] = EDATE(m, -1))",
           "VAR fAA = FILTER(ALL(dMes), dMes[COMPETENCIA] = EDATE(m, -12))",
           "VAR f3 = FILTER(ALL(dMes), dMes[COMPETENCIA] > EDATE(m, -3) && dMes[COMPETENCIA] <= m)",
           "VAR f12 = FILTER(ALL(dMes), dMes[COMPETENCIA] > EDATE(m, -12) && dMes[COMPETENCIA] <= m)"]


def painel(titulo, sub, altura):
    return (f"<div style='{FONTE};{VIDRO};padding:14px 18px 10px;height:{altura}px;overflow:hidden'>"
            f"<div style='font-size:15px;font-weight:700'>{titulo}</div><div style='font-size:12px;color:{C['tinta2']};margin:1px 0 10px'>{sub}</div>")


def cabecalho(titulo):
    return ["VAR m = MAX(fIndicador[COMPETENCIA])",
            'VAR quem = IF(HASONEVALUE(dCliente[CLIENTE]), SELECTEDVALUE(dCliente[CLIENTE]), FORMAT(DISTINCTCOUNT(fIndicador[CODIGO_CLIENTE]), "0") & " clientes")',
            "RETURN",
            "    " + j(s(f"<div style='{FONTE};height:70px;box-sizing:border-box;padding:11px 20px 0;background:linear-gradient(90deg,{C['painel2']},{C['fundo']} 70%);"
                         f"border-bottom:1px solid {C['linha']}'><div style='display:flex;align-items:center;gap:9px;font-size:12.5px;font-weight:600;color:{C['menta']}'>"
                         f"<span style='width:13px;height:13px;border-radius:3px 3px 6px 6px;background:linear-gradient(160deg,#7FF0DC,{C['menta']})'></span>Lastro Contábil</div>"
                         f"<div style='font-size:23px;font-weight:800;letter-spacing:-.4px;line-height:1.2'>{titulo}<span style='font-size:13px;font-weight:400;color:{C['tinta2']};margin-left:12px'>"),
                  "quem", s(", balancete de "), 'FORMAT(m, "mm/yyyy")', s("</span></div></div>"))]


def cartoes(variaveis, specs, altura=112):
    """specs: (rotulo, texto do valor em DAX, texto de apoio em DAX, condicao DAX de 'ruim' ou None)."""
    blocos = []
    for rot, val, apoio, ruim in specs:
        cor = f'IF({ruim}, "{C["coral"]}", "{C["menta"]}")' if ruim else s(C["menta"])
        tinta = f'IF({ruim}, "{C["coral"]}", "{C["tinta"]}")' if ruim else s(C["tinta"])
        blocos.append(j(
            s(f"<div style='flex:1;min-width:0;{VIDRO};height:{altura}px;padding:13px 15px;position:relative;overflow:hidden'>"
              f"<div style='position:absolute;left:0;right:0;top:0;height:3px;background:linear-gradient(90deg,"), cor, s(f",rgba(0,0,0,0))'></div>"
              f"<div style='font-size:12px;color:{C['tinta2']};white-space:nowrap;overflow:hidden;text-overflow:ellipsis'>{rot}</div>"
              f"<div style='font-size:27px;font-weight:700;letter-spacing:-.5px;line-height:1.25;white-space:nowrap;color:"), tinta, s("'>"), val,
            s(f"</div><div style='font-size:11.5px;color:{C['tinta2']};margin-top:5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis'>"), apoio, s("</div></div>")))
    return variaveis + ["RETURN", "    " + s(f"<div style='{FONTE};display:flex;gap:12px'>")] + ["    & " + b for b in blocos] + ["    & " + s("</div>")]


CHIP = (f'"<span style=\'display:inline-block;font-size:10px;font-weight:600;border-radius:99px;padding:0 7px;line-height:13px;border:1px solid " & '
        f'SWITCH(fCarteira[SITUACAO], "Crítico", "{C["coral"]};color:{C["coral"]}", "Atenção", "{C["ambar"]};color:{C["ambar"]}", "{C["menta"]};color:{C["menta"]}") & "\'>" & '
        f'SWITCH(fCarteira[SITUACAO], "Crítico", "✕ ", "Atenção", "! ", "✓ ") & fCarteira[SITUACAO] & "</span>"')
ELIP = "white-space:nowrap;overflow:hidden;text-overflow:ellipsis"


def tabela_carteira(altura):
    grade = "92px 1.5fr .9fr 96px 70px 62px 1.5fr 60px"
    cab = "".join(f"<div style='text-align:{a}'>{t}</div>" for t, a in [("Situação", "left"), ("Cliente", "left"), ("Regime", "left"), ("Faturamento 12 m", "right"),
                                                                      ("Margem", "right"), ("Liquidez", "right"), ("Principal alerta", "left"), ("Resp.", "left")])
    linha = j(
        s(f"<div style='display:grid;grid-template-columns:{grade};gap:10px;align-items:center;font-size:11px;line-height:1.1;padding:1.5px 0;border-bottom:1px solid {C['linha']}'><div>"),
        CHIP, s(f"</div><div style='{ELIP}'><b style='font-weight:600'>"), "[@nome]", s(f"</b> <span style='color:{C['mudo']}'>"), "[@ativ]",
        s(f"</span></div><div style='{ELIP};color:{C['tinta2']}'>"), "[@reg]", s("</div><div style='text-align:right'>"), BRL("fCarteira[RECEITA_BRUTA_12M]"),
        s("</div><div style='text-align:right;color:"), f'IF(fCarteira[MARGEM_LIQUIDA_12M] < 0, "{C["coral"]}", "{C["tinta"]}")', s("'>"),
        'FORMAT(fCarteira[MARGEM_LIQUIDA_12M], "0.0%")', s("</div><div style='text-align:right;color:"),
        f'IF(fCarteira[LIQUIDEZ_CORRENTE] < 1, "{C["coral"]}", "{C["tinta"]}")', s("'>"), 'FORMAT(fCarteira[LIQUIDEZ_CORRENTE], "0.00")',
        s(f"</div><div style='{ELIP};color:{C['tinta2']}'>"), 'IF(ISBLANK([@alerta]), "Sem pontos de atenção", [@alerta])',
        s(f"</div><div style='color:{C['tinta2']}'>"), "[@resp]", s("</div></div>"))
    return [
        'VAR t = ADDCOLUMNS(fCarteira, "@nome", RELATED(dCliente[CLIENTE]), "@ativ", RELATED(dCliente[ATIVIDADE]), "@reg", RELATED(dCliente[REGIME]),',
        '    "@resp", RELATED(dCliente[RESPONSAVEL]),',
        '    "@alerta", VAR cod = fCarteira[CODIGO_CLIENTE] RETURN MAXX(TOPN(1, FILTER(fAlerta, fAlerta[CODIGO_CLIENTE] = cod), fAlerta[ORDEM], ASC), fAlerta[TITULO]))',
        "RETURN",
        "    " + s(painel("Todos os clientes", "Ordenados pelos que mais precisam de atenção. Faturamento e margem dos últimos 12 meses.", altura)
                   + f"<div style='display:grid;grid-template-columns:{grade};gap:10px;font-size:11px;color:{C['mudo']};padding-bottom:4px;border-bottom:1px solid {C['menta2']}'>{cab}</div>"),
        f'    & CONCATENATEX(t, {linha}, "", SWITCH(fCarteira[SITUACAO], "Crítico", 0, "Atenção", 1, 2), ASC, fCarteira[RECEITA_BRUTA_12M], DESC)',
        "    & " + s("</div>")]


def pauta(altura):
    caso = j(
        s(f"<div style='border:1px solid rgba(255,107,94,.3);border-left:4px solid {C['coral']};border-radius:10px;padding:5px 10px;margin-bottom:5px;"
          f"background:linear-gradient(90deg,rgba(255,107,94,.13),rgba(255,107,94,.02) 60%)'><div style='display:flex;justify-content:space-between;gap:8px;align-items:baseline'>"
          f"<b style='font-size:12.5px;font-weight:700;{ELIP}'>"), "[@nome]", s(f"</b><span style='font-size:10.5px;color:#FFC4BD;white-space:nowrap'>"),
        '[@tit] & IF(fCarteira[ALERTAS_CRITICOS] > 1, " +" & FORMAT(fCarteira[ALERTAS_CRITICOS] - 1, "0"), "")',
        s(f"</span></div><div style='font-size:10.5px;color:{C['tinta2']};line-height:1.25;margin-top:1px;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden'>"),
        "[@txt]", s("</div></div>"))
    alerta = lambda col: (f'VAR cod = fCarteira[CODIGO_CLIENTE] RETURN MAXX(TOPN(1, FILTER(fAlerta, fAlerta[CODIGO_CLIENTE] = cod && fAlerta[GRAVIDADE] = "Crítico"), '
                          f'fAlerta[ORDEM], ASC), fAlerta[{col}])')
    return [
        'VAR t = ADDCOLUMNS(FILTER(fCarteira, fCarteira[SITUACAO] = "Crítico"), "@nome", RELATED(dCliente[CLIENTE]),',
        f'    "@tit", {alerta("TITULO")},',
        f'    "@txt", {alerta("TEXTO")})',
        "RETURN",
        "    " + s(painel("Pauta da semana", "Um caso por cliente crítico, com o primeiro ponto a tratar.", altura)),
        f'    & IF(COUNTROWS(t) = 0, {s(f"<div style=\'font-size:12px;color:{C["tinta2"]}\'>Nenhum cliente crítico neste recorte.</div>")},',
        f'        CONCATENATEX(TOPN(7, t, fCarteira[RECEITA_BRUTA_12M], DESC), {caso}, "", fCarteira[RECEITA_BRUTA_12M], DESC))',
        "    & " + s("</div>")]


def alertas_cliente(altura):
    cor = f'IF(fAlerta[GRAVIDADE] = "Crítico", "{C["coral"]}", "{C["ambar"]}")'
    fundo = 'IF(fAlerta[GRAVIDADE] = "Crítico", "rgba(255,107,94,.13)", "rgba(245,192,74,.12)")'
    item = j(s("<div style='border-radius:10px;padding:7px 12px;margin-bottom:6px;border-left:4px solid "), cor, s(";background:linear-gradient(90deg,"), fundo,
             s(",rgba(0,0,0,0) 65%)'><div style='font-size:12.5px;font-weight:700'><span style='font-size:11px;margin-right:8px;color:"), cor, s("'>"),
             'IF(fAlerta[GRAVIDADE] = "Crítico", "✕ Crítico", "! Atenção")', s("</span>"),
             'IF(HASONEVALUE(dCliente[CLIENTE]), "", RELATED(dCliente[CLIENTE]) & ": ")', "fAlerta[TITULO]",
             s(f"</div><div style='font-size:11.5px;color:{C['tinta2']};line-height:1.3;margin-top:1px'>"), "fAlerta[TEXTO]", s("</div></div>"))
    return [
        "VAR n = COUNTROWS(fAlerta)",
        "RETURN",
        "    " + s(painel("O que conversar com o cliente", "Os pontos que o balancete levanta, em linguagem de dono de empresa.", altura)),
        f'    & IF(n = 0, {s(f"<div style=\'border-left:4px solid {C["menta"]};border-radius:10px;padding:9px 12px;background:rgba(52,224,190,.1);font-size:12.5px\'><b>✓ Nenhum ponto de atenção neste balancete</b></div>")},',
        f'        CONCATENATEX(TOPN(4, fAlerta, fAlerta[ORDEM], ASC, fAlerta[CODIGO_CLIENTE], ASC), {item}, "", fAlerta[ORDEM], ASC, fAlerta[CODIGO_CLIENTE], ASC))',
        "    & " + s("</div>")]


def cada_cem(altura):
    partes = [("imp", "Impostos e devoluções", C["lilas"], "#FFFFFF"), ("cus", "Mercadoria e serviço", C["menta2"], "#FFFFFF"), ("pes", "Pessoal", C["laranja"], "#FFFFFF"),
              ("adm", "Despesas administrativas", C["ambar"], C["fundo"]), ("jur", "Juros e tarifas", C["rosa"], C["fundo"]), ("luc", "Lucro", C["menta"], C["fundo"])]
    seg = lambda v, cor, tinta: (f'IF({v} > 0, "<div style=\'flex:" & FORMAT(INT({v}), "0") & " 0 0;min-width:3px;border-radius:7px;background:{cor};color:{tinta};font-size:12px;'
                                 f'font-weight:700;padding:8px 8px;white-space:nowrap;overflow:hidden;box-sizing:border-box\'>" & IF(DIVIDE({v}, rb) > 0.07, FORMAT(DIVIDE({v}, rb) * 100, "0"), "") & "</div>", "")')
    leg = lambda v, nome, cor: (f'IF({v} > 0, "<div style=\'display:flex;align-items:center;gap:7px\'><i style=\'width:10px;height:10px;border-radius:3px;background:{cor}\'></i>{nome}'
                                f'<b style=\'margin-left:auto;color:{C["tinta"]}\'>R$ " & FORMAT(DIVIDE({v}, rb) * 100, "0.0") & "</b></div>", "")')
    return JANELAS + [
        "VAR rb = CALCULATE([Faturamento], f3)",
        "VAR imp = -CALCULATE(SUM(fIndicador[DEDUCOES]) + SUM(fIndicador[IMPOSTO_LUCRO]), f3)",
        "VAR cus = -CALCULATE([Custo total], f3)",
        "VAR pes = -CALCULATE([Despesas com pessoal], f3)",
        "VAR adm = -CALCULATE([Despesas administrativas] + [Despesa de depreciação], f3)",
        "VAR jur = -CALCULATE([Resultado financeiro], f3)",
        "VAR luc = CALCULATE([Lucro líquido], f3)",
        "RETURN",
        "    " + s(painel("Para onde vai cada R$ 100 que entram", "Média do último trimestre.", altura) + "<div style='display:flex;gap:3px;height:34px;margin-bottom:12px'>"),
    ] + ["    & " + seg(v, cor, tinta) for v, _, cor, tinta in partes] + [
        "    & " + s(f"</div><div style='display:grid;grid-template-columns:1fr 1fr;gap:5px 22px;font-size:12px;color:{C['tinta2']}'>"),
    ] + ["    & " + leg(v, nome, cor) for v, nome, cor, _ in partes] + [
        "    & " + s("</div>"),
        f'    & IF(luc < 0, "<div style=\'font-size:11.5px;color:{C["coral"]};margin-top:9px\'>Saíram R$ " & FORMAT(DIVIDE(rb - luc, rb) * 100, "0.0") & " para cada R$ 100 que entraram.</div>", "")',
        "    & " + s("</div>")]


def simples(altura):
    barra = lambda uso, rot, txt: j(
        s(f"<div style='font-size:12px;color:{C['tinta2']};display:flex;justify-content:space-between;margin-top:9px'><span>{rot}</span><b style='color:{C['tinta']}'>"), txt,
        s(f"</b></div><div style='height:9px;border-radius:5px;background:{C['fundo']};margin-top:4px;overflow:hidden'><div style='height:9px;border-radius:5px;"
          f"background:linear-gradient(90deg,{C['menta2']},{C['menta']});width:"), f'FORMAT(INT(MIN(1, {uso}) * 100), "0")', s("%'></div></div>"))
    kv = lambda rot, val: j(s(f"<div style='display:flex;justify-content:space-between;font-size:12.5px;margin-top:7px'><span style='color:{C['tinta2']}'>{rot}</span><b>"), val, s("</b></div>"))
    vazio = lambda txt: s(f"<div style='font-size:12.5px;color:{C['tinta2']}'>{txt}</div>")
    return [
        "VAR um = HASONEVALUE(dCliente[CODIGO_CLIENTE])",
        "VAR ms = MAX(fSimples[COMPETENCIA])",
        "VAR f = FILTER(ALL(dMes), dMes[COMPETENCIA] = ms)",
        "VAR rbt = CALCULATE(MAX(fSimples[RBT12]), f)",
        "VAR faixa = CALCULATE(MAX(fSimples[FAIXA]), f)",
        "VAR usoF = CALCULATE(MAX(fSimples[USO_FAIXA]), f)",
        "VAR usoL = CALCULATE(MAX(fSimples[USO_LIMITE_SIMPLES]), f)",
        "VAR folga = CALCULATE(MAX(fSimples[FOLGA_FAIXA]), f)",
        "VAR lim = CALCULATE(MAX(fSimples[LIMITE_SIMPLES]), f)",
        "VAR aliq = CALCULATE(MAX(fSimples[ALIQUOTA_EFETIVA]), f)",
        "RETURN",
        "    " + s(painel("Simples Nacional", "A faixa depende do faturamento dos últimos 12 meses.", altura)),
        f"    & IF(NOT um, {vazio('Selecione um cliente para ver a faixa do Simples.')},",
        f"        IF(ISBLANK(rbt), {vazio('Este cliente não está no Simples Nacional.')},",
        "            " + j(kv("Faturamento de 12 meses", BRL("rbt")), kv("Faixa atual", 'FORMAT(faixa, "0") & " de 6"'), kv("Alíquota efetiva", 'FORMAT(aliq, "0.0%")'),
                           barra("usoF", "Uso da faixa atual", 'IF(ISBLANK(folga), "acima do limite", "faltam " & ' + BRL("folga") + ')'),
                           barra("usoL", "Uso do limite do Simples", '"faltam " & ' + BRL("lim - rbt")),
                           s(f"<div style='font-size:10.5px;color:{C['mudo']};margin-top:10px'>Tabela simplificada para a demonstração.</div>")) + "))",
        "    & " + s("</div>")]


def dre(altura):
    linhas = [("Receita bruta", "[Faturamento]", 0), ("(-) Impostos e devoluções", "[Impostos e devoluções]", 0), ("Receita líquida", "[Receita líquida]", 1),
              ("(-) Custo da mercadoria e do serviço", "[Custo total]", 0), ("Lucro bruto", "[Lucro bruto]", 1), ("(-) Pessoal", "[Despesas com pessoal]", 0),
              ("(-) Despesas administrativas", "[Despesas administrativas]", 0), ("(-) Depreciação", "[Despesa de depreciação]", 0), ("Resultado operacional", "[Resultado operacional]", 1),
              ("Resultado financeiro", "[Resultado financeiro]", 0), ("(-) IRPJ e CSLL", "[IRPJ e CSLL]", 0), ("Lucro líquido", "[Lucro líquido]", 1)]
    grade = "1.7fr 1fr 1fr 1fr 1.1fr"
    cel = lambda med, f: (f'VAR x = CALCULATE({med}, {f}) RETURN "<div style=\'text-align:right;color:" & IF(x < 0, "{C["coral"]}", "{C["tinta"]}") & "\'>" & '
                          f'IF(ISBLANK(x), "–", FORMAT(x, "#,##0")) & "</div>"')
    out = JANELAS + ["RETURN",
                     "    " + j(s(painel("Resultado do mês", "Demonstração de resultado resumida: mês, mês anterior, um ano antes e últimos 12 meses.", altura)
                                  + f"<div style='display:grid;grid-template-columns:{grade};gap:10px;font-size:11.5px;color:{C['mudo']};padding-bottom:6px;border-bottom:1px solid {C['menta2']}'>"
                                    "<div>Em reais</div><div style='text-align:right'>"), 'FORMAT(m, "mm/yyyy")', s("</div><div style='text-align:right'>"),
                                'FORMAT(EDATE(m, -1), "mm/yyyy")', s("</div><div style='text-align:right'>"), 'FORMAT(EDATE(m, -12), "mm/yyyy")',
                                s("</div><div style='text-align:right'>12 meses</div></div>"))]
    for rot, med, total in linhas:
        peso = "font-weight:700;" if total else ""
        borda = C["menta2"] if total else C["linha"]
        out.append("    & " + s(f"<div style='display:grid;grid-template-columns:{grade};gap:10px;font-size:12.5px;{peso}padding:8px 0;border-bottom:1px solid {borda}'><div>{rot}</div>"))
        for f in ("fM", "fAnt", "fAA", "f12"):
            out.append(f"    & ({cel(med, f)})")
        out.append("    & " + s("</div>"))
    return out + ["    & " + s("</div>")]


V_CART = ['VAR nc = CALCULATE(COUNTROWS(fCarteira), fCarteira[SITUACAO] = "Crítico") + 0',
          'VAR na = CALCULATE(COUNTROWS(fCarteira), fCarteira[SITUACAO] = "Atenção") + 0',
          'VAR nd = CALCULATE(COUNTROWS(fCarteira), fCarteira[SITUACAO] = "Em dia") + 0',
          "VAR atr = CALCULATE(COUNTROWS(dCliente), dCliente[MESES_ATRASO] > 0) + 0",
          "VAR fat = SUM(fCarteira[RECEITA_BRUTA_12M])"]
V_CLI = JANELAS + ["VAR fat = CALCULATE([Faturamento], fM)", "VAR fatAA = CALCULATE([Faturamento], fAA)", "VAR luc = CALCULATE([Lucro líquido], fM)",
                   "VAR luc3 = CALCULATE([Lucro líquido], f3)", "VAR mar3 = CALCULATE([Margem líquida], f3)", "VAR disp = [Disponível]",
                   "VAR liq = [Liquidez corrente]", "VAR carga = CALCULATE([Carga tributária], fM)"]

MEDIDAS["Painel"] = [(n, None, linhas) for n, linhas in {
    "HTML Cabeçalho Carteira": cabecalho("Carteira de clientes"),
    "HTML Cabeçalho Cliente": cabecalho("Painel do cliente"),
    "HTML Cabeçalho Resultado": cabecalho("Resultado e caixa"),
    "HTML Cartões Carteira": cartoes(V_CART, [
        ("Precisam de conversa agora", 'FORMAT(nc, "0")', s("clientes com alerta crítico"), "nc > 0"),
        ("Para acompanhar", 'FORMAT(na, "0")', s("clientes com ponto de atenção"), None),
        ("Em dia", 'FORMAT(nd, "0")', s("sem alerta neste balancete"), None),
        ("Balancetes atrasados", 'FORMAT(atr, "0")', s("painel desses clientes está desatualizado"), "atr > 0"),
        ("Faturamento da carteira", BRL("fat"), s("somado, últimos 12 meses"), None)]),
    "HTML Tabela Carteira": tabela_carteira(498),
    "HTML Pauta": pauta(498),
    "HTML Cartões Cliente": cartoes(V_CLI, [
        ("Faturamento do mês", BRL("fat"), 'IF(ISBLANK(fatAA), "primeiro ano de histórico", FORMAT(DIVIDE(fat, fatAA) - 1, "+0.0%;-0.0%") & " sobre um ano antes")', None),
        ("Lucro do mês", BRL("luc"), BRL("luc3") + ' & " no trimestre"', "luc < 0"),
        ("Sobra de cada R$ 100", '"R$ " & FORMAT(mar3 * 100, "0.0")', s("margem líquida do trimestre"), "mar3 < 0"),
        ("Dinheiro disponível", BRL("disp"), s("caixa e aplicações"), None),
        ("Para cada R$ 1 a pagar", '"R$ " & FORMAT(liq, "0.00")', s("em caixa, a receber e estoque"), "liq < 1"),
        ("Impostos", 'FORMAT(carga, "0.0%")', s("do faturamento do mês"), None)]),
    "HTML Alertas Cliente": alertas_cliente(244),
    "HTML Cada 100": cada_cem(244),
    "HTML Simples": simples(254),
    "HTML DRE": dre(622),
}.items()]
TABELA_DA_MEDIDA = {m[0]: t for t, ms in MEDIDAS.items() for m in ms}

H = "html"
SEG_CLI = [("slicer", 880, 6, 190, 58, None, {"Values": ["dCliente.CLIENTE|Cliente"]}), ("slicer", 1078, 6, 186, 58, None, {"Values": ["dMes.MES_ROTULO|Mês de referência"]})]
PAGINAS = [
    ("carteira", "Carteira", [
        (H, 0, 0, 1280, 70, None, "HTML Cabeçalho Carteira"),
        ("slicer", 880, 6, 190, 58, None, {"Values": ["dCliente.RESPONSAVEL|Responsável"]}), ("slicer", 1078, 6, 186, 58, None, {"Values": ["dCliente.REGIME|Regime"]}),
        (H, 16, 82, 1248, 112, None, "HTML Cartões Carteira"),
        (H, 16, 206, 830, 498, None, "HTML Tabela Carteira"), (H, 858, 206, 406, 498, None, "HTML Pauta")]),
    ("cliente", "Cliente", [
        (H, 0, 0, 1280, 70, None, "HTML Cabeçalho Cliente")] + SEG_CLI + [
        (H, 16, 82, 1248, 112, None, "HTML Cartões Cliente"),
        (H, 16, 206, 618, 244, None, "HTML Alertas Cliente"), (H, 646, 206, 618, 244, None, "HTML Cada 100"),
        ("lineChart", 16, 462, 780, 242, "Faturamento e lucro, mês a mês", {"Category": ["dMes.MES_ROTULO|Mês"], "Y": ["Faturamento", "Lucro líquido"]}),
        (H, 808, 462, 456, 254, None, "HTML Simples")]),
    ("resultado", "Resultado e caixa", [
        (H, 0, 0, 1280, 70, None, "HTML Cabeçalho Resultado")] + SEG_CLI + [
        (H, 16, 82, 760, 622, None, "HTML DRE"),
        ("clusteredColumnChart", 788, 82, 476, 300, "Margem líquida por mês", {"Category": ["dMes.MES_ROTULO|Mês"], "Y": ["Margem líquida"]}),
        ("lineChart", 788, 394, 476, 310, "Dinheiro disponível e dívida bancária", {"Category": ["dMes.MES_ROTULO|Mês"], "Y": ["Disponível", "Dívida bancária"]})]),
]


# ----------------------------------------------------------------------------
def escrever(caminho, texto):
    caminho.parent.mkdir(parents=True, exist_ok=True)
    caminho.write_text(texto, encoding="utf-8", newline="\n")


def js(caminho, obj):
    escrever(caminho, json.dumps(obj, ensure_ascii=False, indent=2) + "\n")


def q(n):
    return n if n.replace("_", "").isalnum() else "'" + n.replace("'", "''") + "'"


def medidas_tmdl(nome):
    L = []
    for med, fmt, dax in MEDIDAS.get(nome, []):
        if isinstance(dax, list):
            L.append(f"\tmeasure {q(med)} =")
            L += ["\t\t\t" + x for x in dax]
        else:
            L.append(f"\tmeasure {q(med)} = {dax}")
        if fmt:
            L.append(f"\t\tformatString: {fmt}")
        L += [f"\t\tlineageTag: {tag(nome, 'm', med)}", ""]
    return L


def tmdl_tabela(nome):
    arq, cols, ocultas, ordenar = TABELAS[nome]
    L = [f"table {nome}", f"\tlineageTag: {tag(nome)}", ""] + medidas_tmdl(nome)
    for c, t in cols.items():
        L += [f"\tcolumn {c}", f"\t\tdataType: {T[t][0]}"]
        if ocultas == "*" or c in ocultas:
            L.append("\t\tisHidden")
        if t == "date":
            L.append("\t\tformatString: dd/mm/yyyy")
        elif t == "int":
            L.append("\t\tformatString: 0")
        L += [f"\t\tlineageTag: {tag(nome, 'c', c)}", "\t\tsummarizeBy: none", f"\t\tsourceColumn: {c}"]
        if c in ordenar:
            L.append(f"\t\tsortByColumn: {ordenar[c]}")
        L += ["", "\t\tannotation SummarizationSetBy = Automatic", ""]
        if t == "date":
            L += ["\t\tannotation UnderlyingDateTimeDataType = Date", ""]
    lista = ", ".join(f'"{c}"' for c in cols)
    tipos = ", ".join(f'{{"{c}", {T[t][1]}}}' for c, t in cols.items())
    L += [f"\tpartition {nome} = m", "\t\tmode: import", "\t\tsource =", "\t\t\t\tlet",
          f'\t\t\t\t    Fonte = Parquet.Document(File.Contents(PastaExports & "{arq}.parquet")),',
          f"\t\t\t\t    Colunas = Table.SelectColumns(Fonte, {{{lista}}}),",
          f"\t\t\t\t    Tipos = Table.TransformColumnTypes(Colunas, {{{tipos}}})", "\t\t\t\tin", "\t\t\t\t    Tipos", "",
          "\tannotation PBI_ResultType = Table", ""]
    return "\n".join(L)


def tmdl_painel():
    L = ["table Painel", f"\tlineageTag: {tag('Painel')}", ""] + medidas_tmdl("Painel")
    L += ["\tcolumn Bloco", "\t\tdataType: string", "\t\tisHidden", f"\t\tlineageTag: {tag('Painel', 'c', 'Bloco')}", "\t\tsummarizeBy: none",
          "\t\tsourceColumn: Bloco", "", "\t\tannotation SummarizationSetBy = Automatic", "", "\tpartition Painel = m", "\t\tmode: import", "\t\tsource =",
          "\t\t\t\tlet", '\t\t\t\t    Fonte = #table(type table [Bloco = text], {{"HTML"}})', "\t\t\t\tin", "\t\t\t\t    Fonte", "", "\tannotation PBI_ResultType = Table", ""]
    return "\n".join(L)


def modelo():
    d = PBI / f"{NOME}.SemanticModel"
    js(d / ".platform", {"$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
                         "metadata": {"type": "SemanticModel", "displayName": NOME}, "config": {"version": "2.0", "logicalId": tag("modelo")}})
    js(d / "definition.pbism", {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
                                "version": "4.2", "settings": {}})
    escrever(d / "definition" / "database.tmdl", "database\n\tcompatibilityLevel: 1606\n\n")
    exp, pasta = d / "definition" / "expressions.tmdl", PASTA_PADRAO
    if exp.exists():                                         # preserva o caminho ja configurado nesta maquina
        m = re.search(r'PastaExports = "([^"]*)"', exp.read_text(encoding="utf-8"))
        pasta = m.group(1) if m else pasta
    escrever(exp, f'expression PastaExports = "{pasta}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n'
                  f"\tlineageTag: {tag('PastaExports')}\n\n\tannotation PBI_ResultType = Text\n\n")
    todas = list(TABELAS) + ["Painel"]
    escrever(d / "definition" / "model.tmdl", "\n".join([
        "model Model", "\tculture: pt-BR", "\tdefaultPowerBIDataSourceVersion: powerBI_V3", "\tdiscourageImplicitMeasures", "\tsourceQueryCulture: pt-BR",
        "\tdataAccessOptions", "\t\tlegacyRedirects", "\t\treturnErrorValuesAsNull", "", "annotation __PBI_TimeIntelligenceEnabled = 0", "",
        f"annotation PBI_QueryOrder = {json.dumps(['PastaExports'] + todas)}", "", 'annotation PBI_ProTooling = ["DevMode"]', ""]
        + [f"ref table {t}" for t in todas] + ["", ""]))
    escrever(d / "definition" / "relationships.tmdl", "\n".join(f"relationship {tag('rel', a, b)}\n\tfromColumn: {a}\n\ttoColumn: {b}\n" for a, b in RELACOES) + "\n")
    for t in TABELAS:
        escrever(d / "definition" / "tables" / f"{t}.tmdl", tmdl_tabela(t))
    escrever(d / "definition" / "tables" / "Painel.tmdl", tmdl_painel())


def lit(v):
    return {"expr": {"Literal": {"Value": v}}}


def cor(v):
    return {"solid": {"color": lit("'" + v + "'")}}


def campo(spec):
    nome, _, rot = spec.partition("|")
    if "." in nome and nome.split(".")[0] in TABELAS:
        tab, col = nome.split(".", 1)
        f = {"Column": {"Expression": {"SourceRef": {"Entity": tab}}, "Property": col}}
    else:
        tab, col = TABELA_DA_MEDIDA[nome], nome
        f = {"Measure": {"Expression": {"SourceRef": {"Entity": tab}}, "Property": col}}
    p = {"field": f, "queryRef": f"{tab}.{col}", "nativeQueryRef": rot or col}
    if rot:
        p["displayName"] = rot
    return p


def visual(pag, i, spec):
    tipo, x, y, w, h, titulo, papeis = spec
    nome = hashlib.sha1(f"{pag}/{i}/{tipo}/{papeis}".encode()).hexdigest()[:20]
    objs = {}
    cont = {"title": [{"properties": {"show": lit("true"), "text": lit("'" + titulo.replace("'", "''") + "'")} if titulo else {"show": lit("false")}}]}
    if tipo == "html":
        tipo = VISUAL_HTML
        query = {"queryState": {"content": {"projections": [campo(papeis)]}}}
        zero = {k: lit("0D") for k in ("top", "bottom", "left", "right")}
        cont.update({"background": [{"properties": {"show": lit("false")}}], "border": [{"properties": {"show": lit("false")}}],
                     "padding": [{"properties": zero}], "visualHeader": [{"properties": {"show": lit("false")}}]})
    else:
        query = {"queryState": {papel: {"projections": [campo(c) for c in campos]} for papel, campos in papeis.items()}}
        if tipo == "slicer":
            objs = {"data": [{"properties": {"mode": lit("'Dropdown'")}}],
                    "header": [{"properties": {"show": lit("true"), "fontColor": cor(C["tinta2"]), "textSize": lit("8D")}}],
                    "items": [{"properties": {"fontColor": cor(C["tinta"]), "background": cor(C["painel2"]), "textSize": lit("9D")}}]}
            cont.update({"background": [{"properties": {"show": lit("true"), "color": cor(C["painel2"]), "transparency": lit("0D")}}],
                         "border": [{"properties": {"show": lit("true"), "color": cor(C["menta2"]), "radius": lit("9D")}}]})
        elif tipo == "clusteredColumnChart":
            objs = {"labels": [{"properties": {"show": lit("false")}}], "legend": [{"properties": {"show": lit("false")}}]}
        elif tipo == "lineChart":
            objs = {"legend": [{"properties": {"show": lit("true"), "position": lit("'Top'"), "labelColor": cor(C["tinta2"])}}]}
    return nome, {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.14.0/schema.json",
                  "name": nome, "position": {"x": x, "y": y, "z": (i + 1) * 1000, "height": h, "width": w, "tabOrder": (i + 1) * 1000},
                  "visual": {"visualType": tipo, "query": query, "objects": objs, "visualContainerObjects": cont, "drillFilterOtherVisuals": True}}


def eixo():
    return {"valueAxis": [{"gridlineColor": {"solid": {"color": C["linha"]}}, "labelColor": {"solid": {"color": C["tinta2"]}}, "showAxisTitle": False}],
            "categoryAxis": [{"labelColor": {"solid": {"color": C["tinta2"]}}, "showAxisTitle": False}]}


TEMA = {
    "name": "Lastro Contabil", "dataColors": [C["menta"], C["tinta"], C["laranja"], C["ambar"], C["lilas"], C["rosa"]],
    "background": C["painel"], "foreground": C["tinta"], "tableAccent": C["menta"], "foregroundNeutralSecondary": C["tinta2"],
    "foregroundNeutralTertiary": C["mudo"], "backgroundLight": C["painel2"], "backgroundNeutral": C["linha"],
    "good": C["menta"], "neutral": C["ambar"], "bad": C["coral"],
    "textClasses": {"callout": {"fontFace": "Segoe UI Semibold", "color": C["tinta"]}, "title": {"fontFace": "Segoe UI Semibold", "fontSize": 11, "color": C["tinta"]},
                    "header": {"fontFace": "Segoe UI Semibold", "fontSize": 10, "color": C["tinta"]}, "label": {"fontFace": "Segoe UI", "fontSize": 9, "color": C["tinta2"]}},
    "visualStyles": {
        "*": {"*": {"background": [{"show": True, "color": {"solid": {"color": C["painel"]}}, "transparency": 0}],
                    "border": [{"show": True, "color": {"solid": {"color": C["linha"]}}, "radius": 16}],
                    "title": [{"fontColor": {"solid": {"color": C["tinta"]}}, "fontSize": 11}], "dropShadow": [{"show": False}]}},
        "lineChart": {"*": eixo()}, "clusteredColumnChart": {"*": eixo()},
        "page": {"*": {"background": [{"color": {"solid": {"color": C["fundo"]}}, "transparency": 0}],
                       "outspace": [{"color": {"solid": {"color": C["fundo"]}}, "transparency": 0}]}}},
}


def relatorio():
    d = PBI / f"{NOME}.Report"
    js(PBI / f"{NOME}.pbip", {"$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json", "version": "1.0",
                              "artifacts": [{"report": {"path": f"{NOME}.Report"}}], "settings": {"enableAutoRecovery": True}})
    js(d / ".platform", {"$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
                         "metadata": {"type": "Report", "displayName": NOME}, "config": {"version": "2.0", "logicalId": tag("relatorio")}})
    js(d / "definition.pbir", {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
                               "version": "4.0", "datasetReference": {"byPath": {"path": f"../{NOME}.SemanticModel"}}})
    js(d / "StaticResources" / "RegisteredResources" / "LastroContabil.json", TEMA)
    ver = {"visual": "2.11.0", "report": "3.4.0", "page": "2.3.1"}
    js(d / "definition" / "report.json", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.3.0/schema.json",
        "themeCollection": {"baseTheme": {"name": "CY26SU07", "reportVersionAtImport": ver, "type": "SharedResources"},
                            "customTheme": {"name": "LastroContabil.json", "reportVersionAtImport": ver, "type": "RegisteredResources"}},
        "publicCustomVisuals": [VISUAL_HTML],
        "resourcePackages": [
            {"name": "SharedResources", "type": "SharedResources", "items": [{"name": "CY26SU07", "path": "BaseThemes/CY26SU07.json", "type": "BaseTheme"}]},
            {"name": "RegisteredResources", "type": "RegisteredResources", "items": [{"name": "LastroContabil.json", "path": "LastroContabil.json", "type": "CustomTheme"}]}],
        "settings": {"useStylableVisualContainerHeader": True, "exportDataMode": "AllowSummarized", "defaultDrillFilterOtherVisuals": True,
                     "allowChangeFilterTypes": True, "useEnhancedTooltips": True, "useDefaultAggregateDisplayName": True}})
    js(d / "definition" / "version.json", {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})
    js(d / "definition" / "pages" / "pages.json", {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json",
                                                   "pageOrder": [p[0] for p in PAGINAS], "activePageName": PAGINAS[0][0]})
    for pid, titulo, visuais in PAGINAS:
        pasta = d / "definition" / "pages" / pid
        if (pasta / "visuals").exists():
            for v in (pasta / "visuals").iterdir():
                shutil.rmtree(v, ignore_errors=True)
                if v.exists():
                    try:
                        v.rmdir()
                    except OSError:
                        pass
        js(pasta / "page.json", {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json",
                                 "name": pid, "displayName": titulo, "displayOption": "FitToPage", "height": 720, "width": 1280,
                                 "objects": {k: [{"properties": {"color": cor(C["fundo"]), "transparency": lit("0D")}}] for k in ("background", "outspace")}})
        for i, spec in enumerate(visuais):
            nome, obj = visual(pid, i, spec)
            js(pasta / "visuals" / nome / "visual.json", obj)
    print(f"{NOME}: {len(TABELAS) + 1} tabelas, {sum(len(m) for m in MEDIDAS.values())} medidas, {len(PAGINAS)} paginas, {sum(len(p[2]) for p in PAGINAS)} visuais")


if __name__ == "__main__":
    modelo()
    relatorio()

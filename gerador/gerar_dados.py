"""Gera a base ficticia do escritorio Lastro Contábil e das suas 20 empresas clientes.

Tudo aqui e inventado: escritorio, empresas, CNPJs (com digitos invalidos de proposito), cidades e numeros.
A semente e fixa, entao a base sai sempre igual.

Cada empresa e simulada mes a mes por lancamentos em partidas dobradas (todo debito tem um credito),
entao o balancete fecha: ativo = passivo + patrimonio liquido + resultado do exercicio.
Cada empresa tem um "cenario" (crescendo, perto do teto do Simples, em prejuizo, endividada...)
para o painel ter o que mostrar.

Saida, no formato em que o escritorio receberia:

  dados/brutos/cadastro_clientes.xlsx        carteira do escritorio
  dados/brutos/balancetes/<cliente>/         1 planilha por mes, em 2 layouts (2 sistemas contabeis diferentes)
  dados/brutos/folha_resumo.csv              funcionarios e total da folha por cliente e mes

As regras de imposto estao SIMPLIFICADAS para a demonstracao (ver config/). Nao usar como base fiscal.

Uso: python gerador/gerar_dados.py
"""
import csv
from datetime import date
from pathlib import Path

import numpy as np
import openpyxl

RAIZ = Path(__file__).resolve().parents[1]
BRUTOS = RAIZ / "dados" / "brutos"
CONFIG = RAIZ / "config"
rng = np.random.default_rng(2026)

MESES = [date(2024 + (10 + i) // 12, (10 + i) % 12 + 1, 1) for i in range(24)]        # nov/2024 a out/2026

# codigo, descricao, grupo, natureza (D devedora / C credora)
PLANO = [
    ("1.1.01", "Caixa e bancos", "Ativo circulante", "D"), ("1.1.02", "Aplicações financeiras", "Ativo circulante", "D"),
    ("1.1.03", "Clientes a receber", "Ativo circulante", "D"), ("1.1.04", "Estoques", "Ativo circulante", "D"),
    ("1.2.01", "Imobilizado", "Ativo não circulante", "D"), ("1.2.02", "(-) Depreciação acumulada", "Ativo não circulante", "C"),
    ("2.1.01", "Fornecedores", "Passivo circulante", "C"), ("2.1.02", "Salários e encargos a pagar", "Passivo circulante", "C"),
    ("2.1.03", "Impostos a recolher", "Passivo circulante", "C"), ("2.1.04", "Empréstimos de curto prazo", "Passivo circulante", "C"),
    ("2.2.01", "Empréstimos de longo prazo", "Passivo não circulante", "C"),
    ("2.3.01", "Capital social", "Patrimônio líquido", "C"), ("2.3.02", "Lucros ou prejuízos acumulados", "Patrimônio líquido", "C"),
    ("3.1.01", "Receita de venda de mercadorias", "Receita bruta", "C"), ("3.1.02", "Receita de prestação de serviços", "Receita bruta", "C"),
    ("3.2.01", "(-) Impostos sobre a receita", "Deduções da receita", "D"), ("3.2.02", "(-) Devoluções e abatimentos", "Deduções da receita", "D"),
    ("4.1.01", "Custo das mercadorias vendidas", "Custos", "D"), ("4.1.02", "Custo dos serviços prestados", "Custos", "D"),
    ("5.1.01", "Salários", "Despesas com pessoal", "D"), ("5.1.02", "Encargos sociais", "Despesas com pessoal", "D"),
    ("5.1.03", "Pró-labore", "Despesas com pessoal", "D"),
    ("5.2.01", "Aluguel", "Despesas administrativas", "D"), ("5.2.02", "Energia, água e telecomunicações", "Despesas administrativas", "D"),
    ("5.2.03", "Serviços de terceiros", "Despesas administrativas", "D"), ("5.2.04", "Marketing e vendas", "Despesas administrativas", "D"),
    ("5.2.05", "Outras despesas administrativas", "Despesas administrativas", "D"),
    ("5.3.01", "Depreciação", "Depreciação", "D"),
    ("5.4.01", "Despesas financeiras", "Resultado financeiro", "D"), ("5.4.02", "Receitas financeiras", "Resultado financeiro", "C"),
    ("5.5.01", "IRPJ e CSLL", "Imposto sobre o lucro", "D"),
]
RESULTADO = [c[0] for c in PLANO if c[0][0] in "345"]

# Simples Nacional: anexo -> faixas (limite da receita em 12 meses, aliquota nominal, parcela a deduzir)
SIMPLES = {
    "I": [(180000, .04, 0), (360000, .073, 5940), (720000, .095, 13860), (1800000, .107, 22500), (3600000, .143, 87300), (4800000, .19, 378000)],
    "II": [(180000, .045, 0), (360000, .078, 5940), (720000, .10, 13860), (1800000, .112, 22500), (3600000, .147, 85500), (4800000, .30, 720000)],
    "III": [(180000, .06, 0), (360000, .112, 9360), (720000, .135, 17640), (1800000, .16, 35640), (3600000, .21, 125640), (4800000, .33, 648000)],
}
SAZ = {
    "plana": [1] * 12,
    "natal": [.85, .82, .92, .95, 1.05, .98, .95, .98, .97, 1.02, 1.18, 1.55],
    "verao": [1.45, 1.35, 1.15, .95, .75, .62, .65, .72, .88, 1.02, 1.2, 1.4],
    "escolar": [1.15, 1.1, 1.02, 1, 1, .98, .9, 1, 1, 1, .98, .95],
    "safra": [.8, .85, 1.1, 1.3, 1.35, 1.1, .9, .85, .95, 1.05, 1.0, .85],
}

# Carteira: cada empresa com seu cenario. receita = mensal inicial; g = crescimento ao mes.
# cmv = custo sobre a receita; pessoal e fixas = despesa mensal inicial; prazo = % vendido a prazo; cobra = % do contas a receber que entra no mes
def emp(cod, nome, atividade, cidade, regime, anexo, cenario, receita, g=0.004, saz="plana", serv=0.0, cmv=0.55, pessoal=0.14, fixas=0.10,
        prazo=0.4, cobra=0.85, estoque=35, divida=0.0, sistema="A", inicio=0, atraso=0, **extra):
    return dict(cod=cod, nome=nome, atividade=atividade, cidade=cidade, regime=regime, anexo=anexo, cenario=cenario, receita=receita, g=g, saz=saz,
                serv=serv, cmv=cmv, pessoal=pessoal, fixas=fixas, prazo=prazo, cobra=cobra, estoque=estoque, divida=divida, sistema=sistema,
                inicio=inicio, atraso=atraso, **extra)


EMPRESAS = [
    emp("C001", "Mercado Bom Preço Ltda", "Supermercado de bairro", "Vale Sereno", "Simples Nacional", "I", "Saudável e crescendo", 185000, g=.012, cmv=.70, pessoal=.09, fixas=.06, prazo=.15, estoque=28),
    emp("C002", "Padaria Trigo Dourado Ltda", "Padaria e confeitaria", "Vale Sereno", "Simples Nacional", "I", "Perto de mudar de faixa no Simples", 131000, g=.005, cmv=.48, pessoal=.20, fixas=.11, prazo=.05, estoque=12),
    emp("C003", "Moda Íris Boutique Ltda", "Comércio de vestuário", "Porto das Pedras", "Simples Nacional", "I", "Sazonal, com pico no fim do ano", 72000, saz="natal", cmv=.50, pessoal=.15, fixas=.13, prazo=.6, estoque=75, sistema="B"),
    emp("C004", "Sorveteria Polo Sul Ltda", "Sorveteria", "Porto das Pedras", "Simples Nacional", "I", "Sazonal, com prejuízo no inverno", 58000, saz="verao", cmv=.38, pessoal=.21, fixas=.16, prazo=.02, estoque=15, sistema="B"),
    emp("C005", "Farmácia Saúde Já Ltda", "Farmácia", "Vale Sereno", "Simples Nacional", "I", "Perto do limite de R$ 4,8 milhões", 333000, g=.007, cmv=.69, pessoal=.09, fixas=.05, prazo=.25, estoque=45),
    emp("C006", "Pet Shop Amigo Fiel Ltda", "Pet shop e banho e tosa", "Lagoa Clara", "Simples Nacional", "I", "Empresa nova, ainda no vermelho", 34000, g=.045, serv=.3, cmv=.54, pessoal=.30, fixas=.28, prazo=.1, estoque=40, inicio=14, sistema="B"),
    emp("C007", "Ótica Visão Clara Ltda", "Ótica", "Lagoa Clara", "Simples Nacional", "I", "Estoque inchado", 64000, cmv=.42, pessoal=.17, fixas=.13, prazo=.7, estoque=95, deriva_estoque=5.0),
    emp("C008", "Restaurante Fogão de Minas Ltda", "Restaurante", "Vale Sereno", "Simples Nacional", "I", "Margem caindo: custo subindo mais que a venda", 96000, g=.003, cmv=.36, pessoal=.26, fixas=.14, prazo=.3, estoque=10, deriva_cmv=.0045),
    emp("C009", "Móveis Arte & Madeira Ltda", "Fábrica de móveis", "Campo Largo", "Simples Nacional", "II", "Crescendo rápido, com caixa apertado", 120000, g=.028, cmv=.58, pessoal=.15, fixas=.08, prazo=.9, cobra=.42, estoque=70, divida=.9),
    emp("C010", "Gráfica Impressa Ltda", "Gráfica", "Campo Largo", "Simples Nacional", "II", "Receita em queda", 88000, g=-.016, cmv=.50, pessoal=.20, fixas=.13, prazo=.6, estoque=30, sistema="B"),
    emp("C011", "TecnoNorte Sistemas Ltda", "Desenvolvimento de software", "Porto das Pedras", "Simples Nacional", "III", "Saudável, margem alta", 92000, g=.014, serv=1, cmv=.30, pessoal=.20, fixas=.09, prazo=.9, cobra=.9, estoque=0),
    emp("C012", "Studio Kairós Pilates Ltda", "Estúdio de pilates", "Lagoa Clara", "Simples Nacional", "III", "Folha pesada", 41000, serv=1, cmv=.38, pessoal=.36, fixas=.17, prazo=.1, estoque=0, sistema="B"),
    emp("C013", "Clínica Vida Plena Ltda", "Clínica médica", "Vale Sereno", "Simples Nacional", "III", "Despesa administrativa disparando", 148000, g=.005, serv=1, cmv=.34, pessoal=.18, fixas=.12, prazo=.5, estoque=0, salto_fixas=(21, 1.75)),
    emp("C014", "Auto Center Rota 7 Ltda", "Oficina mecânica", "Campo Largo", "Simples Nacional", "III", "Inadimplência subindo", 67000, serv=.65, cmv=.45, pessoal=.19, fixas=.12, prazo=.75, cobra=.8, estoque=40, deriva_cobra=-.022),
    emp("C015", "Colégio Saber Vivo Ltda", "Escola de ensino fundamental", "Vale Sereno", "Lucro Presumido", None, "Estável, com balancete atrasado", 310000, saz="escolar", serv=1, cmv=.42, pessoal=.20, fixas=.11, prazo=.95, cobra=.88, estoque=0, atraso=2),
    emp("C016", "Construtora Alicerce Ltda", "Construção civil", "Porto das Pedras", "Lucro Presumido", None, "Endividada", 420000, g=.002, serv=.8, cmv=.66, pessoal=.12, fixas=.07, prazo=.9, cobra=.5, estoque=60, divida=3.4),
    emp("C017", "Transportes Rodamil Ltda", "Transporte rodoviário de cargas", "Campo Largo", "Lucro Presumido", None, "Saindo do prejuízo", 265000, g=.009, serv=1, cmv=.74, pessoal=.11, fixas=.07, prazo=.9, cobra=.75, estoque=0, divida=1.6, deriva_cmv=-.006, sistema="B"),
    emp("C018", "Agropecuária Campo Forte Ltda", "Comércio de insumos agrícolas", "Campo Largo", "Lucro Presumido", None, "Sazonal pela safra", 390000, saz="safra", cmv=.735, pessoal=.07, fixas=.05, prazo=.8, cobra=.6, estoque=80),
    emp("C019", "Metalúrgica Forja Sul Ltda", "Indústria metalúrgica", "Lagoa Clara", "Lucro Real", None, "Prejuízo recorrente", 610000, g=-.004, cmv=.78, pessoal=.11, fixas=.08, prazo=.9, cobra=.6, estoque=65, divida=2.2),
    emp("C020", "Distribuidora Horizonte Azul Ltda", "Distribuidora de alimentos", "Lagoa Clara", "Lucro Real", None, "Grande e saudável, margem fina", 1150000, g=.007, cmv=.83, pessoal=.05, fixas=.035, prazo=.9, cobra=.8, estoque=30, divida=.5),
]
RESPONSAVEIS = ["Aline", "Bruno", "Cecília", "Diego"]


def aliquota_simples(anexo, rbt12):
    """Aliquota efetiva = (RBT12 x aliquota nominal - parcela a deduzir) / RBT12."""
    faixas = SIMPLES[anexo]
    for limite, aliq, ded in faixas:
        if rbt12 <= limite:
            return max(0.0, (rbt12 * aliq - ded) / rbt12) if rbt12 > 0 else faixas[0][1]
    limite, aliq, ded = faixas[-1]
    return (rbt12 * aliq - ded) / rbt12


def simular(e):
    """Devolve uma lista de balancetes mensais: {conta: (saldo anterior, debitos, creditos, saldo final)}, saldos com sinal devedor."""
    saldo = {c[0]: 0.0 for c in PLANO}
    r0 = e["receita"]
    # balanco de abertura
    ab = {"1.1.01": r0 * rng.uniform(.5, .9), "1.1.03": r0 * e["prazo"] / e["cobra"], "1.1.04": r0 * e["cmv"] * e["estoque"] / 30,
          "1.2.01": r0 * rng.uniform(1.8, 3.2), "1.2.02": -r0 * rng.uniform(.4, .9), "2.1.01": -r0 * e["cmv"] * (1 - e["serv"]) * .9,
          "2.1.02": -r0 * e["pessoal"] * .8, "2.1.03": -r0 * .06, "2.1.04": -r0 * e["divida"] * .3, "2.2.01": -r0 * e["divida"] * .7}
    saldo.update(ab)
    pl = -sum(saldo.values())
    saldo["2.3.01"] = min(0.0, pl * .6) if pl < 0 else 0.0
    saldo["2.3.02"] = pl - saldo["2.3.01"]
    historico, receitas = [], []
    rec_prev = [r0 * SAZ[e["saz"]][(MESES[0].month - 1 - k) % 12] for k in range(12, 0, -1)]       # 12 meses anteriores, para o RBT12
    pessoal, fixas, cmv, cobra, est_dias = r0 * e["pessoal"], r0 * e["fixas"], e["cmv"], e["cobra"], e["estoque"]
    for t, mes in enumerate(MESES):
        if t < e["inicio"]:
            historico.append(None)
            continue
        mov = {c[0]: [0.0, 0.0] for c in PLANO}
        ant = dict(saldo)

        def lanca(debito, credito, valor):
            if valor <= 0:
                return
            mov[debito][0] += valor
            mov[credito][1] += valor
            saldo[debito] += valor
            saldo[credito] -= valor

        k = t - e["inicio"]
        receita = r0 * (1 + e["g"]) ** k * SAZ[e["saz"]][mes.month - 1] * rng.normal(1, .035)
        cmv = min(.95, cmv + e.get("deriva_cmv", 0) * (1 if k else 0))
        cobra = float(np.clip(cobra + e.get("deriva_cobra", 0) * (1 if k else 0), .25, .98))
        est_dias += e.get("deriva_estoque", 0) * (1 if k else 0)
        pessoal *= 1.004 + (.004 if mes.month == 5 else 0)                       # reajuste anual em maio
        fixas *= 1.0035
        if e.get("salto_fixas") and t == e["salto_fixas"][0]:
            fixas *= e["salto_fixas"][1]
        serv, merc = receita * e["serv"], receita * (1 - e["serv"])

        # pagamentos do que ficou do mes anterior
        lanca("2.1.01", "1.1.01", -ant["2.1.01"] * .92)
        lanca("2.1.02", "1.1.01", -ant["2.1.02"])
        lanca("2.1.03", "1.1.01", -ant["2.1.03"])
        lanca("1.1.01", "1.1.03", ant["1.1.03"] * cobra)
        # vendas
        for conta, valor in (("3.1.01", merc), ("3.1.02", serv)):
            lanca("1.1.03", conta, valor * e["prazo"])
            lanca("1.1.01", conta, valor * (1 - e["prazo"]))
        lanca("3.2.02", "1.1.03", merc * rng.uniform(.004, .014))
        # impostos sobre a receita (simplificado)
        rbt12 = sum((rec_prev + receitas)[-12:])
        if e["regime"] == "Simples Nacional":
            imposto = receita * aliquota_simples(e["anexo"], rbt12)
        elif e["regime"] == "Lucro Presumido":
            imposto = receita * (.0365 + (.05 if e["serv"] > .5 else .04))
        else:
            imposto = receita * .0925 * .45                                       # nao cumulativo, liquido de creditos
        lanca("3.2.01", "2.1.03", imposto)
        # custo: mercadoria sai do estoque; servico e mao de obra direta
        custo_merc = merc * cmv
        compra = max(0.0, custo_merc + (custo_merc * est_dias / 30 - saldo["1.1.04"]) * .5)
        lanca("1.1.04", "2.1.01", compra)
        lanca("4.1.01", "1.1.04", min(custo_merc, saldo["1.1.04"]))
        lanca("4.1.02", "2.1.02", serv * cmv * .7)
        lanca("4.1.02", "1.1.01", serv * cmv * .3)
        # pessoal e despesas fixas
        lanca("5.1.01", "2.1.02", pessoal * .62)
        lanca("5.1.02", "2.1.02", pessoal * (.08 if e["regime"] == "Simples Nacional" else .24))
        lanca("5.1.03", "1.1.01", pessoal * .14)
        for conta, parte in (("5.2.01", .30), ("5.2.02", .14), ("5.2.03", .22), ("5.2.04", .12), ("5.2.05", .22)):
            lanca(conta, "1.1.01", fixas * parte * rng.normal(1, .06))
        lanca("5.3.01", "1.2.02", saldo["1.2.01"] * .10 / 12)
        # financeiro
        divida = -(saldo["2.1.04"] + saldo["2.2.01"])
        lanca("5.4.01", "1.1.01", divida * .019 + receita * .004)
        lanca("1.1.02", "5.4.02", saldo["1.1.02"] * .008)
        lanca("2.1.04", "1.1.01", -saldo["2.1.04"] * .06)
        lanca("2.2.01", "2.1.04", -saldo["2.2.01"] * .03)
        # imposto sobre o lucro (regimes fora do Simples)
        if e["regime"] == "Lucro Presumido":
            lanca("5.5.01", "2.1.03", receita * (.32 if e["serv"] > .5 else .08) * .24)
        elif e["regime"] == "Lucro Real":
            lucro_mes = -sum(mov[c][0] - mov[c][1] for c in RESULTADO)
            lanca("5.5.01", "2.1.03", max(0.0, lucro_mes) * .34)
        # caixa: faltou, toma emprestimo; sobrou, aplica e distribui
        minimo = receita * .12
        if saldo["1.1.01"] < minimo:
            falta = minimo - saldo["1.1.01"]
            resgate = min(falta, saldo["1.1.02"])
            lanca("1.1.01", "1.1.02", resgate)
            lanca("1.1.01", "2.1.04", (falta - resgate) * 1.15)
        elif saldo["1.1.01"] > receita * .55:
            sobra = saldo["1.1.01"] - receita * .55
            lucro_retido = -(saldo["2.3.02"] + sum(saldo[c] for c in RESULTADO))      # acumulado + resultado do ano em curso
            if lucro_retido > 0:                                                      # distribuicao de lucros aos socios
                lanca("2.3.02", "1.1.01", min(sobra * .8, lucro_retido * .7))
            lanca("1.1.02", "1.1.01", max(0.0, saldo["1.1.01"] - receita * .55) * .7)
        # investimento ocasional
        if rng.random() < .05:
            lanca("1.2.01", "1.1.01", min(receita * rng.uniform(.1, .3), max(0.0, saldo["1.1.01"] - minimo)))
        receitas.append(receita)
        historico.append({c: (ant[c], mov[c][0], mov[c][1], saldo[c]) for c in saldo})
        assert abs(sum(saldo.values())) < 0.01, (e["cod"], mes)
        # encerramento do exercicio: o balancete de dezembro sai ANTES do encerramento; em janeiro as contas de
        # resultado abrem zeradas e o resultado do ano ja esta em lucros ou prejuizos acumulados
        if mes.month == 12:
            saldo["2.3.02"] += sum(saldo[c] for c in RESULTADO)
            for c in RESULTADO:
                saldo[c] = 0.0
    return historico


def gravar_balancete(caminho, e, mes, bal):
    """Dois layouts, como dois sistemas contabeis diferentes exportariam."""
    wb = openpyxl.Workbook()
    ws = wb.active
    nomes = {c[0]: c[1] for c in PLANO}
    if e["sistema"] == "A":
        ws.title = "Balancete"
        ws.append([f"BALANCETE DE VERIFICAÇÃO - {e['nome'].upper()}"])
        ws.append([f"Período: 01/{mes:%m/%Y} a {ultimo_dia(mes):%d/%m/%Y}", None, None, "Emissão pelo sistema contábil"])
        ws.append([])
        ws.append(["Conta", "Descrição", "Saldo anterior", "Débito", "Crédito", "Saldo atual"])
        for c, (ant, deb, cre, fim) in bal.items():
            if abs(ant) + deb + cre + abs(fim) > .005:
                ws.append([c, nomes[c], round(ant, 2), round(deb, 2), round(cre, 2), round(fim, 2)])
    else:                                                    # layout B: saldo sem sinal + coluna D/C, codigo sem pontos
        ws.title = "Plan1"
        ws.append([e["nome"]])
        ws.append([f"Balancete {mes:%m/%Y}"])
        ws.append(["Código", "Nome da conta", "Sld. Ant.", "D/C", "Débitos", "Créditos", "Sld. Final", "D/C"])
        for c, (ant, deb, cre, fim) in bal.items():
            if abs(ant) + deb + cre + abs(fim) > .005:
                ws.append([c.replace(".", ""), nomes[c].upper(), round(abs(ant), 2), "D" if ant >= 0 else "C", round(deb, 2), round(cre, 2),
                           round(abs(fim), 2), "D" if fim >= 0 else "C"])
    wb.save(caminho)


def ultimo_dia(mes):
    prox = date(mes.year + mes.month // 12, mes.month % 12 + 1, 1)
    return date.fromordinal(prox.toordinal() - 1)


def main():
    (BRUTOS / "balancetes").mkdir(parents=True, exist_ok=True)
    CONFIG.mkdir(exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Clientes"
    ws.append(["LASTRO CONTÁBIL - CARTEIRA DE CLIENTES"])
    ws.append([])
    ws.append(["Código", "Razão social", "CNPJ", "Atividade", "Cidade", "Regime tributário", "Anexo do Simples", "Sistema contábil",
               "Responsável", "Cliente desde", "Honorário mensal"])
    folha, resumo, n_arq = [], [], 0
    for i, e in enumerate(EMPRESAS):
        hist = simular(e)
        pasta = BRUTOS / "balancetes" / e["cod"]
        pasta.mkdir(exist_ok=True)
        for f in pasta.glob("*.xlsx"):
            f.unlink()
        validos = [(m, b) for m, b in zip(MESES, hist) if b is not None]
        entregues = validos[:len(validos) - e["atraso"]] if e["atraso"] else validos
        for mes, bal in entregues:
            gravar_balancete(pasta / f"balancete_{mes:%Y-%m}.xlsx", e, mes, bal)
            n_arq += 1
        for mes, bal in validos:
            sal = bal["5.1.01"][1] + bal["5.1.02"][1] + bal["5.1.03"][1] + bal["4.1.02"][1] * .35     # 35% do custo do servico e mao de obra
            folha.append([e["cod"], f"{mes:%Y-%m}", max(1, int(round(sal / rng.uniform(3300, 4100)))), round(sal, 2)])
        # CNPJ ficticio: raiz 00.9xx.xxx e digitos verificadores 00 (invalidos de proposito)
        cnpj = f"00.9{i + 10:02d}.{100 + i * 7:03d}/0001-00"
        honorario = int(round((450 + e["receita"] * .0035) / 10) * 10)
        ws.append([e["cod"], e["nome"], cnpj, e["atividade"], e["cidade"], e["regime"], e["anexo"], "Sistema " + e["sistema"],
                   RESPONSAVEIS[i % 4], validos[0][0], honorario])
        ult = [b for _, b in validos][-12:]
        rec = sum(b["3.1.01"][2] + b["3.1.02"][2] - b["3.1.01"][1] - b["3.1.02"][1] for b in ult)
        luc = sum(b[c][2] - b[c][1] for b in ult for c in RESULTADO)
        fim = validos[-1][1]
        ac = fim["1.1.01"][3] + fim["1.1.02"][3] + fim["1.1.03"][3] + fim["1.1.04"][3]
        pc = -(fim["2.1.01"][3] + fim["2.1.02"][3] + fim["2.1.03"][3] + fim["2.1.04"][3])
        div = -(fim["2.1.04"][3] + fim["2.2.01"][3])
        resumo.append((e["cod"], e["nome"][:30], e["regime"][:9] + (" " + e["anexo"] if e["anexo"] else ""), e["cenario"], rec, luc / rec,
                       ac / pc, div / (rec / len(ult)), len(entregues)))
    wb.save(BRUTOS / "cadastro_clientes.xlsx")
    with open(BRUTOS / "folha_resumo.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["CODIGO_CLIENTE", "COMPETENCIA", "FUNCIONARIOS", "TOTAL_FOLHA"])
        w.writerows([[a, b, c, str(d).replace(".", ",")] for a, b, c, d in folha])
    with open(CONFIG / "plano_de_contas.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["CONTA", "DESCRICAO", "GRUPO", "NATUREZA"])
        w.writerows(PLANO)
    with open(CONFIG / "simples_nacional.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["ANEXO", "FAIXA", "RECEITA_12M_ATE", "ALIQUOTA_NOMINAL", "PARCELA_A_DEDUZIR"])
        for anexo, faixas in SIMPLES.items():
            w.writerows([[anexo, n, lim, aliq, ded] for n, (lim, aliq, ded) in enumerate(faixas, 1)])
    (RAIZ / "docs").mkdir(exist_ok=True)
    with open(RAIZ / "docs" / "cenarios.md", "w", encoding="utf-8", newline="\n") as f:
        f.write("# Cenários da carteira fictícia\n\nEscritório Lastro Contábil, 20 empresas clientes, novembro de 2024 a outubro de 2026. "
                "Tudo inventado. Os números são dos últimos 12 meses de cada empresa.\n\n"
                "| Cliente | Empresa | Atividade | Regime | Cenário | Receita 12 m | Margem líquida | Liquidez corrente | Balancetes |\n|---|---|---|---|---|---:|---:|---:|---:|\n")
        for e, (cod, nome, reg, cen, rec, mar, liq, div, n) in zip(EMPRESAS, resumo):
            regime = e["regime"] + (f" (Anexo {e['anexo']})" if e["anexo"] else "")
            f.write(f"| {cod} | {e['nome']} | {e['atividade']} | {regime} | {cen} | R$ {rec:,.0f} | {mar:.1%} | {liq:.2f} | {n} |\n".replace(",", "."))
    print(f"{len(EMPRESAS)} empresas, {n_arq} balancetes, {len(folha)} linhas de folha\n")
    print(f"{'cod':5}{'empresa':31}{'regime':13}{'receita 12m':>12}{'margem':>8}{'liq.corr':>9}{'div/mes':>8}{'bal.':>5}  cenario")
    for cod, nome, reg, cen, rec, mar, liq, div, n in resumo:
        print(f"{cod:5}{nome:31}{reg:13}{rec:12,.0f}{mar:8.1%}{liq:9.2f}{div:8.1f}{n:5}  {cen}")


if __name__ == "__main__":
    main()

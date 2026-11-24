-- =========================================================
-- TRUSTED - BALANCO RESUMIDO POR MES
--
-- Do saldo final de cada conta patrimonial saem os blocos que o
-- empresario precisa enxergar: quanto tem em caixa, quanto tem a
-- receber, quanto deve e a quem.
--
-- O saldo vem com sinal (credor negativo), entao passivo e
-- patrimonio liquido sao invertidos para ficarem positivos.
--
-- Patrimonio liquido = capital + lucros acumulados + resultado do
-- exercicio em curso (que ainda esta nas contas de resultado).
--
-- DIFERENCA_BALANCO deve ser zero: e a prova de que o balancete fecha.
-- =========================================================

CREATE OR REPLACE TABLE balanco_mensal AS

SELECT
    b.CODIGO_CLIENTE,
    b.COMPETENCIA,
    SUM(b.SALDO_FINAL) FILTER (WHERE b.CONTA IN ('1.1.01', '1.1.02')) AS DISPONIVEL,
    COALESCE(SUM(b.SALDO_FINAL) FILTER (WHERE b.CONTA = '1.1.03'), 0) AS CLIENTES_A_RECEBER,
    COALESCE(SUM(b.SALDO_FINAL) FILTER (WHERE b.CONTA = '1.1.04'), 0) AS ESTOQUES,
    SUM(b.SALDO_FINAL) FILTER (WHERE c.GRUPO = 'Ativo circulante') AS ATIVO_CIRCULANTE,
    COALESCE(SUM(b.SALDO_FINAL) FILTER (WHERE c.GRUPO = 'Ativo não circulante'), 0) AS ATIVO_NAO_CIRCULANTE,
    SUM(b.SALDO_FINAL) FILTER (WHERE c.CLASSE = 'Ativo') AS ATIVO_TOTAL,
    COALESCE(-SUM(b.SALDO_FINAL) FILTER (WHERE b.CONTA = '2.1.01'), 0) AS FORNECEDORES,
    -SUM(b.SALDO_FINAL) FILTER (WHERE c.GRUPO = 'Passivo circulante') AS PASSIVO_CIRCULANTE,
    COALESCE(-SUM(b.SALDO_FINAL) FILTER (WHERE c.GRUPO = 'Passivo não circulante'), 0) AS PASSIVO_NAO_CIRCULANTE,
    COALESCE(-SUM(b.SALDO_FINAL) FILTER (WHERE b.CONTA IN ('2.1.04', '2.2.01')), 0) AS DIVIDA_BANCARIA,
    -SUM(b.SALDO_FINAL) FILTER (WHERE c.GRUPO = 'Patrimônio líquido' OR c.CLASSE = 'Resultado') AS PATRIMONIO_LIQUIDO,
    ROUND(SUM(b.SALDO_FINAL), 2) AS DIFERENCA_BALANCO
FROM balancete b
JOIN dim_conta c USING (CONTA)
GROUP BY 1, 2;

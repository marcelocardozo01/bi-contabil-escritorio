-- =========================================================
-- TRUSTED - DRE GERENCIAL POR MES
--
-- O balancete traz o movimento do mes em cada conta de resultado.
-- A DRE do mes e esse movimento: credito menos debito.
-- Receita sai positiva; deducao, custo e despesa saem negativos.
--
--   Receita bruta
--   (-) Deducoes            impostos sobre a receita e devolucoes
--   = Receita liquida
--   (-) Custos              mercadoria vendida e servico prestado
--   = Lucro bruto
--   (-) Despesas com pessoal, administrativas e depreciacao
--   = Resultado operacional
--   (+/-) Resultado financeiro
--   (-) IRPJ e CSLL
--   = Lucro liquido
--
-- Por que usar o movimento e nao o saldo: o saldo das contas de
-- resultado acumula no ano e zera em janeiro. O movimento e sempre
-- o resultado daquele mes, em qualquer mes do ano.
-- =========================================================

CREATE OR REPLACE TABLE dre_mensal AS

WITH mov AS (
    SELECT b.CODIGO_CLIENTE, b.COMPETENCIA, c.GRUPO, b.CONTA, b.CREDITO - b.DEBITO AS VALOR
    FROM balancete b JOIN dim_conta c USING (CONTA)
    WHERE c.CLASSE = 'Resultado'
),

grupos AS (
    SELECT CODIGO_CLIENTE, COMPETENCIA,
           COALESCE(SUM(VALOR) FILTER (WHERE GRUPO = 'Receita bruta'), 0) AS RECEITA_BRUTA,
           COALESCE(SUM(VALOR) FILTER (WHERE GRUPO = 'Deduções da receita'), 0) AS DEDUCOES,
           COALESCE(SUM(VALOR) FILTER (WHERE CONTA = '3.2.01'), 0) AS IMPOSTOS_RECEITA,
           COALESCE(SUM(VALOR) FILTER (WHERE GRUPO = 'Custos'), 0) AS CUSTOS,
           COALESCE(SUM(VALOR) FILTER (WHERE GRUPO = 'Despesas com pessoal'), 0) AS DESPESAS_PESSOAL,
           COALESCE(SUM(VALOR) FILTER (WHERE GRUPO = 'Despesas administrativas'), 0) AS DESPESAS_ADMINISTRATIVAS,
           COALESCE(SUM(VALOR) FILTER (WHERE GRUPO = 'Depreciação'), 0) AS DEPRECIACAO,
           COALESCE(SUM(VALOR) FILTER (WHERE GRUPO = 'Resultado financeiro'), 0) AS RESULTADO_FINANCEIRO,
           COALESCE(SUM(VALOR) FILTER (WHERE GRUPO = 'Imposto sobre o lucro'), 0) AS IMPOSTO_LUCRO
    FROM mov
    GROUP BY 1, 2
)

SELECT
    *,
    RECEITA_BRUTA + DEDUCOES AS RECEITA_LIQUIDA,
    RECEITA_BRUTA + DEDUCOES + CUSTOS AS LUCRO_BRUTO,
    RECEITA_BRUTA + DEDUCOES + CUSTOS + DESPESAS_PESSOAL + DESPESAS_ADMINISTRATIVAS + DEPRECIACAO AS RESULTADO_OPERACIONAL,
    RECEITA_BRUTA + DEDUCOES + CUSTOS + DESPESAS_PESSOAL + DESPESAS_ADMINISTRATIVAS + DEPRECIACAO
        + RESULTADO_FINANCEIRO + IMPOSTO_LUCRO AS LUCRO_LIQUIDO
FROM grupos;

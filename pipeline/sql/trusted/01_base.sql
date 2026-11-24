-- =========================================================
-- TRUSTED - BASE: clientes, contas e balancete conferido
--
--   dim_cliente : carteira do escritorio
--   dim_conta   : plano de contas referencial (config/plano_de_contas.csv)
--   balancete   : uma linha por cliente, mes e conta, saldo com sinal
--                 (devedor positivo, credor negativo)
--
-- Regras:
--   1. Conta que nao existe no plano referencial fica de fora e aparece
--      na checagem de qualidade (precisa de de-para antes de entrar).
--   2. MES_REFERENCIA e o ultimo mes em que algum cliente entregou
--      balancete. E contra ele que se mede o atraso de cada cliente.
-- =========================================================

CREATE OR REPLACE TABLE dim_conta AS
SELECT CONTA, DESCRICAO, GRUPO, NATUREZA,
       CASE LEFT(CONTA, 1) WHEN '1' THEN 'Ativo' WHEN '2' THEN 'Passivo e patrimônio líquido' ELSE 'Resultado' END AS CLASSE
FROM cfg_plano_de_contas;

CREATE OR REPLACE TABLE balancete AS
SELECT b.CODIGO_CLIENTE, b.COMPETENCIA, b.CONTA, b.SALDO_ANTERIOR, b.DEBITO, b.CREDITO, b.SALDO_FINAL
FROM balancete_bruto b
JOIN dim_conta c USING (CONTA);

-- meses com balancete, com o rotulo usado nos paineis (out/26)
CREATE OR REPLACE TABLE dim_mes AS
SELECT DISTINCT
    COMPETENCIA,
    STRFTIME(COMPETENCIA, '%Y-%m') AS ANO_MES,
    ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'][MONTH(COMPETENCIA)] || '/' || STRFTIME(COMPETENCIA, '%y') AS MES_ROTULO
FROM balancete;

CREATE OR REPLACE TABLE dim_cliente AS
WITH ult AS (
    SELECT CODIGO_CLIENTE, MAX(COMPETENCIA) AS ULTIMO_BALANCETE, MIN(COMPETENCIA) AS PRIMEIRO_BALANCETE, COUNT(DISTINCT COMPETENCIA) AS MESES_ENTREGUES
    FROM balancete GROUP BY 1
), ref AS (SELECT MAX(COMPETENCIA) AS MES_REFERENCIA FROM balancete)
SELECT
    c.CODIGO_CLIENTE,
    c.RAZAO_SOCIAL,
    REPLACE(c.RAZAO_SOCIAL, ' Ltda', '') AS CLIENTE,
    c.CNPJ, c.ATIVIDADE, c.CIDADE, c.REGIME, c.ANEXO_SIMPLES, c.SISTEMA, c.RESPONSAVEL, c.CLIENTE_DESDE, c.HONORARIO,
    u.PRIMEIRO_BALANCETE, u.ULTIMO_BALANCETE, u.MESES_ENTREGUES,
    r.MES_REFERENCIA,
    DATE_DIFF('month', u.ULTIMO_BALANCETE, r.MES_REFERENCIA) AS MESES_ATRASO
FROM cadastro_clientes c
JOIN ult u USING (CODIGO_CLIENTE)
CROSS JOIN ref r;

-- =========================================================
-- TRUSTED - INDICADORES POR CLIENTE E MES
--
-- Junta DRE, balanco e folha e calcula os indicadores do painel.
--
--   Margens          bruta e liquida, sobre a receita liquida
--   Liquidez corrente ativo circulante / passivo circulante
--   Liquidez seca    (ativo circulante - estoques) / passivo circulante
--   Folego de caixa  disponivel / saidas medias do mes (em meses)
--   Prazo de recebimento  clientes a receber / receita bruta x 30 (dias)
--   Dias de estoque  estoques / custo do mes x 30
--   Divida em meses de receita   divida bancaria / receita bruta media de 3 meses
--   Carga tributaria (impostos sobre a receita + IRPJ/CSLL) / receita bruta
--   Peso da folha    total da folha / receita bruta
--
-- Janelas:
--   _3M   soma dos ultimos 3 meses entregues (tira o ruido de um mes so);
--         nos 2 primeiros meses da empresa a media usa os meses que existem
--   _12M  soma dos ultimos 12 meses (RBT12 e a base do Simples Nacional)
--   _AA   o mesmo trimestre um ano antes: e contra ele que se compara,
--         para empresa sazonal nao gerar alerta falso todo inverno.
-- =========================================================

CREATE OR REPLACE TABLE indicadores_mensal AS

WITH base AS (
    SELECT
        d.*,
        b.DISPONIVEL, b.CLIENTES_A_RECEBER, b.ESTOQUES, b.ATIVO_CIRCULANTE, b.ATIVO_TOTAL, b.FORNECEDORES,
        b.PASSIVO_CIRCULANTE, b.PASSIVO_NAO_CIRCULANTE, b.DIVIDA_BANCARIA, b.PATRIMONIO_LIQUIDO,
        f.FUNCIONARIOS, f.TOTAL_FOLHA,
        ROW_NUMBER() OVER (PARTITION BY d.CODIGO_CLIENTE ORDER BY d.COMPETENCIA) AS N_MES
    FROM dre_mensal d
    JOIN balanco_mensal b USING (CODIGO_CLIENTE, COMPETENCIA)
    LEFT JOIN folha_resumo f USING (CODIGO_CLIENTE, COMPETENCIA)
),

janelas AS (
    SELECT
        *,
        SUM(RECEITA_BRUTA) OVER m3 AS RECEITA_BRUTA_3M,
        SUM(RECEITA_LIQUIDA) OVER m3 AS RECEITA_LIQUIDA_3M,
        SUM(LUCRO_BRUTO) OVER m3 AS LUCRO_BRUTO_3M,
        SUM(LUCRO_LIQUIDO) OVER m3 AS LUCRO_LIQUIDO_3M,
        SUM(DESPESAS_ADMINISTRATIVAS) OVER m3 AS DESPESAS_ADM_3M,
        SUM(RECEITA_BRUTA) OVER m12 AS RECEITA_BRUTA_12M,
        SUM(RECEITA_LIQUIDA) OVER m12 AS RECEITA_LIQUIDA_12M,
        SUM(LUCRO_LIQUIDO) OVER m12 AS LUCRO_LIQUIDO_12M,
        COUNT(*) OVER m12 AS MESES_NA_JANELA_12M
    FROM base
    WINDOW m3 AS (PARTITION BY CODIGO_CLIENTE ORDER BY COMPETENCIA ROWS BETWEEN 2 PRECEDING AND CURRENT ROW),
           m12 AS (PARTITION BY CODIGO_CLIENTE ORDER BY COMPETENCIA ROWS BETWEEN 11 PRECEDING AND CURRENT ROW)
)

SELECT
    j.*,
    j.LUCRO_BRUTO / NULLIF(j.RECEITA_LIQUIDA, 0) AS MARGEM_BRUTA,
    j.LUCRO_LIQUIDO / NULLIF(j.RECEITA_LIQUIDA, 0) AS MARGEM_LIQUIDA,
    j.LUCRO_BRUTO_3M / NULLIF(j.RECEITA_LIQUIDA_3M, 0) AS MARGEM_BRUTA_3M,
    j.LUCRO_LIQUIDO_3M / NULLIF(j.RECEITA_LIQUIDA_3M, 0) AS MARGEM_LIQUIDA_3M,
    j.LUCRO_LIQUIDO_12M / NULLIF(j.RECEITA_LIQUIDA_12M, 0) AS MARGEM_LIQUIDA_12M,
    j.ATIVO_CIRCULANTE / NULLIF(j.PASSIVO_CIRCULANTE, 0) AS LIQUIDEZ_CORRENTE,
    (j.ATIVO_CIRCULANTE - j.ESTOQUES) / NULLIF(j.PASSIVO_CIRCULANTE, 0) AS LIQUIDEZ_SECA,
    j.DISPONIVEL / NULLIF((j.RECEITA_BRUTA_3M - j.LUCRO_LIQUIDO_3M) / LEAST(j.N_MES, 3), 0) AS FOLEGO_CAIXA_MESES,
    j.CLIENTES_A_RECEBER / NULLIF(j.RECEITA_BRUTA_3M / LEAST(j.N_MES, 3), 0) * 30 AS PRAZO_RECEBIMENTO_DIAS,
    j.ESTOQUES / NULLIF(-j.CUSTOS, 0) * 30 AS DIAS_ESTOQUE,
    j.DIVIDA_BANCARIA / NULLIF(j.RECEITA_BRUTA_3M / LEAST(j.N_MES, 3), 0) AS DIVIDA_MESES_RECEITA,
    -(j.IMPOSTOS_RECEITA + j.IMPOSTO_LUCRO) / NULLIF(j.RECEITA_BRUTA, 0) AS CARGA_TRIBUTARIA,
    j.TOTAL_FOLHA / NULLIF(j.RECEITA_BRUTA, 0) AS PESO_FOLHA,
    j.RECEITA_BRUTA / NULLIF(j.FUNCIONARIOS, 0) AS RECEITA_POR_FUNCIONARIO,
    -- o mesmo ponto, 3, 6 e 12 meses antes
    LAG(j.RECEITA_BRUTA) OVER w AS RECEITA_BRUTA_MES_ANTERIOR,
    LAG(j.LUCRO_LIQUIDO) OVER w AS LUCRO_LIQUIDO_MES_ANTERIOR,
    LAG(j.RECEITA_BRUTA, 12) OVER w AS RECEITA_BRUTA_AA,
    LAG(j.LUCRO_LIQUIDO, 12) OVER w AS LUCRO_LIQUIDO_AA,
    LAG(j.RECEITA_BRUTA_3M, 12) OVER w AS RECEITA_BRUTA_3M_AA,
    LAG(j.LUCRO_BRUTO_3M / NULLIF(j.RECEITA_LIQUIDA_3M, 0), 12) OVER w AS MARGEM_BRUTA_3M_AA,
    LAG(j.DESPESAS_ADM_3M, 3) OVER w AS DESPESAS_ADM_3M_ANTERIOR,
    LAG(j.CLIENTES_A_RECEBER / NULLIF(j.RECEITA_BRUTA_3M / LEAST(j.N_MES, 3), 0) * 30, 6) OVER w AS PRAZO_RECEBIMENTO_DIAS_6M_ANTES,
    LAG(j.ESTOQUES / NULLIF(-j.CUSTOS, 0) * 30, 6) OVER w AS DIAS_ESTOQUE_6M_ANTES
FROM janelas j
WINDOW w AS (PARTITION BY j.CODIGO_CLIENTE ORDER BY j.COMPETENCIA);

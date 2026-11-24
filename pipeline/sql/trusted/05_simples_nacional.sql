-- =========================================================
-- TRUSTED - SIMPLES NACIONAL: FAIXA E DISTANCIA DO LIMITE
--
-- So para clientes do Simples. A faixa depende da receita bruta dos
-- ultimos 12 meses (RBT12) e do anexo da atividade.
--
--   Aliquota efetiva = (RBT12 x aliquota nominal - parcela a deduzir) / RBT12
--
-- Regras:
--   1. Empresa com menos de 12 meses: RBT12 proporcional
--      (media mensal x 12), marcada em RBT12_PROPORCIONAL.
--   2. FOLGA_FAIXA  = quanto ainda cabe de receita antes de subir de faixa.
--   3. RBT12_PROJETADO = RBT12 se os proximos 3 meses repetirem o
--      crescimento do ultimo trimestre sobre o mesmo trimestre do ano
--      anterior. E um aviso, nao uma previsao.
--   4. LIMITE_SIMPLES = teto da ultima faixa (config/simples_nacional.csv).
--
-- ATENCAO: tabelas simplificadas para a demonstracao. Conferir com a
-- legislacao vigente antes de qualquer uso real.
-- =========================================================

CREATE OR REPLACE TABLE simples_mensal AS

WITH base AS (
    SELECT
        i.CODIGO_CLIENTE, i.COMPETENCIA, c.ANEXO_SIMPLES AS ANEXO, i.RECEITA_BRUTA,
        i.MESES_NA_JANELA_12M < 12 AS RBT12_PROPORCIONAL,
        CASE WHEN i.MESES_NA_JANELA_12M < 12 THEN i.RECEITA_BRUTA_12M / i.MESES_NA_JANELA_12M * 12 ELSE i.RECEITA_BRUTA_12M END AS RBT12,
        i.RECEITA_BRUTA_3M / NULLIF(i.RECEITA_BRUTA_3M_AA, 0) AS CRESCIMENTO_TRIMESTRE
    FROM indicadores_mensal i
    JOIN dim_cliente c USING (CODIGO_CLIENTE)
    WHERE c.REGIME = 'Simples Nacional'
),

teto AS (SELECT ANEXO, MAX(RECEITA_12M_ATE) AS LIMITE_SIMPLES, MAX(FAIXA) AS ULTIMA_FAIXA FROM cfg_simples_nacional GROUP BY 1)

SELECT
    b.CODIGO_CLIENTE, b.COMPETENCIA, b.ANEXO, b.RBT12, b.RBT12_PROPORCIONAL,
    COALESCE(f.FAIXA, t.ULTIMA_FAIXA + 1) AS FAIXA,
    f.RECEITA_12M_ATE AS LIMITE_FAIXA,
    f.RECEITA_12M_ATE - b.RBT12 AS FOLGA_FAIXA,
    b.RBT12 / f.RECEITA_12M_ATE AS USO_FAIXA,
    (b.RBT12 * f.ALIQUOTA_NOMINAL - f.PARCELA_A_DEDUZIR) / NULLIF(b.RBT12, 0) AS ALIQUOTA_EFETIVA,
    t.LIMITE_SIMPLES,
    b.RBT12 / t.LIMITE_SIMPLES AS USO_LIMITE_SIMPLES,
    b.RBT12 > t.LIMITE_SIMPLES AS ACIMA_DO_LIMITE,
    b.RBT12 * (1 + (COALESCE(b.CRESCIMENTO_TRIMESTRE, 1) - 1) * 0.25) AS RBT12_PROJETADO
FROM base b
JOIN teto t USING (ANEXO)
LEFT JOIN cfg_simples_nacional f
       ON f.ANEXO = b.ANEXO
      AND f.FAIXA = (SELECT MIN(x.FAIXA) FROM cfg_simples_nacional x WHERE x.ANEXO = b.ANEXO AND x.RECEITA_12M_ATE >= b.RBT12);

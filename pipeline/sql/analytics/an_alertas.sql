-- =========================================================
-- ANALYTICS - ALERTAS EM LINGUAGEM DE DONO DE EMPRESA
--
-- Cada regra olha o ultimo balancete entregue de cada cliente e, se
-- disparar, escreve uma frase com o numero dentro. O contador nao
-- precisa montar o texto: recebe a pauta da conversa pronta.
--
-- Gravidade: 'Crítico' pede conversa agora; 'Atenção' pede acompanhar.
--
-- Comparacoes de resultado usam o mesmo trimestre do ano anterior,
-- para a sorveteria nao receber alerta todo inverno.
--
-- Os limites (liquidez 1,0, folha 40%, divida 3 meses...) sao pontos
-- de partida da demonstracao. Num escritorio real ficam em config,
-- por atividade.
-- =========================================================

CREATE OR REPLACE MACRO brl(x) AS 'R$ ' || REPLACE(FORMAT('{:,.0f}', ROUND(CAST(x AS DOUBLE))), ',', '.');
CREATE OR REPLACE MACRO pct1(x) AS REPLACE(FORMAT('{:.1f}', CAST(x AS DOUBLE) * 100), '.', ',') || '%';
CREATE OR REPLACE MACRO dec1(x) AS REPLACE(FORMAT('{:.1f}', CAST(x AS DOUBLE)), '.', ',');

CREATE OR REPLACE TABLE an_alertas AS

WITH u AS (
    SELECT i.*, c.REGIME, c.MESES_ATRASO, c.MES_REFERENCIA
    FROM indicadores_mensal i
    JOIN dim_cliente c ON c.CODIGO_CLIENTE = i.CODIGO_CLIENTE AND c.ULTIMO_BALANCETE = i.COMPETENCIA
),
s AS (
    SELECT s.* FROM simples_mensal s
    JOIN dim_cliente c ON c.CODIGO_CLIENTE = s.CODIGO_CLIENTE AND c.ULTIMO_BALANCETE = s.COMPETENCIA
)

SELECT CODIGO_CLIENTE, 'Crítico' AS GRAVIDADE, 'Resultado' AS TEMA, 10 AS ORDEM, 'Prejuízo no trimestre' AS TITULO,
       'A empresa perdeu ' || brl(-LUCRO_LIQUIDO_3M) || ' nos últimos três meses. A cada R$ 100 que entram, saem R$ '
       || dec1(100 - MARGEM_LIQUIDA_3M * 100) || '.' AS TEXTO
FROM u WHERE LUCRO_LIQUIDO_3M < 0

UNION ALL
SELECT CODIGO_CLIENTE, 'Crítico', 'Caixa', 20, 'Contas de curto prazo maiores que os recursos',
       'Para cada R$ 1,00 que vence nos próximos meses, a empresa tem R$ ' || REPLACE(FORMAT('{:.2f}', LIQUIDEZ_CORRENTE), '.', ',')
       || ' em caixa, contas a receber e estoque.'
FROM u WHERE LIQUIDEZ_CORRENTE < 1

UNION ALL
SELECT CODIGO_CLIENTE, 'Atenção', 'Caixa', 21, 'Pouca folga para pagar as contas',
       'Para cada R$ 1,00 que vence nos próximos meses, há R$ ' || REPLACE(FORMAT('{:.2f}', LIQUIDEZ_CORRENTE), '.', ',')
       || ' disponível. Qualquer atraso de cliente aperta o caixa.'
FROM u WHERE LIQUIDEZ_CORRENTE >= 1 AND LIQUIDEZ_CORRENTE < 1.2

UNION ALL
SELECT CODIGO_CLIENTE, 'Crítico', 'Dívida', 30, 'Dívida alta para o tamanho da empresa',
       'A dívida com bancos soma ' || brl(DIVIDA_BANCARIA) || ', o equivalente a ' || dec1(DIVIDA_MESES_RECEITA)
       || ' meses de faturamento.'
FROM u WHERE DIVIDA_MESES_RECEITA > 3

UNION ALL
SELECT CODIGO_CLIENTE, 'Atenção', 'Resultado', 40, 'Margem menor que há um ano',
       'De cada R$ 100 vendidos sobravam R$ ' || dec1(MARGEM_BRUTA_3M_AA * 100) || ' depois do custo; agora sobram R$ '
       || dec1(MARGEM_BRUTA_3M * 100) || '. O custo está subindo mais rápido que o preço.'
FROM u WHERE MARGEM_BRUTA_3M_AA - MARGEM_BRUTA_3M > 0.03

UNION ALL
SELECT CODIGO_CLIENTE, 'Atenção', 'Vendas', 50, 'Faturamento abaixo do ano passado',
       'O trimestre faturou ' || brl(RECEITA_BRUTA_3M) || ', ' || pct1(1 - RECEITA_BRUTA_3M / RECEITA_BRUTA_3M_AA)
       || ' a menos que o mesmo trimestre do ano anterior.'
FROM u WHERE RECEITA_BRUTA_3M < 0.9 * RECEITA_BRUTA_3M_AA

UNION ALL
SELECT CODIGO_CLIENTE, 'Atenção', 'Despesas', 60, 'Despesa administrativa subiu',
       'As despesas administrativas passaram de ' || brl(-DESPESAS_ADM_3M_ANTERIOR) || ' para ' || brl(-DESPESAS_ADM_3M)
       || ' no trimestre, alta de ' || pct1(DESPESAS_ADM_3M / DESPESAS_ADM_3M_ANTERIOR - 1) || '.'
FROM u WHERE DESPESAS_ADM_3M < 1.25 * DESPESAS_ADM_3M_ANTERIOR AND DESPESAS_ADM_3M_ANTERIOR < 0

UNION ALL
SELECT CODIGO_CLIENTE, 'Atenção', 'Recebimento', 70, 'Clientes demorando mais para pagar',
       'O prazo médio de recebimento passou de ' || CAST(ROUND(PRAZO_RECEBIMENTO_DIAS_6M_ANTES) AS INTEGER) || ' para '
       || CAST(ROUND(PRAZO_RECEBIMENTO_DIAS) AS INTEGER) || ' dias em seis meses. Há ' || brl(CLIENTES_A_RECEBER) || ' na rua.'
FROM u WHERE PRAZO_RECEBIMENTO_DIAS - PRAZO_RECEBIMENTO_DIAS_6M_ANTES > 10

UNION ALL
SELECT CODIGO_CLIENTE, 'Atenção', 'Recebimento', 71, 'Dinheiro preso em contas a receber',
       'A empresa leva em média ' || CAST(ROUND(PRAZO_RECEBIMENTO_DIAS) AS INTEGER) || ' dias para receber o que vende. São '
       || brl(CLIENTES_A_RECEBER) || ' a receber, ' || dec1(CLIENTES_A_RECEBER / NULLIF(RECEITA_BRUTA_3M / 3, 0)) || ' meses de faturamento.'
FROM u WHERE PRAZO_RECEBIMENTO_DIAS > 60 AND NOT COALESCE(PRAZO_RECEBIMENTO_DIAS - PRAZO_RECEBIMENTO_DIAS_6M_ANTES > 10, FALSE)

UNION ALL
SELECT CODIGO_CLIENTE, 'Atenção', 'Estoque', 80, 'Estoque parado crescendo',
       'O estoque cobre ' || CAST(ROUND(DIAS_ESTOQUE) AS INTEGER) || ' dias de venda; há seis meses cobria '
       || CAST(ROUND(DIAS_ESTOQUE_6M_ANTES) AS INTEGER) || '. São ' || brl(ESTOQUES) || ' em mercadoria.'
FROM u WHERE DIAS_ESTOQUE > 90 AND DIAS_ESTOQUE - DIAS_ESTOQUE_6M_ANTES > 15

UNION ALL
SELECT CODIGO_CLIENTE, 'Atenção', 'Pessoas', 90, 'Folha pesada',
       'A folha consome ' || pct1(PESO_FOLHA) || ' do faturamento do mês.'
FROM u WHERE PESO_FOLHA > 0.40

UNION ALL
SELECT s.CODIGO_CLIENTE, 'Crítico', 'Impostos', 5, 'Acima do limite do Simples Nacional',
       'O faturamento de 12 meses chegou a ' || brl(s.RBT12) || ', acima do limite de ' || brl(s.LIMITE_SIMPLES)
       || '. É preciso planejar a mudança de regime.'
FROM s WHERE s.ACIMA_DO_LIMITE

UNION ALL
SELECT s.CODIGO_CLIENTE, 'Crítico', 'Impostos', 6, 'Perto do limite do Simples Nacional',
       'O faturamento de 12 meses está em ' || brl(s.RBT12) || ', ' || pct1(s.USO_LIMITE_SIMPLES) || ' do limite de '
       || brl(s.LIMITE_SIMPLES) || '. Faltam ' || brl(s.LIMITE_SIMPLES - s.RBT12) || ' para sair do Simples.'
FROM s WHERE NOT s.ACIMA_DO_LIMITE AND s.USO_LIMITE_SIMPLES >= 0.9

UNION ALL
SELECT s.CODIGO_CLIENTE, 'Atenção', 'Impostos', 7, 'Perto de subir de faixa no Simples',
       'Faltam ' || brl(s.FOLGA_FAIXA) || ' de faturamento em 12 meses para passar da faixa ' || s.FAIXA || ' para a faixa '
       || (s.FAIXA + 1) || ', com alíquota maior.'
FROM s WHERE NOT s.ACIMA_DO_LIMITE AND s.USO_LIMITE_SIMPLES < 0.9 AND s.USO_FAIXA >= 0.95

UNION ALL
SELECT CODIGO_CLIENTE, CASE WHEN MESES_ATRASO >= 2 THEN 'Crítico' ELSE 'Atenção' END, 'Documentos', 1, 'Balancete atrasado',
       'O último balancete entregue é de ' || STRFTIME(ULTIMO_BALANCETE, '%m/%Y') || ': ' || MESES_ATRASO
       || CASE WHEN MESES_ATRASO = 1 THEN ' mês' ELSE ' meses' END || ' de atraso. Os números deste painel param nessa data.'
FROM dim_cliente WHERE MESES_ATRASO > 0;


-- ---------------------------------------------------------
-- Visao da carteira: uma linha por cliente, para o contador
-- ---------------------------------------------------------
CREATE OR REPLACE TABLE an_carteira AS
SELECT
    c.CODIGO_CLIENTE, c.CLIENTE, c.ATIVIDADE, c.CIDADE, c.REGIME, c.ANEXO_SIMPLES, c.RESPONSAVEL, c.HONORARIO,
    c.ULTIMO_BALANCETE, c.MESES_ATRASO, c.MESES_ENTREGUES,
    i.RECEITA_BRUTA, i.RECEITA_BRUTA_12M, i.LUCRO_LIQUIDO, i.LUCRO_LIQUIDO_3M, i.MARGEM_LIQUIDA_3M, i.MARGEM_LIQUIDA_12M,
    i.LIQUIDEZ_CORRENTE, i.DIVIDA_MESES_RECEITA, i.CARGA_TRIBUTARIA, i.PESO_FOLHA,
    s.FAIXA AS FAIXA_SIMPLES, s.USO_FAIXA, s.USO_LIMITE_SIMPLES, s.RBT12,
    COALESCE(a.CRITICOS, 0) AS ALERTAS_CRITICOS,
    COALESCE(a.ATENCOES, 0) AS ALERTAS_ATENCAO,
    CASE WHEN COALESCE(a.CRITICOS, 0) > 0 THEN 'Crítico' WHEN COALESCE(a.ATENCOES, 0) > 0 THEN 'Atenção' ELSE 'Em dia' END AS SITUACAO
FROM dim_cliente c
JOIN indicadores_mensal i ON i.CODIGO_CLIENTE = c.CODIGO_CLIENTE AND i.COMPETENCIA = c.ULTIMO_BALANCETE
LEFT JOIN simples_mensal s ON s.CODIGO_CLIENTE = c.CODIGO_CLIENTE AND s.COMPETENCIA = c.ULTIMO_BALANCETE
LEFT JOIN (SELECT CODIGO_CLIENTE, COUNT(*) FILTER (WHERE GRAVIDADE = 'Crítico') AS CRITICOS,
                  COUNT(*) FILTER (WHERE GRAVIDADE = 'Atenção') AS ATENCOES
           FROM an_alertas GROUP BY 1) a ON a.CODIGO_CLIENTE = c.CODIGO_CLIENTE;

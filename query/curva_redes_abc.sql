/* =====================================================================
   CURVA REDES ABC  -  vendas por REDE x PRODUTO x MES (12 meses fechados)
   M. Ferretti  |  SQL Server / Protheus  |  Banco CO136Y_160463_PR_PD
   ---------------------------------------------------------------------
   Rede    = grupo de vendas  (SA1010.A1_GRPVEN -> ACY010.ACY_DESCRI)
   Metrica = valor faturado   (SC6010.C6_VALOR)
   ---------------------------------------------------------------------
   FILIAIS neste banco:
     - Cadastros SA1010 / ACY010 -> FILIAL = ''  (vazio)
     - Produto  SB1010 / SBM010  -> FILIAL = '01'
     - Movimento SC5010 / SC6010 -> FILIAL = '01'
   Janela: 12 MESES FECHADOS + o MES ATUAL (parcial).
     Ex.: rodando em 24/jul/2026 => jul/2025 ate 24/jul/2026 (13 meses).
     Os 12 fechados (jul/2025..jun/2026) sao a base de CALCULO no site;
     o mes atual (jul/2026) vem so para VISUALIZACAO (nao entra nas
     medias, totais nem tendencias - o site separa via mesParcial).
   Exclui o grupo guarda-chuva "SEM GRUPO OU REDE" (nao e rede real).
   ===================================================================== */

DECLARE @MesAtual DATE  = DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1);
DECLARE @IniDate  DATE  = DATEADD(MONTH, -12, @MesAtual);   -- 1o dia, 12 meses atras
DECLARE @FimDate  DATE  = CAST(GETDATE() AS DATE);          -- HOJE (inclui o mes atual parcial)
DECLARE @DataIni  CHAR(8) = CONVERT(CHAR(8), @IniDate, 112);
DECLARE @DataFim  CHAR(8) = CONVERT(CHAR(8), @FimDate, 112);

SELECT
    RTRIM(ACY.ACY_DESCRI)                                  AS rede,
    A1.A1_GRPVEN                                           AS redeCod,
    SUBSTRING(C5.C5_EMISSAO,1,4) + '-' +
        SUBSTRING(C5.C5_EMISSAO,5,2)                       AS anoMes,
    RTRIM(C6.C6_PRODUTO)                                   AS prodCod,
    RTRIM(B1.B1_DESC)                                      AS prodDesc,
    RTRIM(B1.B1_ZZMARCA)                                   AS marca,
    RTRIM(ISNULL(BM.BM_DESC,''))                           AS grupoProduto,
    SUM(C6.C6_VALOR)                                       AS valor,
    SUM(C6.C6_QTDVEN)                                      AS qtd
FROM SC6010 C6
INNER JOIN SC5010 C5
        ON C5.C5_FILIAL  = C6.C6_FILIAL
       AND C5.C5_NUM     = C6.C6_NUM
       AND C5.D_E_L_E_T_ = ' '
INNER JOIN SA1010 A1
        ON A1.A1_FILIAL  = ''
       AND A1.A1_COD     = C6.C6_CLI
       AND A1.A1_LOJA    = C6.C6_LOJA
       AND A1.D_E_L_E_T_ = ' '
INNER JOIN ACY010 ACY
        ON ACY.ACY_FILIAL = ''
       AND ACY.ACY_GRPVEN = A1.A1_GRPVEN
       AND ACY.D_E_L_E_T_ = ' '
INNER JOIN SB1010 B1
        ON B1.B1_FILIAL  = '01'
       AND B1.B1_COD     = C6.C6_PRODUTO
       AND B1.D_E_L_E_T_ = ' '
LEFT JOIN SBM010 BM
        ON BM.BM_FILIAL  = '01'
       AND BM.BM_GRUPO   = B1.B1_GRUPO
       AND BM.D_E_L_E_T_ = ' '
WHERE C6.C6_FILIAL   = '01'
  AND C6.D_E_L_E_T_  = ' '
  AND A1.A1_GRPVEN  <> ''
  AND ACY.ACY_DESCRI NOT LIKE '%SEM GRUPO%'
  AND ACY.ACY_DESCRI NOT LIKE '%SEM REDE%'
  AND C5.C5_EMISSAO >= @DataIni
  AND C5.C5_EMISSAO <= @DataFim
GROUP BY
    RTRIM(ACY.ACY_DESCRI),
    A1.A1_GRPVEN,
    SUBSTRING(C5.C5_EMISSAO,1,4) + '-' + SUBSTRING(C5.C5_EMISSAO,5,2),
    RTRIM(C6.C6_PRODUTO),
    RTRIM(B1.B1_DESC),
    RTRIM(B1.B1_ZZMARCA),
    RTRIM(ISNULL(BM.BM_DESC,''))
HAVING SUM(C6.C6_VALOR) <> 0
ORDER BY rede, anoMes, valor DESC;

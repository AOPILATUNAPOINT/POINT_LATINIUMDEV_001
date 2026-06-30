IF OBJECT_ID('[dbo].[Com_ConsultaPresupuestoEmpresaContado]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ConsultaPresupuestoEmpresaContado]
GO

CREATE PROCEDURE [dbo].[Com_ConsultaPresupuestoEmpresaContado]
    @Periodo Int
AS
BEGIN
    SET NOCOUNT ON;

    Declare @sSQL     Varchar(MAX) = ''
    Declare @Columnas Varchar(Max) = ''

    Create Table #Meses (Mes Varchar(100))
    Exec ('Insert Into #Meses (Mes) Select Descripcion From PeriodoMes Order By idPeriodoMes')

    Select @Columnas = COALESCE(@Columnas + '[' + Mes + '],', '')
    FROM (Select Distinct Mes From #Meses) as a

    Set @Columnas = LEFT(@Columnas, LEN(@Columnas) - 1)

    Set @sSQL =
        'SELECT * FROM
        (Select Nombre, Descripcion, e.Valor, CONVERT(DECIMAL(18, 2), 0) As TOTAL
         From Com_PresupuestoContado e
         Inner Join Bodega b On b.bodega = e.Bodega
         Inner Join PeriodoMes p On p.idPeriodoMes = e.idPeriodoMes
         Where e.Periodo = ' + CAST(@Periodo As Varchar(4)) + ') AS SourceTable
        PIVOT
        (SUM(Valor)
         FOR Descripcion IN (' + @Columnas + ')) AS PivotTable;'

    Exec (@sSQL)
END
GO

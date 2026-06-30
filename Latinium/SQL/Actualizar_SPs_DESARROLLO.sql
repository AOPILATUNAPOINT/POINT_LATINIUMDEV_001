-- =============================================================================
-- SCRIPT: Actualizar SPs de Presupuesto Empresa en base DESARROLLO
-- Generado: 2026-06-30
-- Fuente:   POINT (produccion)
-- Aplica a: DESARROLLO
-- =============================================================================
-- INSTRUCCIONES:
--   1. Conectarse a la base DESARROLLO en SSMS
--   2. Ejecutar este script completo (F5)
--   3. Los SPs marcados como PENDIENTE deben copiarse manualmente desde POINT
-- =============================================================================

PRINT '===== INICIO ACTUALIZACION SPs PRESUPUESTO EMPRESA ====='
GO

-- ---------------------------------------------------------------------------
-- 1. Com_ConsultaPresupuestoEmpresa
-- ---------------------------------------------------------------------------
PRINT '>>> Com_ConsultaPresupuestoEmpresa'
GO
IF OBJECT_ID('[dbo].[Com_ConsultaPresupuestoEmpresa]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ConsultaPresupuestoEmpresa]
GO

CREATE PROCEDURE [dbo].[Com_ConsultaPresupuestoEmpresa]
    @Periodo Int
AS
BEGIN
    SET NOCOUNT ON;

    Declare @sSQL    Varchar(MAX) = ''
    Declare @Columnas Varchar(Max) = ''

    Create Table #Meses (Mes Varchar(100))
    Exec ('Insert Into #Meses (Mes) Select Descripcion From PeriodoMes Order By idPeriodoMes')

    Select @Columnas = COALESCE(@Columnas + '[' + Mes + '],', '')
    FROM (Select Distinct Mes From #Meses) as a

    Set @Columnas = LEFT(@Columnas, LEN(@Columnas) - 1)

    Set @sSQL =
        'SELECT * FROM
        (Select Nombre, Descripcion, e.Valor, CONVERT(DECIMAL(18, 2), 0) As TOTAL
         From Com_Empresa e
         Inner Join Bodega b On b.bodega = e.Bodega
         Inner Join PeriodoMes p On p.idPeriodoMes = e.idPeriodoMes
         Where e.Periodo = ' + CAST(@Periodo As Varchar(4)) + ') AS SourceTable
        PIVOT
        (SUM(Valor)
         FOR Descripcion IN (' + @Columnas + ')) AS PivotTable;'

    Exec (@sSQL)
END
GO

-- ---------------------------------------------------------------------------
-- 2. Com_ConsultaPresupuestoEmpresaCredito
-- ---------------------------------------------------------------------------
PRINT '>>> Com_ConsultaPresupuestoEmpresaCredito'
GO
IF OBJECT_ID('[dbo].[Com_ConsultaPresupuestoEmpresaCredito]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ConsultaPresupuestoEmpresaCredito]
GO

CREATE PROCEDURE [dbo].[Com_ConsultaPresupuestoEmpresaCredito]
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
         From Com_PresupuestoCredito e
         Inner Join Bodega b On b.bodega = e.Bodega
         Inner Join PeriodoMes p On p.idPeriodoMes = e.idPeriodoMes
         Where e.Periodo = ' + CAST(@Periodo As Varchar(4)) + ') AS SourceTable
        PIVOT
        (SUM(Valor)
         FOR Descripcion IN (' + @Columnas + ')) AS PivotTable;'

    Exec (@sSQL)
END
GO

-- ---------------------------------------------------------------------------
-- 3. Com_ConsultaPresupuestoEmpresaGE
-- ---------------------------------------------------------------------------
PRINT '>>> Com_ConsultaPresupuestoEmpresaGE'
GO
IF OBJECT_ID('[dbo].[Com_ConsultaPresupuestoEmpresaGE]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ConsultaPresupuestoEmpresaGE]
GO

CREATE PROCEDURE [dbo].[Com_ConsultaPresupuestoEmpresaGE]
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
         From Com_PresupuestoEmpresaGE e
         Inner Join Bodega b On b.bodega = e.Bodega
         Inner Join PeriodoMes p On p.idPeriodoMes = e.idPeriodoMes
         Where e.Periodo = ' + CAST(@Periodo As Varchar(4)) + ') AS SourceTable
        PIVOT
        (SUM(Valor)
         FOR Descripcion IN (' + @Columnas + ')) AS PivotTable;'

    Exec (@sSQL)
END
GO

-- ---------------------------------------------------------------------------
-- 4. Com_ConsultaPresupuestoEmpresaContado
-- ---------------------------------------------------------------------------
PRINT '>>> Com_ConsultaPresupuestoEmpresaContado'
GO
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

-- ---------------------------------------------------------------------------
-- 5. Com_ActualizaValoresPresupuestoEmpresa
-- ---------------------------------------------------------------------------
PRINT '>>> Com_ActualizaValoresPresupuestoEmpresa'
GO
IF OBJECT_ID('[dbo].[Com_ActualizaValoresPresupuestoEmpresa]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ActualizaValoresPresupuestoEmpresa]
GO

CREATE PROCEDURE [dbo].[Com_ActualizaValoresPresupuestoEmpresa]
    @Periodo       Int,
    @Local         Varchar(200),
    @Mes           Varchar(50),
    @Valor         Decimal(18, 2)
AS
BEGIN
    SET NOCOUNT ON;

    Declare @Bodega Int, @idPeriodoMes Int

    Select @Bodega       = Bodega       From Bodega    Where Nombre       = @Local
    Select @idPeriodoMes = idPeriodoMes From PeriodoMes Where Descripcion = @Mes

    Update Com_Empresa
    Set    Valor = @Valor
    Where  Periodo      = @Periodo
      And  Bodega       = @Bodega
      And  idPeriodoMes = @idPeriodoMes

    /* ACTUALIZA ASIGNACION VENDEDORES */

    Declare @Ceros         Varchar(2),
            @PeriodoAV     Varchar(7),
            @DiasMes       Int,
            @PeriodoActual Varchar(8)

    Set @Ceros         = SUBSTRING('00', 1, 2 - LEN(@idPeriodoMes))
    Set @PeriodoAV     = @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '/' + CAST(@Periodo As Varchar(4))
    Set @PeriodoActual = CONVERT(varchar(4), @Periodo) + @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '01'
    Set @DiasMes       = DATEPART(DAY, DATEADD(s, -1, DATEADD(mm, DATEDIFF(m, 0, CONVERT(Datetime, @PeriodoActual, 112)) + 1, 0)))

    If Exists(Select idCom_AsignacionDeVendedores From Com_AsignacionDeVendedores Where Bodega = @Bodega And Periodo = @PeriodoAV)
    Begin
        Update Com_AsignacionDeVendedores
        Set    MetaMensual = dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Periodo, @idPeriodoMes, @Bodega, idCargo, 1) / @DiasMes * NumeroDias
        Where  Bodega       = @Bodega
          And  Periodo      = @PeriodoAV
          And  idCom_Estado In (1, 3)
    End

    If Exists(Select idCom_AsignacionJefes From Com_AsignacionJefes Where Bodega = @Bodega And Periodo = @PeriodoAV)
        Update Com_AsignacionJefes
        Set    MetaMensual = (@Valor / @DiasMes) * Dias
        Where  Bodega  = @Bodega
          And  Periodo = @PeriodoAV

END
GO

-- ---------------------------------------------------------------------------
-- 6. Com_ActualizaValoresPresupuestoEmpresaCredito
-- ---------------------------------------------------------------------------
PRINT '>>> Com_ActualizaValoresPresupuestoEmpresaCredito'
GO
IF OBJECT_ID('[dbo].[Com_ActualizaValoresPresupuestoEmpresaCredito]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ActualizaValoresPresupuestoEmpresaCredito]
GO

CREATE PROCEDURE [dbo].[Com_ActualizaValoresPresupuestoEmpresaCredito]
    @Periodo       Int,
    @Local         Varchar(200),
    @Mes           Varchar(50),
    @Valor         Decimal(18, 2)
AS
BEGIN
    SET NOCOUNT ON;

    Declare @Bodega Int, @idPeriodoMes Int

    Select @Bodega       = Bodega       From Bodega    Where Nombre       = @Local
    Select @idPeriodoMes = idPeriodoMes From PeriodoMes Where Descripcion = @Mes

    Update Com_PresupuestoCredito
    Set    Valor = @Valor
    Where  Periodo      = @Periodo
      And  Bodega       = @Bodega
      And  idPeriodoMes = @idPeriodoMes

    /* ACTUALIZA ASIGNACION VENDEDORES */

    Declare @Ceros         Varchar(2),
            @PeriodoAV     Varchar(7),
            @DiasMes       Int,
            @PeriodoActual Varchar(8)

    Set @Ceros         = SUBSTRING('00', 1, 2 - LEN(@idPeriodoMes))
    Set @PeriodoAV     = @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '/' + CAST(@Periodo As Varchar(4))
    Set @PeriodoActual = CONVERT(varchar(4), @Periodo) + @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '01'
    Set @DiasMes       = DATEPART(DAY, DATEADD(s, -1, DATEADD(mm, DATEDIFF(m, 0, CONVERT(Datetime, @PeriodoActual, 112)) + 1, 0)))

    If Exists(Select idCom_AsignacionDeVendedores From Com_AsignacionDeVendedores Where Bodega = @Bodega And Periodo = @PeriodoAV)
    Begin
        Update Com_AsignacionDeVendedores
        Set    MetaMensualCredito = dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Periodo, @idPeriodoMes, @Bodega, idCargo, 3) / @DiasMes * NumeroDias
        Where  Bodega       = @Bodega
          And  Periodo      = @PeriodoAV
          And  idCom_Estado In (1, 3)
    End

    If Exists(Select idCom_AsignacionJefes From Com_AsignacionJefes Where Bodega = @Bodega And Periodo = @PeriodoAV)
        Update Com_AsignacionJefes
        Set    MetaMensualCredito = (@Valor / @DiasMes) * Dias
        Where  Bodega  = @Bodega
          And  Periodo = @PeriodoAV

END
GO

-- ---------------------------------------------------------------------------
-- 7. Com_ActualizaValoresPresupuestoAccesorios
-- ---------------------------------------------------------------------------
PRINT '>>> Com_ActualizaValoresPresupuestoAccesorios'
GO
IF OBJECT_ID('[dbo].[Com_ActualizaValoresPresupuestoAccesorios]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ActualizaValoresPresupuestoAccesorios]
GO

CREATE PROCEDURE [dbo].[Com_ActualizaValoresPresupuestoAccesorios]
    @Periodo       Int,
    @Local         Varchar(200),
    @Mes           Varchar(50),
    @Valor         Decimal(18, 2)
AS
BEGIN
    SET NOCOUNT ON;

    Declare @Bodega Int, @idPeriodoMes Int

    Select @Bodega       = Bodega       From Bodega    Where Nombre       = @Local
    Select @idPeriodoMes = idPeriodoMes From PeriodoMes Where Descripcion = @Mes

    Update Com_Accesorios
    Set    Valor = @Valor
    Where  Periodo      = @Periodo
      And  Bodega       = @Bodega
      And  idPeriodoMes = @idPeriodoMes

    /* ACTUALIZA ASIGNACION VENDEDORES */

    Declare @Ceros         Varchar(2),
            @PeriodoAV     Varchar(7),
            @DiasMes       Int,
            @PeriodoActual Varchar(8)

    Set @Ceros         = SUBSTRING('00', 1, 2 - LEN(@idPeriodoMes))
    Set @PeriodoAV     = @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '/' + CAST(@Periodo As Varchar(4))
    Set @PeriodoActual = CONVERT(varchar(4), @Periodo) + @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '01'
    Set @DiasMes       = DATEPART(DAY, DATEADD(s, -1, DATEADD(mm, DATEDIFF(m, 0, CONVERT(Datetime, @PeriodoActual, 112)) + 1, 0)))

    If Exists(Select idCom_AsignacionDeVendedores From Com_AsignacionDeVendedores Where Bodega = @Bodega And Periodo = @PeriodoAV)
    Begin
        Update Com_AsignacionDeVendedores
        Set    MetaMensualAccesorios = dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Periodo, @idPeriodoMes, @Bodega, idCargo, 6) / @DiasMes * NumeroDias
        Where  Bodega       = @Bodega
          And  Periodo      = @PeriodoAV
          And  idCom_Estado In (1, 3)
    End

    If Exists(Select idCom_AsignacionJefes From Com_AsignacionJefes Where Bodega = @Bodega And Periodo = @PeriodoAV)
        Update Com_AsignacionJefes
        Set    MetaMensualAccesorios = (@Valor / @DiasMes) * Dias
        Where  Bodega  = @Bodega
          And  Periodo = @PeriodoAV

END
GO

-- ---------------------------------------------------------------------------
-- 8. Com_ActualizaValoresPresupuestoMotos
-- ---------------------------------------------------------------------------
PRINT '>>> Com_ActualizaValoresPresupuestoMotos'
GO
IF OBJECT_ID('[dbo].[Com_ActualizaValoresPresupuestoMotos]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ActualizaValoresPresupuestoMotos]
GO

CREATE PROCEDURE [dbo].[Com_ActualizaValoresPresupuestoMotos]
    @Periodo       Int,
    @Local         Varchar(200),
    @Mes           Varchar(50),
    @Valor         Decimal(18, 2)
AS
BEGIN
    SET NOCOUNT ON;

    Declare @Bodega Int, @idPeriodoMes Int

    Select @Bodega       = Bodega       From Bodega    Where Nombre       = @Local
    Select @idPeriodoMes = idPeriodoMes From PeriodoMes Where Descripcion = @Mes

    Update Com_PresupuestoMotos
    Set    Valor = @Valor
    Where  Periodo      = @Periodo
      And  Bodega       = @Bodega
      And  idPeriodoMes = @idPeriodoMes

    /* ACTUALIZA ASIGNACION VENDEDORES */

    Declare @Ceros         Varchar(2),
            @PeriodoAV     Varchar(7),
            @DiasMes       Int,
            @PeriodoActual Varchar(8)

    Set @Ceros         = SUBSTRING('00', 1, 2 - LEN(@idPeriodoMes))
    Set @PeriodoAV     = @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '/' + CAST(@Periodo As Varchar(4))
    Set @PeriodoActual = CONVERT(varchar(4), @Periodo) + @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '01'
    Set @DiasMes       = DATEPART(DAY, DATEADD(s, -1, DATEADD(mm, DATEDIFF(m, 0, CONVERT(Datetime, @PeriodoActual, 112)) + 1, 0)))

    If Exists(Select idCom_AsignacionDeVendedores From Com_AsignacionDeVendedores Where Bodega = @Bodega And Periodo = @PeriodoAV)
    Begin
        Update Com_AsignacionDeVendedores
        Set    MetaMensualMotos = dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Periodo, @idPeriodoMes, @Bodega, idCargo, 8) / @DiasMes * NumeroDias
        Where  Bodega       = @Bodega
          And  Periodo      = @PeriodoAV
          And  idCom_Estado In (1, 3)
    End

    If Exists(Select idCom_AsignacionJefes From Com_AsignacionJefes Where Bodega = @Bodega And Periodo = @PeriodoAV)
        Update Com_AsignacionJefes
        Set    MetaMensualMotos = (@Valor / @DiasMes) * Dias
        Where  Bodega  = @Bodega
          And  Periodo = @PeriodoAV

END
GO

-- ---------------------------------------------------------------------------
-- 9. Com_ActualizaValoresPresupuestoEmpresaGE
-- ---------------------------------------------------------------------------
PRINT '>>> Com_ActualizaValoresPresupuestoEmpresaGE'
GO
IF OBJECT_ID('[dbo].[Com_ActualizaValoresPresupuestoEmpresaGE]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ActualizaValoresPresupuestoEmpresaGE]
GO

CREATE PROCEDURE [dbo].[Com_ActualizaValoresPresupuestoEmpresaGE]
    @Periodo       Int,
    @Local         Varchar(200),
    @Mes           Varchar(50),
    @Valor         Decimal(18, 2)
AS
BEGIN
    SET NOCOUNT ON;

    Declare @Bodega Int, @idPeriodoMes Int

    Select @Bodega       = Bodega       From Bodega    Where Nombre       = @Local
    Select @idPeriodoMes = idPeriodoMes From PeriodoMes Where Descripcion = @Mes

    Update Com_PresupuestoEmpresaGE
    Set    Valor = @Valor
    Where  Periodo      = @Periodo
      And  Bodega       = @Bodega
      And  idPeriodoMes = @idPeriodoMes

    /* ACTUALIZA ASIGNACION VENDEDORES */

    Declare @Ceros         Varchar(2),
            @PeriodoAV     Varchar(7),
            @DiasMes       Int,
            @PeriodoActual Varchar(8)

    Set @Ceros         = SUBSTRING('00', 1, 2 - LEN(@idPeriodoMes))
    Set @PeriodoAV     = @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '/' + CAST(@Periodo As Varchar(4))
    Set @PeriodoActual = CONVERT(varchar(4), @Periodo) + @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '01'
    Set @DiasMes       = DATEPART(DAY, DATEADD(s, -1, DATEADD(mm, DATEDIFF(m, 0, CONVERT(Datetime, @PeriodoActual, 112)) + 1, 0)))

    If Exists(Select idCom_AsignacionDeVendedores From Com_AsignacionDeVendedores Where Bodega = @Bodega And Periodo = @PeriodoAV)
    Begin
        Update Com_AsignacionDeVendedores
        Set    MetaMensualGEx = dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Periodo, @idPeriodoMes, @Bodega, idCargo, 4) / @DiasMes * NumeroDias
        Where  Bodega       = @Bodega
          And  Periodo      = @PeriodoAV
          And  idCom_Estado In (1, 3)
    End

    If Exists(Select idCom_AsignacionJefes From Com_AsignacionJefes Where Bodega = @Bodega And Periodo = @PeriodoAV)
        Update Com_AsignacionJefes
        Set    MetaMensualGex = (@Valor / @DiasMes) * Dias
        Where  Bodega  = @Bodega
          And  Periodo = @PeriodoAV

END
GO

-- ---------------------------------------------------------------------------
-- 10. Com_ActualizaValoresPresupuestoEmpresaContado
-- ---------------------------------------------------------------------------
PRINT '>>> Com_ActualizaValoresPresupuestoEmpresaContado'
GO
IF OBJECT_ID('[dbo].[Com_ActualizaValoresPresupuestoEmpresaContado]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ActualizaValoresPresupuestoEmpresaContado]
GO

CREATE PROCEDURE [dbo].[Com_ActualizaValoresPresupuestoEmpresaContado]
    @Periodo       Int,
    @Local         Varchar(200),
    @Mes           Varchar(50),
    @Valor         Decimal(18, 2)
AS
BEGIN
    SET NOCOUNT ON;

    Declare @Bodega Int, @idPeriodoMes Int

    Select @Bodega       = Bodega       From Bodega    Where Nombre       = @Local
    Select @idPeriodoMes = idPeriodoMes From PeriodoMes Where Descripcion = @Mes

    Update Com_PresupuestoContado
    Set    Valor = @Valor
    Where  Periodo      = @Periodo
      And  Bodega       = @Bodega
      And  idPeriodoMes = @idPeriodoMes

    /* ACTUALIZA ASIGNACION VENDEDORES */

    Declare @Ceros         Varchar(2),
            @PeriodoAV     Varchar(7),
            @DiasMes       Int,
            @PeriodoActual Varchar(8)

    Set @Ceros         = SUBSTRING('00', 1, 2 - LEN(@idPeriodoMes))
    Set @PeriodoAV     = @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '/' + CAST(@Periodo As Varchar(4))
    Set @PeriodoActual = CONVERT(varchar(4), @Periodo) + @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '01'
    Set @DiasMes       = DATEPART(DAY, DATEADD(s, -1, DATEADD(mm, DATEDIFF(m, 0, CONVERT(Datetime, @PeriodoActual, 112)) + 1, 0)))

    If Exists(Select idCom_AsignacionDeVendedores From Com_AsignacionDeVendedores Where Bodega = @Bodega And Periodo = @PeriodoAV)
    Begin
        Update Com_AsignacionDeVendedores
        Set    MetaMensualContado = dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Periodo, @idPeriodoMes, @Bodega, idCargo, 2) / @DiasMes * NumeroDias
        Where  Bodega       = @Bodega
          And  Periodo      = @PeriodoAV
          And  idCom_Estado In (1, 3)
    End

    If Exists(Select idCom_AsignacionJefes From Com_AsignacionJefes Where Bodega = @Bodega And Periodo = @PeriodoAV)
        Update Com_AsignacionJefes
        Set    MetaMensualContado = (@Valor / @DiasMes) * Dias
        Where  Bodega  = @Bodega
          And  Periodo = @PeriodoAV

END
GO

-- ---------------------------------------------------------------------------
-- 11. Com_ActualizaValoresPresupuestoEmpresaCarpa
-- ---------------------------------------------------------------------------
PRINT '>>> Com_ActualizaValoresPresupuestoEmpresaCarpa'
GO
IF OBJECT_ID('[dbo].[Com_ActualizaValoresPresupuestoEmpresaCarpa]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ActualizaValoresPresupuestoEmpresaCarpa]
GO

CREATE PROCEDURE [dbo].[Com_ActualizaValoresPresupuestoEmpresaCarpa]
    @Periodo       Int,
    @Local         Varchar(200),
    @Mes           Varchar(50),
    @Valor         Decimal(18, 2)
AS
BEGIN
    SET NOCOUNT ON;

    Declare @Bodega Int, @idPeriodoMes Int

    Select @Bodega       = Bodega       From Bodega    Where Nombre       = @Local
    Select @idPeriodoMes = idPeriodoMes From PeriodoMes Where Descripcion = @Mes

    Update Com_PresupuestoCarpa
    Set    Valor = @Valor
    Where  Periodo      = @Periodo
      And  Bodega       = @Bodega
      And  idPeriodoMes = @idPeriodoMes

    /* ACTUALIZA ASIGNACION VENDEDORES */

    Declare @Ceros         Varchar(2),
            @PeriodoAV     Varchar(7),
            @DiasMes       Int,
            @PeriodoActual Varchar(8)

    Set @Ceros         = SUBSTRING('00', 1, 2 - LEN(@idPeriodoMes))
    Set @PeriodoAV     = @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '/' + CAST(@Periodo As Varchar(4))
    Set @PeriodoActual = CONVERT(varchar(4), @Periodo) + @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '01'
    Set @DiasMes       = DATEPART(DAY, DATEADD(s, -1, DATEADD(mm, DATEDIFF(m, 0, CONVERT(Datetime, @PeriodoActual, 112)) + 1, 0)))

    If Exists(Select idCom_AsignacionDeVendedores From Com_AsignacionDeVendedores Where Bodega = @Bodega And Periodo = @PeriodoAV)
    Begin
        Update Com_AsignacionDeVendedores
        Set    MensualCarpa = dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Periodo, @idPeriodoMes, @Bodega, idCargo, 5) / @DiasMes * NumeroDias
        Where  Bodega       = @Bodega
          And  Periodo      = @PeriodoAV
          And  idCom_Estado In (1, 3)
    End

    If Exists(Select idCom_AsignacionJefes From Com_AsignacionJefes Where Bodega = @Bodega And Periodo = @PeriodoAV)
        Update Com_AsignacionJefes
        Set    MetaMensualCarpa = (@Valor / @DiasMes) * Dias
        Where  Bodega  = @Bodega
          And  Periodo = @PeriodoAV

END
GO

-- ---------------------------------------------------------------------------
-- 12. Com_GeneraPresupuestoEmpresa
-- ---------------------------------------------------------------------------
PRINT '>>> Com_GeneraPresupuestoEmpresa'
GO
IF OBJECT_ID('[dbo].[Com_GeneraPresupuestoEmpresa]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_GeneraPresupuestoEmpresa]
GO

-- Exec Com_GeneraPresupuestoEmpresa 2026
CREATE PROCEDURE [dbo].[Com_GeneraPresupuestoEmpresa]
    @Periodo Int
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @Bodega Int

    /* GENERAL */
    Declare cCom_Empresa Cursor For
        Select Bodega From Bodega
        Where  Activo = 1 And Factura = 1 And Credito = 1
          And  CONVERT(Date, ISNULL(FechaCierre, '20301231')) > '20241231'
        Order By Nombre

    OPEN cCom_Empresa
    Fetch Next From cCom_Empresa Into @Bodega
    While (@@FETCH_STATUS = 0)
    Begin
        Insert Into Com_Empresa (Periodo, Bodega, idPeriodoMes, Valor)
        Select @Periodo, @Bodega, idPeriodoMes, 0 From PeriodoMes Order By idPeriodoMes
        Fetch Next From cCom_Empresa Into @Bodega
    End
    Close cCom_Empresa
    DeAllocate cCom_Empresa

    /* CREDITO */
    Declare cCom_EmpresaCredito Cursor For
        Select Bodega From Bodega
        Where  Activo = 1 And Factura = 1 And Credito = 1
          And  CONVERT(Date, ISNULL(FechaCierre, '20301231')) > '20241231'
        Order By Nombre

    OPEN cCom_EmpresaCredito
    Fetch Next From cCom_EmpresaCredito Into @Bodega
    While (@@FETCH_STATUS = 0)
    Begin
        Insert Into Com_PresupuestoCredito (Periodo, Bodega, idPeriodoMes, Valor)
        Select @Periodo, @Bodega, idPeriodoMes, 0 From PeriodoMes Order By idPeriodoMes
        Fetch Next From cCom_EmpresaCredito Into @Bodega
    End
    Close cCom_EmpresaCredito
    DeAllocate cCom_EmpresaCredito

    /* GARANTIA EXTENDIDA */
    Declare cCom_EmpresaGEx Cursor For
        Select Bodega From Bodega
        Where  Activo = 1 And Factura = 1 And Credito = 1
          And  CONVERT(Date, ISNULL(FechaCierre, '20301231')) > '20241231'
        Order By Nombre

    OPEN cCom_EmpresaGEx
    Fetch Next From cCom_EmpresaGEx Into @Bodega
    While (@@FETCH_STATUS = 0)
    Begin
        Insert Into Com_PresupuestoEmpresaGE (Periodo, Bodega, idPeriodoMes, Valor)
        Select @Periodo, @Bodega, idPeriodoMes, 0 From PeriodoMes Order By idPeriodoMes
        Fetch Next From cCom_EmpresaGEx Into @Bodega
    End
    Close cCom_EmpresaGEx
    DeAllocate cCom_EmpresaGEx

    /* META CONSULTA CEDULAS */
    Declare cCom_EmpresaConsultaCedulas Cursor For
        Select Bodega From Bodega
        Where  Activo = 1 And Factura = 1 And Credito = 1
          And  CONVERT(Date, ISNULL(FechaCierre, '20301231')) > '20241231'
        Order By Nombre

    OPEN cCom_EmpresaConsultaCedulas
    Fetch Next From cCom_EmpresaConsultaCedulas Into @Bodega
    While (@@FETCH_STATUS = 0)
    Begin
        Insert Into Com_MetaConsultaCedulas (Periodo, Bodega, idPeriodoMes, Valor)
        Select @Periodo, @Bodega, idPeriodoMes, 0 From PeriodoMes Order By idPeriodoMes
        Fetch Next From cCom_EmpresaConsultaCedulas Into @Bodega
    End
    Close cCom_EmpresaConsultaCedulas
    DeAllocate cCom_EmpresaConsultaCedulas

END
GO

-- ---------------------------------------------------------------------------
-- 13. Com_GeneraPresupuestoEmpresaDesdePeriodoAnterior
-- ---------------------------------------------------------------------------
PRINT '>>> Com_GeneraPresupuestoEmpresaDesdePeriodoAnterior'
GO
IF OBJECT_ID('[dbo].[Com_GeneraPresupuestoEmpresaDesdePeriodoAnterior]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_GeneraPresupuestoEmpresaDesdePeriodoAnterior]
GO

CREATE PROCEDURE [dbo].[Com_GeneraPresupuestoEmpresaDesdePeriodoAnterior]
    @Periodo       Int,
    @PE            Int,
    @PersonalLocal Bit,
    @PL            Int,
    @AsigVend      Bit,
    @AV            Int
AS
BEGIN
    SET NOCOUNT ON;

    /* GENERA PRESUPUESTO EMPRESA */
    Insert Into Com_Empresa (Periodo, Bodega, idPeriodoMes, Valor)
    Select @Periodo, c.Bodega, idPeriodoMes,
        Case @PE
            When 1 Then Valor
            When 2 Then Case idPeriodoMes When 1  Then Valor Else 0 End
            When 3 Then Case idPeriodoMes When 12 Then Valor Else 0 End
        End
    From Com_Empresa c
    Inner Join Bodega b On b.Bodega = c.Bodega
    Where Periodo = @Periodo - 1
      And DATEPART(YEAR, ISNULL(b.FechaCierre, '20301231')) > @Periodo

    Insert Into Com_PresupuestoCredito (Periodo, Bodega, idPeriodoMes, Valor)
    Select @Periodo, c.Bodega, idPeriodoMes,
        Case @PE
            When 1 Then Valor
            When 2 Then Case idPeriodoMes When 1  Then Valor Else 0 End
            When 3 Then Case idPeriodoMes When 12 Then Valor Else 0 End
        End
    From Com_PresupuestoCredito c
    Inner Join Bodega b On b.Bodega = c.Bodega
    Where Periodo = @Periodo - 1
      And DATEPART(YEAR, ISNULL(b.FechaCierre, '20301231')) > @Periodo

    /* GENERA PERSONAL POR LOCAL */
    Declare @PeriodoInicial Varchar(7)
    Set @PeriodoInicial = '01/' + CAST(@Periodo As Varchar(4))

    If (@PersonalLocal = 1)
    Begin
        If (@PL = 1)
        Begin
            Insert Into Com_PersonalLocales (Periodo, Bodega, idCargo, Cantidad)
            Select SUBSTRING(Periodo, 1, 2) + '/' + CAST(@Periodo As Varchar(4)), p.Bodega, idCargo, Cantidad
            From Com_PersonalLocales p
            Inner Join Bodega b On b.Bodega = p.Bodega
            Where SUBSTRING(Periodo, 4, 7) = @Periodo - 1
              And DATEPART(YEAR, ISNULL(b.FechaCierre, '20301231')) > @Periodo
        End
        Else
        Begin
            If (@PL = 2) Set @PeriodoInicial = '01/' + CAST(@Periodo - 1 As Varchar(4))
            If (@PL = 3) Set @PeriodoInicial = '12/' + CAST(@Periodo - 1 As Varchar(4))

            Insert Into Com_PersonalLocales (Periodo, Bodega, idCargo, Cantidad)
            Select SUBSTRING(Periodo, 1, 2) + '/' + CAST(@Periodo As Varchar(4)), p.Bodega, idCargo, Cantidad
            From Com_PersonalLocales p
            Inner Join Bodega b On b.Bodega = p.Bodega
            Where Periodo = @PeriodoInicial
              And DATEPART(YEAR, ISNULL(b.FechaCierre, '20301231')) > @Periodo
        End
    End
    Else
    Begin
        Exec Com_GeneraPersonalLocales @PeriodoInicial
    End

    /* GENERA PRESUPUESTO MENSUAL POR LOCAL */
    Insert Into Com_PresupuestoMensualLocales (Periodo, Bodega, idPeriodoMes, idCargo, Valor)
    Select @Periodo, p.Bodega, idPeriodoMes, idCargo,
        Case @PE
            When 1 Then Valor
            When 2 Then Case idPeriodoMes When 1  Then Valor Else 0 End
            When 3 Then Case idPeriodoMes When 12 Then Valor Else 0 End
        End
    From Com_PresupuestoMensualLocales p
    Inner Join Bodega b On b.Bodega = p.Bodega
    Where Periodo = @Periodo - 1
      And DATEPART(YEAR, ISNULL(b.FechaCierre, '20301231')) > @Periodo

    If (@AsigVend = 1)
    Begin
        If (@AV = 1)
        Begin
            Insert Into Com_AsignacionDeVendedores (Periodo, Bodega, idCargo, idPersonal, idCom_Rango, Desde, Hasta, Diario, Mensual, idCom_Estado)
            Select SUBSTRING(Periodo, 1, 2) + '/' + CAST(@Periodo As Varchar(4)),
                   a.Bodega, idCargo, idPersonal, idCom_Rango,
                   DATEADD(YEAR, 1, Desde), DATEADD(YEAR, 1, Hasta),
                   Diario, Mensual, idCom_Estado
            From Com_AsignacionDeVendedores a
            Inner Join Bodega b On b.Bodega = a.Bodega
            Where SUBSTRING(Periodo, 4, 7) = @Periodo - 1
              And DATEPART(YEAR, ISNULL(b.FechaCierre, '20301231')) > @Periodo
              And idCom_Estado = 1
        End
    End

END
GO

-- ---------------------------------------------------------------------------
-- 14. Com_ConsultaPresupuestoCarpa
-- ---------------------------------------------------------------------------
PRINT '>>> Com_ConsultaPresupuestoCarpa'
GO
IF OBJECT_ID('[dbo].[Com_ConsultaPresupuestoCarpa]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ConsultaPresupuestoCarpa]
GO

CREATE PROCEDURE [dbo].[Com_ConsultaPresupuestoCarpa]
    @Periodo Int
AS
BEGIN
    SET NOCOUNT ON;

    Declare @sSQL     Varchar(MAX) = ''
    Declare @Columnas Varchar(Max) = ''

    Create Table #Meses (Mes Varchar(100))
    Declare @sSQLCC Varchar(MAX)
    Set @sSQLCC = 'Insert Into #Meses (Mes) Select Descripcion From PeriodoMes Order By idPeriodoMes'
    Exec (@sSQLCC)

    Select @Columnas = COALESCE(@Columnas + '[' + Mes + '],', '')
    FROM (Select Distinct Mes From #Meses) as a

    Set @Columnas = LEFT(@Columnas, LEN(@Columnas) - 1)

    Set @sSQL =
        'SELECT * FROM
        (Select Nombre, Descripcion, c.Valor, CONVERT(DECIMAL(18, 2), 0) As TOTAL
         From Com_PresupuestoCarpa c
         Inner Join Bodega b On b.bodega = c.Bodega
         Inner Join PeriodoMes p On p.idPeriodoMes = c.idPeriodoMes
         Where c.Periodo = ' + CAST(@Periodo As Varchar(4)) + ') AS SourceTable
        PIVOT
        (SUM(Valor)
         FOR Descripcion IN (' + @Columnas + ')) AS PivotTable;'

    Exec (@sSQL)
END
GO

-- ---------------------------------------------------------------------------
-- 15. Com_ConsultaPresupuestoAccesorios
-- ---------------------------------------------------------------------------
PRINT '>>> Com_ConsultaPresupuestoAccesorios'
GO
IF OBJECT_ID('[dbo].[Com_ConsultaPresupuestoAccesorios]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ConsultaPresupuestoAccesorios]
GO

CREATE PROCEDURE [dbo].[Com_ConsultaPresupuestoAccesorios]
    @Periodo Int
AS
BEGIN
    SET NOCOUNT ON;

    Declare @sSQL     Varchar(MAX) = ''
    Declare @Columnas Varchar(Max) = ''

    Create Table #Meses (Mes Varchar(100))
    Declare @sSQLCC Varchar(MAX)
    Set @sSQLCC = 'Insert Into #Meses (Mes) Select Descripcion From PeriodoMes Order By idPeriodoMes'
    Exec (@sSQLCC)

    Select @Columnas = COALESCE(@Columnas + '[' + Mes + '],', '')
    FROM (Select Distinct Mes From #Meses) as a

    Set @Columnas = LEFT(@Columnas, LEN(@Columnas) - 1)

    Set @sSQL =
        'SELECT * FROM
        (Select Nombre, Descripcion, e.Valor, CONVERT(DECIMAL(18, 2), 0) As TOTAL
         From Com_Accesorios e
         Inner Join Bodega b On b.bodega = e.Bodega
         Inner Join PeriodoMes p On p.idPeriodoMes = e.idPeriodoMes
         Where e.Periodo = ' + CAST(@Periodo As Varchar(4)) + ') AS SourceTable
        PIVOT
        (SUM(Valor)
         FOR Descripcion IN (' + @Columnas + ')) AS PivotTable;'

    Exec (@sSQL)
END
GO

-- ---------------------------------------------------------------------------
-- 16. Com_ConsultaMetaConsultaCedulas
-- ---------------------------------------------------------------------------
PRINT '>>> Com_ConsultaMetaConsultaCedulas'
GO
IF OBJECT_ID('[dbo].[Com_ConsultaMetaConsultaCedulas]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ConsultaMetaConsultaCedulas]
GO

CREATE PROCEDURE [dbo].[Com_ConsultaMetaConsultaCedulas]
    @Periodo Int
AS
BEGIN
    SET NOCOUNT ON;

    Declare @sSQL     Varchar(MAX) = ''
    Declare @Columnas Varchar(Max) = ''

    Create Table #Meses (Mes Varchar(100))
    Declare @sSQLCC Varchar(MAX)
    Set @sSQLCC = 'Insert Into #Meses (Mes) Select Descripcion From PeriodoMes Order By idPeriodoMes'
    Exec (@sSQLCC)

    Select @Columnas = COALESCE(@Columnas + '[' + Mes + '],', '')
    FROM (Select Distinct Mes From #Meses) as a

    Set @Columnas = LEFT(@Columnas, LEN(@Columnas) - 1)

    Set @sSQL =
        'SELECT * FROM
        (Select Nombre, Descripcion, e.Valor, CONVERT(DECIMAL(18, 2), 0) As TOTAL
         From Com_MetaConsultaCedulas e
         Inner Join Bodega b On b.bodega = e.Bodega
         Inner Join PeriodoMes p On p.idPeriodoMes = e.idPeriodoMes
         Where e.Periodo = ' + CAST(@Periodo As Varchar(4)) + ') AS SourceTable
        PIVOT
        (SUM(Valor)
         FOR Descripcion IN (' + @Columnas + ')) AS PivotTable;'

    Exec (@sSQL)
END
GO

-- ---------------------------------------------------------------------------
-- 17. Com_ConsultaPresupuestoMotos
-- ---------------------------------------------------------------------------
PRINT '>>> Com_ConsultaPresupuestoMotos'
GO
IF OBJECT_ID('[dbo].[Com_ConsultaPresupuestoMotos]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ConsultaPresupuestoMotos]
GO

CREATE PROCEDURE [dbo].[Com_ConsultaPresupuestoMotos]
    @Periodo Int
AS
BEGIN
    SET NOCOUNT ON;

    Declare @sSQL     Varchar(MAX) = ''
    Declare @Columnas Varchar(Max) = ''

    Create Table #Meses (Mes Varchar(100))
    Declare @sSQLCC Varchar(MAX)
    Set @sSQLCC = 'Insert Into #Meses (Mes) Select Descripcion From PeriodoMes Order By idPeriodoMes'
    Exec (@sSQLCC)

    Select @Columnas = COALESCE(@Columnas + '[' + Mes + '],', '')
    FROM (Select Distinct Mes From #Meses) as a

    Set @Columnas = LEFT(@Columnas, LEN(@Columnas) - 1)

    Set @sSQL =
        'SELECT * FROM
        (Select Nombre, Descripcion, e.Valor, CONVERT(DECIMAL(18, 2), 0) As TOTAL
         From Com_PresupuestoMotos e
         Inner Join Bodega b On b.bodega = e.Bodega
         Inner Join PeriodoMes p On p.idPeriodoMes = e.idPeriodoMes
         Where e.Periodo = ' + CAST(@Periodo As Varchar(4)) + ') AS SourceTable
        PIVOT
        (SUM(Valor)
         FOR Descripcion IN (' + @Columnas + ')) AS PivotTable;'

    Exec (@sSQL)
END
GO

-- ---------------------------------------------------------------------------
-- 18. Com_ActualizaValoresConsultaCedulas
-- ---------------------------------------------------------------------------
PRINT '>>> Com_ActualizaValoresConsultaCedulas'
GO
IF OBJECT_ID('[dbo].[Com_ActualizaValoresConsultaCedulas]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ActualizaValoresConsultaCedulas]
GO

CREATE PROCEDURE [dbo].[Com_ActualizaValoresConsultaCedulas]
    @Periodo       Int,
    @Local         Varchar(200),
    @Mes           Varchar(50),
    @Valor         Decimal(18, 2)
AS
BEGIN
    SET NOCOUNT ON;

    Declare @Bodega Int, @idPeriodoMes Int

    Select @Bodega       = Bodega       From Bodega    Where Nombre       = @Local
    Select @idPeriodoMes = idPeriodoMes From PeriodoMes Where Descripcion = @Mes

    Update Com_MetaConsultaCedulas
    Set    Valor = @Valor
    Where  Periodo      = @Periodo
      And  Bodega       = @Bodega
      And  idPeriodoMes = @idPeriodoMes

    Declare @Ceros         Varchar(2),
            @PeriodoAV     Varchar(7),
            @DiasMes       Int,
            @PeriodoActual Varchar(8)

    Set @Ceros         = SUBSTRING('00', 1, 2 - LEN(@idPeriodoMes))
    Set @PeriodoAV     = @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '/' + CAST(@Periodo As Varchar(4))
    Set @PeriodoActual = CONVERT(varchar(4), @Periodo) + @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '01'
    Set @DiasMes       = DATEPART(DAY, DATEADD(s, -1, DATEADD(mm, DATEDIFF(m, 0, CONVERT(Datetime, @PeriodoActual, 112)) + 1, 0)))

    If Exists(Select idCom_AsignacionDeVendedores From Com_AsignacionDeVendedores Where Bodega = @Bodega And Periodo = @PeriodoAV)
    Begin
        Update Com_AsignacionDeVendedores
        Set    MetaConsultaCedulas = dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Periodo, @idPeriodoMes, @Bodega, idCargo, 7) / @DiasMes * NumeroDias
        Where  Bodega       = @Bodega
          And  Periodo      = @PeriodoAV
          And  idCom_Estado In (1, 3)
    End

    If Exists(Select idCom_AsignacionJefes From Com_AsignacionJefes Where Bodega = @Bodega And Periodo = @PeriodoAV)
        Update Com_AsignacionJefes
        Set    MetaConsultaCedulas = (@Valor / @DiasMes) * Dias
        Where  Bodega  = @Bodega
          And  Periodo = @PeriodoAV

END
GO

-- ---------------------------------------------------------------------------
-- 19. Com_AgregaLocalNuevo
-- ---------------------------------------------------------------------------
PRINT '>>> Com_AgregaLocalNuevo'
GO
IF OBJECT_ID('[dbo].[Com_AgregaLocalNuevo]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_AgregaLocalNuevo]
GO

CREATE PROCEDURE [dbo].[Com_AgregaLocalNuevo]
    @Periodo Varchar(7),
    @Bodega  Int
AS
BEGIN
    SET NOCOUNT ON;

    Declare @PeriodoEmpresa Int
    Set @PeriodoEmpresa = SUBSTRING(@Periodo, 4, 7)

    /* Mensual */
    Insert Into Com_Empresa (Periodo, Bodega, idPeriodoMes, Valor)
    Select @PeriodoEmpresa, @Bodega, idPeriodoMes, 0 From PeriodoMes Order By idPeriodoMes

    /* Carpa */
    Insert Into Com_PresupuestoCarpa (Periodo, Bodega, idPeriodoMes, Valor)
    Select @PeriodoEmpresa, @Bodega, idPeriodoMes, 0 From PeriodoMes Order By idPeriodoMes

    /* Credito */
    Insert Into Com_PresupuestoCredito (Periodo, Bodega, idPeriodoMes, Valor)
    Select @PeriodoEmpresa, @Bodega, idPeriodoMes, 0 From PeriodoMes Order By idPeriodoMes

    /* Garantia Extendida */
    Insert Into Com_PresupuestoEmpresaGE (Periodo, Bodega, idPeriodoMes, Valor)
    Select @PeriodoEmpresa, @Bodega, idPeriodoMes, 0 From PeriodoMes Order By idPeriodoMes

    /* Accesorios */
    Insert Into Com_Accesorios (Periodo, Bodega, idPeriodoMes, Valor)
    Select @PeriodoEmpresa, @Bodega, idPeriodoMes, 0 From PeriodoMes Order By idPeriodoMes

    /* Personal Por Local */
    Insert Into Com_PersonalLocales (Periodo, Bodega, idCargo, Cantidad)
    Select @Periodo, @Bodega, idCargo, 0
    From   Com_CargosDeVentas
    Where  AV = 1 And idCargo Not In (21, 22, 23)
    Order By Orden

END
GO

-- ---------------------------------------------------------------------------
-- 20. Com_CopiaMesAnterior
-- ---------------------------------------------------------------------------
PRINT '>>> Com_CopiaMesAnterior'
GO
IF OBJECT_ID('[dbo].[Com_CopiaMesAnterior]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_CopiaMesAnterior]
GO

-- Exec Com_CopiaMesAnterior '03/2026'
CREATE PROCEDURE [dbo].[Com_CopiaMesAnterior]
    @Periodo Varchar(7)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRANSACTION;

    BEGIN TRY

        Declare @Desde         Date,
                @Hasta         Date,
                @Mes           Int,
                @Año           Int,
                @PeriodoNuevo  Varchar(7),
                @Ceros         Varchar(2),
                @DiasMes       Int,
                @Numero        Int

        Select Top 1 @Numero = Numero From Com_AsignacionDeVendedores Where Periodo = @Periodo
        Set @Numero = @Numero + 1

        Set @Mes  = SUBSTRING(@Periodo, 1, 2)
        Set @Mes  = @Mes + 1
        Set @Año  = SUBSTRING(@Periodo, 4, 7)
        Set @Ceros = SUBSTRING('00', 1, 2 - LEN(@Mes))
        Set @PeriodoNuevo = @Ceros + CONVERT(Varchar(2), @Mes) + '/' + CONVERT(Varchar(4), @Año)
        Set @Desde   = SUBSTRING(@PeriodoNuevo, 4, 7) + SUBSTRING(@PeriodoNuevo, 1, 2) + '01'
        Set @Hasta   = DATEADD(s, -1, DATEADD(mm, DATEDIFF(m, 0, @Desde) + 1, 0))
        Set @DiasMes = DATEDIFF(DAY, @Desde, @Hasta) + 1

        /* PRESUPUESTO MENSUAL */
        Update Com_Empresa
        Set    Valor = (Select Valor From Com_Empresa Where idPeriodoMes = (@Mes - 1) And Periodo = @Año And Bodega = c.Bodega)
        From   Com_Empresa c
        Where  idPeriodoMes = @Mes
          And  Periodo = @Año
          And  Bodega In (Select Bodega From Bodega Where ISNULL(FechaCierre, '20301231') >= @Hasta)

        /* CARPA */
        Update Com_PresupuestoMensualCarpa
        Set    Valor = (Select Valor From Com_PresupuestoMensualCarpa Where idPeriodoMes = (@Mes - 1) And Periodo = @Año And Bodega = c.Bodega)
        From   Com_PresupuestoMensualCarpa c
        Where  idPeriodoMes = @Mes
          And  Periodo = @Año
          And  Bodega In (Select Bodega From Bodega Where ISNULL(FechaCierre, '20301231') >= @Hasta)

        /* CREDITO */
        Update Com_PresupuestoCredito
        Set    Valor = (Select Valor From Com_PresupuestoCredito Where idPeriodoMes = (@Mes - 1) And Periodo = @Año And Bodega = c.Bodega)
        From   Com_PresupuestoCredito c
        Where  idPeriodoMes = @Mes
          And  Periodo = @Año
          And  Bodega In (Select Bodega From Bodega Where ISNULL(FechaCierre, '20301231') >= @Hasta)

        /* GARANTIA EXTENDIDA */
        Update Com_PresupuestoEmpresaGE
        Set    Valor = (Select Valor From Com_PresupuestoEmpresaGE Where idPeriodoMes = (@Mes - 1) And Periodo = @Año And Bodega = c.Bodega)
        From   Com_PresupuestoEmpresaGE c
        Where  idPeriodoMes = @Mes
          And  Periodo = @Año
          And  Bodega In (Select Bodega From Bodega Where ISNULL(FechaCierre, '20301231') >= @Hasta)

        /* PRESUPUESTO ACCESORIOS */
        Update Com_Accesorios
        Set    Valor = (Select Valor From Com_Accesorios Where idPeriodoMes = (@Mes - 1) And Periodo = @Año And Bodega = c.Bodega)
        From   Com_Accesorios c
        Where  idPeriodoMes = @Mes
          And  Periodo = @Año
          And  Bodega In (Select Bodega From Bodega Where ISNULL(FechaCierre, '20301231') >= @Hasta)

        /* PRESUPUESTO MOTOS */
        Update Com_PresupuestoMotos
        Set    Valor = (Select Valor From Com_PresupuestoMotos Where idPeriodoMes = (@Mes - 1) And Periodo = @Año And Bodega = c.Bodega)
        From   Com_PresupuestoMotos c
        Where  idPeriodoMes = @Mes
          And  Periodo = @Año
          And  Bodega In (Select Bodega From Bodega Where ISNULL(FechaCierre, '20301231') >= @Hasta)

        /* META CONSULTA CEDULAS */
        Update Com_MetaConsultaCedulas
        Set    Valor = (Select Valor From Com_MetaConsultaCedulas Where idPeriodoMes = (@Mes - 1) And Periodo = @Año And Bodega = c.Bodega)
        From   Com_MetaConsultaCedulas c
        Where  idPeriodoMes = @Mes
          And  Periodo = @Año
          And  Bodega In (Select Bodega From Bodega Where ISNULL(FechaCierre, '20301231') >= @Hasta)

        /* PERSONAL POR LOCAL */
        Declare @Bodega Int
        Declare oCom_PersonalLocales Cursor For
            Select Bodega From Bodega
            Where  Activo = 1 And Factura = 1 And Credito = 1
              And  ISNULL(FechaCierre, '20301231') >= @Hasta
            Order By Nombre
        Open oCom_PersonalLocales
        Fetch Next From oCom_PersonalLocales Into @Bodega
        While (@@FETCH_STATUS = 0)
        Begin
            If Not Exists(Select 1 From Com_PersonalLocales Where Periodo = @PeriodoNuevo And Bodega = @Bodega)
                Insert Into Com_PersonalLocales (Periodo, Bodega, idCargo, Cantidad)
                Select @PeriodoNuevo, @Bodega, idCargo, 0
                From   Com_CargosDeVentas
                Where  AV = 1 And idCargo Not In (21, 22, 23)
                Order By Orden
            Fetch Next From oCom_PersonalLocales Into @Bodega
        End
        Close oCom_PersonalLocales
        DeAllocate oCom_PersonalLocales

        Update Com_PersonalLocales
        Set    Cantidad = (Select SUM(Cantidad) From Com_PersonalLocales Where Periodo = @Periodo And Bodega = c.Bodega And idCargo = c.idCargo)
        From   Com_PersonalLocales c
        Where  Periodo = @PeriodoNuevo

        /* ASIGNACION DE VENDEDORES - mes completo */
        Insert Into Com_AsignacionDeVendedores
            (Periodo, SumaPuerteo, Bodega, idPersonal, idCom_Rango, idCargo, Desde, Hasta,
             NumeroDias, MetaMensual, MensualCarpa, MetaMensualCredito, MetaMensualGEx,
             MetaMensualAccesorios, MetaMensualMotos, idCom_Estado, Numero, EnvioMail, Estrellas)
        Select @PeriodoNuevo, SumaPuerteo, c.Bodega, c.idPersonal, c.idCom_Rango, c.idCargo,
               @Desde, @Hasta, @DiasMes,
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 1),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 5),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 3),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 4),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 6),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 8),
               n.idCom_Estado, @Numero, 1, c.Estrellas
        From   Com_AsignacionDeVendedores c
        Inner Join Nomina n On n.idPersonal = c.idPersonal
        Inner Join Com_Empresa e On e.idPeriodoMes = @Mes And e.Periodo = @Año And e.Bodega = c.Bodega
        Where  c.Periodo = @Periodo
          And  CONVERT(Date, Hasta) = DATEADD(DAY, -1, @Desde)
          And  n.idCom_Estado In (1, 3)

        /* ASIGNACION DE VENDEDORES - reingresos (salidas tipo > 2) */
        Insert Into Com_AsignacionDeVendedores
            (Periodo, SumaPuerteo, Bodega, idPersonal, idCom_Rango, idCargo, Desde, Hasta,
             NumeroDias, MetaMensual, MensualCarpa, MetaMensualCredito, MetaMensualGEx,
             MetaMensualAccesorios, MetaMensualMotos, idCom_Estado, Numero, EnvioMail, Estrellas)
        Select @PeriodoNuevo, SumaPuerteo, c.Bodega, c.idPersonal, c.idCom_Rango, c.idCargo,
               DATEADD(DAY, 1, ns.Hasta), @Hasta, DATEDIFF(DAY, ns.Hasta, @Hasta),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 1),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 5),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 3),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 4),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 6),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 8),
               n.idCom_Estado, @Numero, 1, c.Estrellas
        From   Com_AsignacionDeVendedores c
        Inner Join Nomina n          On n.idPersonal  = c.idPersonal
        Inner Join NominaSalidas ns  On ns.idNomina   = n.idNomina
        Inner Join Com_Empresa e     On e.idPeriodoMes = @Mes And e.Periodo = @Año And e.Bodega = c.Bodega
        Where  c.Periodo = @Periodo
          And  CONVERT(Date, ns.Hasta) Between DATEADD(DAY, -1, @Desde) And @Hasta
          And  n.idCom_Estado In (1, 3)
          And  ns.idNominaTiposDeSalidas > 2
          And  c.idCargo > 1

        /* ASIGNACION DE VENDEDORES - reingresos (salidas tipo = 2, cargo = 1) */
        Insert Into Com_AsignacionDeVendedores
            (Periodo, SumaPuerteo, Bodega, idPersonal, idCom_Rango, idCargo, Desde, Hasta,
             NumeroDias, MetaMensual, MensualCarpa, MetaMensualCredito, MetaMensualGEx,
             MetaMensualAccesorios, MetaMensualMotos, idCom_Estado, Numero, EnvioMail, Estrellas)
        Select @PeriodoNuevo, SumaPuerteo, c.Bodega, c.idPersonal, c.idCom_Rango, c.idCargo,
               DATEADD(DAY, 1, ns.Hasta), @Hasta, DATEDIFF(DAY, ns.Hasta, @Hasta),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 1),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 5),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 3),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 4),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 6),
               dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Año, @Mes, c.Bodega, c.idCargo, 8),
               n.idCom_Estado, @Numero, 1, c.Estrellas
        From   Com_AsignacionDeVendedores c
        Inner Join Nomina n          On n.idPersonal  = c.idPersonal
        Inner Join NominaSalidas ns  On ns.idNomina   = n.idNomina
        Inner Join Com_Empresa e     On e.idPeriodoMes = @Mes And e.Periodo = @Año And e.Bodega = c.Bodega
        Where  c.Periodo = @Periodo
          And  CONVERT(Date, ns.Hasta) Between DATEADD(DAY, -1, @Desde) And @Hasta
          And  n.idCom_Estado In (1, 3)
          And  ns.idNominaTiposDeSalidas = 2
          And  c.idCargo = 1

        /* ASIGNACION DE GERENTES / JEFES ZONALES / SUPERVISORES */
        Insert Into Com_AsignacionJefes
            (idPersonal, Periodo, Bodega, Desde, Hasta, Dias,
             MetaMensual, MetaMensualCarpa, MetaMensualCredito, MetaMensualGex,
             MetaMensualAccesorios, MetaMensualMotos)
        Select c.idPersonal, @PeriodoNuevo, c.Bodega, @Desde, @Hasta, @DiasMes,
               MetaMensual, MetaMensualCarpa, MetaMensualCredito, MetaMensualGex,
               c.MetaMensualAccesorios, MetaMensualMotos
        From   Com_AsignacionJefes c
        Inner Join Nomina n         On n.idPersonal = c.idPersonal
        Inner Join RolCargo rc      On rc.idCargo   = n.idCargo
        Inner Join Com_CargosDeVentas cv On cv.idCargo = rc.Com_IdCargo And AV = 0
        Inner Join Bodega b         On b.Bodega = c.Bodega
                                   And @Hasta Between FechaApertura And ISNULL(FechaCierre, '20301231')
                                   And b.Factura = 1 And Consignacion = 0 And Credito = 1
        Where  Periodo = @Periodo
          And  n.idCom_Estado In (1, 3)

        /* META CREDITO TIPO CLIENTE */
        Insert Into MetaCreditoTipoCliente (Bodega, idTipoCliente, Meta, Periodo)
        Select c.Bodega, c.idTipoCliente, c.Meta, @Hasta
        From   MetaCreditoTipoCliente c
        Where  CONVERT(Date, c.Periodo) = DATEADD(DAY, -1, @Desde)

        /* MONTO BONO CLIENTES BUENOS */
        Insert Into MontoBonoClientesBuenos (Valor, Desde, Hasta, idTipoCliente)
        Select c.Valor, @Desde, @Hasta, c.idTipoCliente
        From   MontoBonoClientesBuenos c
        Where  CONVERT(Date, c.Desde) = DATEADD(MONTH, -1, @Desde)
          And  CONVERT(Date, c.Hasta) = DATEADD(DAY,   -1, @Desde)

        IF @@TRANCOUNT > 0
            COMMIT TRANSACTION;

    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH

END
GO

PRINT '===== FIN ACTUALIZACION SPs PRESUPUESTO EMPRESA ====='
GO

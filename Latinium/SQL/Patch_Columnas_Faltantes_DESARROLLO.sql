-- =============================================================================
-- PASO 1: Ejecutar este script en DESARROLLO antes de re-crear los 3 SPs
-- Agrega las columnas que existen en POINT pero faltan en DESARROLLO
-- =============================================================================

PRINT '>>> Verificando columnas faltantes...'
GO

-- Com_AsignacionDeVendedores.MetaMensualMotos
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('Com_AsignacionDeVendedores')
      AND name = 'MetaMensualMotos'
)
BEGIN
    ALTER TABLE [dbo].[Com_AsignacionDeVendedores]
        ADD [MetaMensualMotos] Decimal(18, 2) NOT NULL CONSTRAINT DF_ComAsigVend_MetaMensualMotos DEFAULT(0)
    PRINT '  + Agregada: Com_AsignacionDeVendedores.MetaMensualMotos'
END
ELSE
    PRINT '  OK (ya existe): Com_AsignacionDeVendedores.MetaMensualMotos'
GO

-- Com_AsignacionDeVendedores.MetaConsultaCedulas
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('Com_AsignacionDeVendedores')
      AND name = 'MetaConsultaCedulas'
)
BEGIN
    ALTER TABLE [dbo].[Com_AsignacionDeVendedores]
        ADD [MetaConsultaCedulas] Decimal(18, 2) NOT NULL CONSTRAINT DF_ComAsigVend_MetaConsultaCedulas DEFAULT(0)
    PRINT '  + Agregada: Com_AsignacionDeVendedores.MetaConsultaCedulas'
END
ELSE
    PRINT '  OK (ya existe): Com_AsignacionDeVendedores.MetaConsultaCedulas'
GO

-- Com_AsignacionJefes.MetaConsultaCedulas
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('Com_AsignacionJefes')
      AND name = 'MetaConsultaCedulas'
)
BEGIN
    ALTER TABLE [dbo].[Com_AsignacionJefes]
        ADD [MetaConsultaCedulas] Decimal(18, 2) NOT NULL CONSTRAINT DF_ComAsigJefes_MetaConsultaCedulas DEFAULT(0)
    PRINT '  + Agregada: Com_AsignacionJefes.MetaConsultaCedulas'
END
ELSE
    PRINT '  OK (ya existe): Com_AsignacionJefes.MetaConsultaCedulas'
GO

-- Com_AsignacionJefes.MetaMensualMotos
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('Com_AsignacionJefes')
      AND name = 'MetaMensualMotos'
)
BEGIN
    ALTER TABLE [dbo].[Com_AsignacionJefes]
        ADD [MetaMensualMotos] Decimal(18, 2) NOT NULL CONSTRAINT DF_ComAsigJefes_MetaMensualMotos DEFAULT(0)
    PRINT '  + Agregada: Com_AsignacionJefes.MetaMensualMotos'
END
ELSE
    PRINT '  OK (ya existe): Com_AsignacionJefes.MetaMensualMotos'
GO

PRINT '>>> Columnas listas. Ahora ejecuta el bloque de abajo.'
GO

-- =============================================================================
-- PASO 2: Re-crear los 3 SPs que fallaron
-- =============================================================================

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

PRINT '===== LISTO - Los 3 SPs recreados correctamente ====='
GO

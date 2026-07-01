-- ============================================================
-- SCRIPT: Promoción "Segunda Visita Mitad de Precio"
-- BASE DE DATOS: PROSERVICIOS / PROSERVICIOS_PRUEBAS ÚNICAMENTE
-- Autor: Angel Pilatuna — Point Technology
-- Fecha: 2026-07-01
-- ============================================================
-- ORDEN DE EJECUCIÓN:
--   1. ALTER TABLE CompraPromocion       (columnas nuevas)
--   2. CREATE TABLE CuponesSegundaVisita (tabla de cupones)
--   3. INSERT TiposDeAnticipos           (tipo 6)
--   4. ALTER PROCEDURE GrabaPromociones  (2 parámetros nuevos)
--   5. CREATE PROCEDURE ValidaConfigPromoSegundaVisita
--   6. CREATE PROCEDURE ValidaSaldoCuponSegundaVisita
--   7. CREATE PROCEDURE GeneraCuponSegundaVisita
--   8. CREATE PROCEDURE AnulaCuponSegundaVisita
--   9. CREATE PROCEDURE VenceCuponesSegundaVisita (job nocturno)
-- ============================================================

USE [PROSERVICIOS]
GO

PRINT '============================================================'
PRINT 'PASO 1 — Columnas nuevas en CompraPromocion'
PRINT '============================================================'

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('CompraPromocion') AND name = 'ValidezCuponDias')
    ALTER TABLE CompraPromocion ADD ValidezCuponDias INT NULL DEFAULT(0)
ELSE PRINT 'CompraPromocion.ValidezCuponDias ya existe — omitido'

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('CompraPromocion') AND name = 'PorcentajeDescuento')
    ALTER TABLE CompraPromocion ADD PorcentajeDescuento DECIMAL(5,2) NULL DEFAULT(0)
ELSE PRINT 'CompraPromocion.PorcentajeDescuento ya existe — omitido'

PRINT 'PASO 1 completado'
GO

-- ============================================================
PRINT '============================================================'
PRINT 'PASO 2 — Tabla CuponesSegundaVisita'
PRINT '============================================================'

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('CuponesSegundaVisita') AND type = 'U')
BEGIN
    CREATE TABLE CuponesSegundaVisita (
        idCupon           INT IDENTITY(1,1) PRIMARY KEY,
        idAnticipo        INT NOT NULL,          -- FK Anticipos.idAnticipo
        idCompraOrigen    INT NOT NULL,           -- Factura donde se vendió el cupón
        idCliente         INT NOT NULL,
        idHabitacion      INT NOT NULL,
        idCategoria       INT NOT NULL,           -- Para validar misma categoría al canjear
        idBodega          INT NOT NULL,           -- Local donde se emitió
        ValorOriginal     DECIMAL(18,2) NOT NULL, -- Precio normal de la habitación
        PorcentajeDesc    DECIMAL(5,2)  NOT NULL, -- 50.00
        ValorDescuento    DECIMAL(18,2) NOT NULL, -- ValorOriginal * PorcentajeDesc / 100
        ValorCancelado    DECIMAL(18,2) NOT NULL, -- Lo que pagó por el cupón
        FechaEmision      DATETIME      NOT NULL DEFAULT(GETDATE()),
        FechaVencimiento  DATETIME      NOT NULL, -- FechaEmision + ValidezCuponDias
        -- Estados: 1=Pendiente, 3=Utilizado, 4=Anulado, 5=Vencido
        Estado            INT           NOT NULL DEFAULT(1),
        idCompraUso       INT           NULL,     -- Factura donde se canjeó
        FechaUso          DATETIME      NULL,
        idCajeroUso       INT           NULL,
        Observacion       VARCHAR(200)  NULL
    )
    PRINT 'Tabla CuponesSegundaVisita creada'
END
ELSE PRINT 'CuponesSegundaVisita ya existe — omitido'
GO

-- ============================================================
PRINT '============================================================'
PRINT 'PASO 3 — TiposDeAnticipos: Tipo 6 = Segunda Visita 50%'
PRINT '============================================================'

IF NOT EXISTS (SELECT 1 FROM TiposDeAnticipos WHERE Tipo = 6)
BEGIN
    INSERT INTO TiposDeAnticipos (Tipo, Descripcion, Reporte)
    VALUES (6, 'Segunda Visita 50%', 'RptCuponSegundaVisita.rpt')
    PRINT 'TiposDeAnticipos Tipo=6 insertado'
END
ELSE PRINT 'TiposDeAnticipos Tipo=6 ya existe — omitido'
GO

-- ============================================================
PRINT '============================================================'
PRINT 'PASO 4 — GrabaPromociones: agrega 2 parámetros opcionales'
PRINT '         (retrocompatible — llamadas existentes no cambian)'
PRINT '============================================================'
GO

ALTER PROCEDURE [dbo].[GrabaPromociones]
    @idPromocion          Int,
    @Tipo                 Varchar(20),
    @Desde                DateTime      = Null,
    @Hasta                DateTime      = Null,
    @Descripcion          Varchar(200)  = Null,
    @PrecioMinimo         Decimal(18,2) = Null,
    @PrecioMaximo         Decimal(18,2) = Null,
    @Cuotas               Int           = Null,
    @CuotaReal            Decimal(18,2) = Null,
    @CuotaCalculada       Decimal(18,2) = Null,
    @NumIngreso           Int           = Null,
    @Estado               Varchar(20)   = Null,
    @idArticulo           Int           = Null,
    @Variable             Bit           = 0,
    @SumaDescuento        Bit           = 0,
    @Valor                Bit           = 0,
    @Adicional            Bit           = 0,
    @Opcional             Bit           = 0,
    @Cupon                Bit           = 0,
    @Condicion            Bit           = 0,
    @ValidaCargaProducto  Bit           = 0,
    @HoraInicio           Datetime      = '20000101',
    @HoraFinaliza         Datetime      = '20000101',
    @PuntosFinal          Int           = 0,
    @TiempoEstadia        Int           = 0,
    @ToqueQueda           Bit           = 0,
    @Encuesta             Bit           = 0,
    @Preventa             Bit           = 0,
    -- Parámetros nuevos — Segunda Visita 50% (opcionales, default 0 = no aplica)
    @ValidezCuponDias     Int           = 0,
    @PorcentajeDescuento  Decimal(5,2)  = 0
AS
BEGIN
    SET NOCOUNT ON;

    Declare @CodPromocion Int

    If Not Exists(Select idArticulo From CompraPromocion Where idPromocion = @idPromocion)
        Begin
            PRINT ('NUEVO')
            Insert Into CompraPromocion (idArticulo, Codigo, NumIngreso, Estado, Valor, Variable, SumaDescuento, Opcional, ValidaCargaProducto, TiempoEstadia, ToqueQueda, Encuesta, Preventa)
                Values (@idArticulo, (Select Count(*) From CompraPromocion) + 1, 1, @Estado, @Valor, @Variable, @SumaDescuento, @Opcional, @ValidaCargaProducto, @TiempoEstadia, @ToqueQueda, @Encuesta, @Preventa)

            Set @CodPromocion = SCOPE_IDENTITY()
        End
    Else
        Begin
            Print ('UPDATE VALCP : ' + CAST(@ValidaCargaProducto AS VARCHAR(10)))

            Update CompraPromocion
            Set Tipo                = @Tipo,
                FechaDesde          = @Desde,
                FechaHasta          = @Hasta,
                Descripcion         = @Descripcion,
                PrecioMinimo        = @PrecioMinimo,
                PrecioMaximo        = @PrecioMaximo,
                PuntosInicio        = @Cuotas,
                CuotaReal           = @CuotaReal,
                CuotaCalculada      = @CuotaCalculada,
                NumIngreso          = @NumIngreso,
                Estado              = @Estado,
                idArticulo          = @idArticulo,
                Variable            = @Variable,
                SumaDescuento       = @SumaDescuento,
                Valor               = @Valor,
                Adicional           = @Adicional,
                Opcional            = @Opcional,
                Cupon               = @Cupon,
                HoraInicia          = @HoraInicio,
                HoraFinaliza        = @HoraFinaliza,
                Condicion           = @Condicion,
                ValidaCargaProducto = @ValidaCargaProducto,
                PuntosFinal         = @PuntosFinal,
                TiempoEstadia       = @TiempoEstadia,
                ToqueQueda          = @ToqueQueda,
                Encuesta            = @Encuesta,
                Preventa            = @Preventa,
                ValidezCuponDias    = @ValidezCuponDias,
                PorcentajeDescuento = @PorcentajeDescuento
            Where idPromocion = @idPromocion

            Set @CodPromocion = @idPromocion
        End

    Select @CodPromocion
END
GO
PRINT 'PASO 4 completado'
GO

-- ============================================================
PRINT '============================================================'
PRINT 'PASO 5 — ValidaConfigPromoSegundaVisita'
PRINT '         Devuelve la config de la promo activa para una'
PRINT '         fecha y bodega dadas. El C# lo usa para habilitar'
PRINT '         el botón en frmFacturacionHabitaciones.'
PRINT '============================================================'
GO

IF OBJECT_ID('ValidaConfigPromoSegundaVisita') IS NOT NULL
    DROP PROCEDURE ValidaConfigPromoSegundaVisita
GO

CREATE PROCEDURE [dbo].[ValidaConfigPromoSegundaVisita]
    @idBodega INT,
    @Fecha    DATETIME
AS
BEGIN
    SET NOCOUNT ON;

    -- Devuelve idPromocion, Descripcion, PorcentajeDescuento, ValidezCuponDias
    -- si existe una promo de tipo 'SEGUNDA VISITA' activa en esa bodega y ese día.
    -- Días: 1=Dom, 2=Lun, 3=Mar, 4=Mié, 5=Jue, 6=Vie, 7=Sab
    -- Viernes(6) y Sábado(7) están excluidos por regla de negocio.
    SELECT TOP 1
        cp.idPromocion,
        cp.Descripcion,
        cp.PorcentajeDescuento,
        cp.ValidezCuponDias
    FROM CompraPromocion cp
        INNER JOIN PromocionCiudadBodega pcb ON pcb.idPromocion = cp.idPromocion
        INNER JOIN PromocionDiasHabitacion pdh ON pdh.idPromocion = cp.idPromocion AND pdh.Activa = 1
    WHERE cp.Tipo   = 'SEGUNDA VISITA'
      AND cp.Estado = 'ACTIVO'
      AND @Fecha BETWEEN cp.FechaDesde AND cp.FechaHasta
      AND pcb.Bodega = @idBodega
      AND DATEPART(WEEKDAY, @Fecha) NOT IN (6, 7)   -- excluye Viernes y Sábado
      AND DATEPART(WEEKDAY, @Fecha) = pdh.idDia      -- día habilitado en la config
END
GO
PRINT 'PASO 5 completado'
GO

-- ============================================================
PRINT '============================================================'
PRINT 'PASO 6 — ValidaSaldoCuponSegundaVisita'
PRINT '         Verifica que el cliente tenga un cupón válido'
PRINT '         al momento del canje (segunda visita).'
PRINT '============================================================'
GO

IF OBJECT_ID('ValidaSaldoCuponSegundaVisita') IS NOT NULL
    DROP PROCEDURE ValidaSaldoCuponSegundaVisita
GO

CREATE PROCEDURE [dbo].[ValidaSaldoCuponSegundaVisita]
    @idCliente   INT,
    @idBodega    INT,
    @idCategoria INT,
    @Fecha       DATETIME
AS
BEGIN
    SET NOCOUNT ON;

    -- Devuelve: idCupon, idAnticipo, ValorDescuento, FechaVencimiento
    -- Valida: estado Pendiente(1), no vencido, mismo local, misma categoría,
    -- y que el día no sea Viernes(6) ni Sábado(7).
    SELECT
        c.idCupon,
        c.idAnticipo,
        c.ValorDescuento,
        c.FechaVencimiento,
        c.ValorOriginal,
        c.PorcentajeDesc
    FROM CuponesSegundaVisita c
        INNER JOIN Anticipos a ON a.idAnticipo = c.idAnticipo
    WHERE c.idCliente   = @idCliente
      AND c.Estado      = 1                          -- Pendiente
      AND c.FechaVencimiento >= @Fecha               -- No vencido
      AND c.idBodega    = @idBodega                  -- Mismo local
      AND c.idCategoria = @idCategoria               -- Misma categoría
      AND DATEPART(WEEKDAY, @Fecha) NOT IN (6, 7)    -- No Viernes ni Sábado
      AND a.Estado      = 1                          -- Anticipo aún activo
END
GO
PRINT 'PASO 6 completado'
GO

-- ============================================================
PRINT '============================================================'
PRINT 'PASO 7 — GeneraCuponSegundaVisita'
PRINT '         Crea el anticipo Tipo=6 y el registro en'
PRINT '         CuponesSegundaVisita al finalizar la venta.'
PRINT '============================================================'
GO

IF OBJECT_ID('GeneraCuponSegundaVisita') IS NOT NULL
    DROP PROCEDURE GeneraCuponSegundaVisita
GO

CREATE PROCEDURE [dbo].[GeneraCuponSegundaVisita]
    @idCompraOrigen      INT,
    @idCliente           INT,
    @idHabitacion        INT,
    @idCategoria         INT,
    @idBodega            INT,
    @ValorOriginal       DECIMAL(18,2),
    @PorcentajeDesc      DECIMAL(5,2),   -- ej. 50.00
    @ValidezCuponDias    INT,            -- ej. 30
    @RUC                 VARCHAR(20)     = '',
    @idTipoRuc           INT             = 5
AS
BEGIN
    SET NOCOUNT ON;

    Declare @ValorDescuento  DECIMAL(18,2)
    Declare @ValorCancelado  DECIMAL(18,2)
    Declare @FechaVenc       DATETIME
    Declare @idAnticipo      INT

    Set @ValorDescuento = ROUND(@ValorOriginal * @PorcentajeDesc / 100, 2)
    Set @ValorCancelado = @ValorDescuento
    Set @FechaVenc      = DATEADD(DAY, @ValidezCuponDias, GETDATE())

    -- 1. Crear el anticipo en la tabla Anticipos (Tipo=6 = Segunda Visita)
    INSERT INTO Anticipos (idCliente, ClienteProveedor, Tipo, Estado, Abono, Saldo, Fecha, idOrigen)
    VALUES (@idCliente, 0, 6, 1, @ValorCancelado, @ValorCancelado, GETDATE(), @idCompraOrigen)

    Set @idAnticipo = SCOPE_IDENTITY()

    -- 2. Registrar el cupón con todos sus datos
    INSERT INTO CuponesSegundaVisita
        (idAnticipo, idCompraOrigen, idCliente, idHabitacion, idCategoria, idBodega,
         ValorOriginal, PorcentajeDesc, ValorDescuento, ValorCancelado,
         FechaEmision, FechaVencimiento, Estado)
    VALUES
        (@idAnticipo, @idCompraOrigen, @idCliente, @idHabitacion, @idCategoria, @idBodega,
         @ValorOriginal, @PorcentajeDesc, @ValorDescuento, @ValorCancelado,
         GETDATE(), @FechaVenc, 1)

    -- Retorna el idAnticipo para que el C# lo use en ImpresionAutomaticaAnticipos
    Select @idAnticipo As idAnticipo
END
GO
PRINT 'PASO 7 completado'
GO

-- ============================================================
PRINT '============================================================'
PRINT 'PASO 8 — AnulaCuponSegundaVisita'
PRINT '         Anulación manual por supervisor.'
PRINT '============================================================'
GO

IF OBJECT_ID('AnulaCuponSegundaVisita') IS NOT NULL
    DROP PROCEDURE AnulaCuponSegundaVisita
GO

CREATE PROCEDURE [dbo].[AnulaCuponSegundaVisita]
    @idCupon      INT,
    @Observacion  VARCHAR(200) = ''
AS
BEGIN
    SET NOCOUNT ON;

    Declare @idAnticipo INT
    Select @idAnticipo = idAnticipo From CuponesSegundaVisita Where idCupon = @idCupon

    If (@idAnticipo IS NULL)
    Begin
        Select 0 As Ok, 'Cupón no encontrado' As Mensaje
        Return
    End

    If Exists(Select 1 From CuponesSegundaVisita Where idCupon = @idCupon And Estado <> 1)
    Begin
        Select 0 As Ok, 'El cupón ya fue utilizado, vencido o anulado' As Mensaje
        Return
    End

    -- Estado 4 = Anulado
    Update CuponesSegundaVisita Set Estado = 4, Observacion = @Observacion Where idCupon = @idCupon
    Update Anticipos Set Estado = 4, Saldo = 0 Where idAnticipo = @idAnticipo

    Select 1 As Ok, 'Cupón anulado correctamente' As Mensaje
END
GO
PRINT 'PASO 8 completado'
GO

-- ============================================================
PRINT '============================================================'
PRINT 'PASO 9 — VenceCuponesSegundaVisita'
PRINT '         Marca como Vencidos(5) los cupones Pendientes(1)'
PRINT '         cuya FechaVencimiento ya pasó.'
PRINT '         Se ejecuta como SQL Agent Job — frecuencia: diaria.'
PRINT '============================================================'
GO

IF OBJECT_ID('VenceCuponesSegundaVisita') IS NOT NULL
    DROP PROCEDURE VenceCuponesSegundaVisita
GO

CREATE PROCEDURE [dbo].[VenceCuponesSegundaVisita]
AS
BEGIN
    SET NOCOUNT ON;

    Declare @hoy DATETIME = GETDATE()

    -- Marcar cupones vencidos
    Update CuponesSegundaVisita
    Set Estado = 5
    Where Estado = 1
      And FechaVencimiento < @hoy

    -- Marcar los anticipos correspondientes como vencidos
    Update Anticipos
    Set Estado = 5, Saldo = 0
    Where idAnticipo IN (
        Select idAnticipo From CuponesSegundaVisita Where Estado = 5
    )
    And Estado = 1

    Select @@ROWCOUNT As CuponesVencidos
END
GO
PRINT 'PASO 9 completado'
GO

-- ============================================================
PRINT '============================================================'
PRINT 'VERIFICACIÓN FINAL'
PRINT '============================================================'

SELECT 'CompraPromocion.ValidezCuponDias'   AS Columna, CASE WHEN EXISTS(SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('CompraPromocion') AND name='ValidezCuponDias')   THEN 'OK' ELSE 'FALTA' END AS Estado
UNION ALL
SELECT 'CompraPromocion.PorcentajeDescuento', CASE WHEN EXISTS(SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('CompraPromocion') AND name='PorcentajeDescuento') THEN 'OK' ELSE 'FALTA' END
UNION ALL
SELECT 'Tabla CuponesSegundaVisita',          CASE WHEN OBJECT_ID('CuponesSegundaVisita') IS NOT NULL THEN 'OK' ELSE 'FALTA' END
UNION ALL
SELECT 'TiposDeAnticipos Tipo=6',             CASE WHEN EXISTS(SELECT 1 FROM TiposDeAnticipos WHERE Tipo=6) THEN 'OK' ELSE 'FALTA' END
UNION ALL
SELECT 'SP GrabaPromociones',                 CASE WHEN OBJECT_ID('GrabaPromociones') IS NOT NULL THEN 'OK' ELSE 'FALTA' END
UNION ALL
SELECT 'SP ValidaConfigPromoSegundaVisita',   CASE WHEN OBJECT_ID('ValidaConfigPromoSegundaVisita') IS NOT NULL THEN 'OK' ELSE 'FALTA' END
UNION ALL
SELECT 'SP ValidaSaldoCuponSegundaVisita',    CASE WHEN OBJECT_ID('ValidaSaldoCuponSegundaVisita') IS NOT NULL THEN 'OK' ELSE 'FALTA' END
UNION ALL
SELECT 'SP GeneraCuponSegundaVisita',         CASE WHEN OBJECT_ID('GeneraCuponSegundaVisita') IS NOT NULL THEN 'OK' ELSE 'FALTA' END
UNION ALL
SELECT 'SP AnulaCuponSegundaVisita',          CASE WHEN OBJECT_ID('AnulaCuponSegundaVisita') IS NOT NULL THEN 'OK' ELSE 'FALTA' END
UNION ALL
SELECT 'SP VenceCuponesSegundaVisita',        CASE WHEN OBJECT_ID('VenceCuponesSegundaVisita') IS NOT NULL THEN 'OK' ELSE 'FALTA' END
GO

PRINT '============================================================'
PRINT 'Script completado — PROSERVICIOS / PROSERVICIOS_PRUEBAS'
PRINT '============================================================'

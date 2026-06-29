-- CREQ: Cierre de Operaciones con Descuento Especial
-- SP de validación individual: recibe una fila del Excel y devuelve
-- si es válida o el motivo del error. No modifica nada en BD.

IF OBJECT_ID('dbo.CierreOperacionDescuentoValidar', 'P') IS NOT NULL
    DROP PROCEDURE dbo.CierreOperacionDescuentoValidar
GO

CREATE PROCEDURE [dbo].[CierreOperacionDescuentoValidar]
    @Cedula          VARCHAR(20),
    @Factura         VARCHAR(50),
    @Segmento        VARCHAR(100),
    @ValorDescuento  DECIMAL(18,2),
    -- Outputs
    @idCompra        INT           OUTPUT,
    @Cliente         VARCHAR(200)  OUTPUT,
    @SaldoReal       DECIMAL(18,4) OUTPUT,
    @ValorAjustado   DECIMAL(18,4) OUTPUT,
    @Error           VARCHAR(500)  OUTPUT
AS
BEGIN
    SET NOCOUNT ON

    SET @idCompra      = 0
    SET @Cliente       = ''
    SET @SaldoReal     = 0
    SET @ValorAjustado = @ValorDescuento
    SET @Error         = ''

    -- 1. Campos obligatorios
    IF LTRIM(RTRIM(ISNULL(@Cedula,  ''))) = '' OR
       LTRIM(RTRIM(ISNULL(@Factura, ''))) = ''
    BEGIN
        SET @Error = 'CEDULA o FACTURA vacíos'
        RETURN
    END

    IF ISNULL(@ValorDescuento, 0) <= 0
    BEGIN
        SET @Error = 'VALOR debe ser mayor a cero'
        RETURN
    END

    -- 2. Resolver segmento → idEntidadFinanciera
    DECLARE @idEF INT
    SELECT @idEF = idEntidadFinanciera
    FROM   Cre_EntidadFinanciera
    WHERE  LTRIM(RTRIM(UPPER(Nombre))) = LTRIM(RTRIM(UPPER(@Segmento)))

    IF ISNULL(@idEF, 0) = 0
    BEGIN
        SET @Error = 'SEGMENTO "' + @Segmento + '" no existe en Cre_EntidadFinanciera — verifique el nombre exacto'
        RETURN
    END

    -- 3. Buscar en Compra filtrando por segmento (idEntidadFinanciera) exacto
    -- NOTA: NO filtramos por Saldo>0 porque Compra.Saldo puede estar en 0
    --       aunque existan cuotas pendientes en Cre_TablaDeAmortizacion
    DECLARE @Saldo     DECIMAL(18,2),
            @RucCompra VARCHAR(20)

    SELECT TOP 1
        @idCompra  = c.idCompra,
        @Saldo     = ISNULL(c.Saldo, 0),
        @RucCompra = ISNULL(c.Ruc, ''),
        @Cliente   = ISNULL(c.Nombre, '')
    FROM Compra c
    WHERE c.Numero                = @Factura
      AND c.Ruc                   = @Cedula
      AND c.Borrar                = 0
      AND c.Estado               <> 6
      AND c.idTipoFactura        IN (1, 27)
      AND c.idFormaPago           = 9
      AND c.idEntidadFinanciera   = @idEF
    ORDER BY c.idCompra DESC

    IF ISNULL(@idCompra, 0) = 0
    BEGIN
        -- Dar diagnóstico: ¿existe con otro segmento?
        DECLARE @NombreEFEncontrado VARCHAR(100)
        SELECT TOP 1
            @NombreEFEncontrado = ef.Nombre
        FROM Compra c
        JOIN Cre_EntidadFinanciera ef ON ef.idEntidadFinanciera = c.idEntidadFinanciera
        WHERE c.Numero         = @Factura
          AND c.Ruc            = @Cedula
          AND c.Borrar         = 0
          AND c.Estado        <> 6
          AND c.idTipoFactura IN (1, 27)
          AND c.idFormaPago    = 9
        ORDER BY c.idCompra DESC

        IF @NombreEFEncontrado IS NOT NULL
            SET @Error = 'FACTURA ' + @Factura + ' existe en CREDIPOINT pero con segmento "'
                + @NombreEFEncontrado + '", no "' + @Segmento + '" — corrija el segmento en el Excel'
        ELSE
            SET @Error = 'FACTURA ' + @Factura + ' no encontrada para cédula ' + @Cedula
                + ' con segmento "' + @Segmento + '" (CREDIPOINT)'
        RETURN
    END

    -- 3. Cédula coincide (redundante con el WHERE, pero deja mensaje explícito)
    IF LTRIM(RTRIM(@RucCompra)) <> LTRIM(RTRIM(@Cedula))
    BEGIN
        SET @Error = 'CÉDULA NO CORRESPONDE a la factura (BD: ' + @RucCompra + ')'
        RETURN
    END

    -- 6. No procesado ya
    IF EXISTS (SELECT 1 FROM CierreOperacionDescuentoLog
               WHERE idCompra = @idCompra AND Estado = 'PROCESADO')
    BEGIN
        SET @Error = 'YA PROCESADO anteriormente'
        RETURN
    END

    -- 7. Sin deuda en abogados
    DECLARE @EstadoAbog INT
    SELECT @EstadoAbog = ISNULL(MAX(EstadoAbogados), 0)
    FROM Anticipos
    WHERE idOrigen = @idCompra

    IF ISNULL(@EstadoAbog, 0) > 0
    BEGIN
        SET @Error = 'DEUDA EN MANOS DE ABOGADOS (Estado=' + CAST(@EstadoAbog AS VARCHAR(5)) + ') — requiere autorización especial'
        RETURN
    END

    -- 8. Cuotas pendientes y saldo real en amortización
    DECLARE @CntCuotas INT

    SELECT @CntCuotas = COUNT(1),
           @SaldoReal  = ISNULL(SUM(Capital - ABCapital), 0)
    FROM   Cre_TablaDeAmortizacion
    WHERE  Vendida  = 0
      AND  Estado   NOT IN (2, 4)
      AND  idCompra = @idCompra

    IF ISNULL(@CntCuotas, 0) = 0
    BEGIN
        SET @Error = 'SIN CUOTAS PENDIENTES en la tabla de amortización'
        RETURN
    END

    IF ISNULL(@SaldoReal, 0) <= 0
    BEGIN
        SET @Error = 'SALDO REAL DE CUOTAS es cero — la operación ya está cerrada en amortización'
        RETURN
    END

    -- 9. Siempre usar el saldo real, igual que el proceso original
    SET @ValorAjustado = @SaldoReal

    -- 10. Sin anticipos pendientes sin aplicar
    DECLARE @AntPend INT
    SELECT @AntPend = COUNT(1)
    FROM Anticipos
    WHERE idOrigen = @idCompra AND Estado = 1

    IF ISNULL(@AntPend, 0) > 0
    BEGIN
        SET @Error = 'HAY ' + CAST(@AntPend AS VARCHAR(5)) + ' ANTICIPO(S) PENDIENTE(S) sin aplicar — liquidar primero'
        RETURN
    END

    -- Todo OK — @Error queda vacío
END
GO

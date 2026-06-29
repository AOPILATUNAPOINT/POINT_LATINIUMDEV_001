-- CREQ: Cierre de Operaciones con Descuento Especial
-- Sigue exactamente el mismo flujo que el proceso de cierre de cobranza estándar
-- (TEMPORAL_SALDOS_COBRANZAS). La única diferencia: se paga @ValorDescuento
-- en lugar del saldo completo, y las cuotas restantes se condonan.

ALTER PROCEDURE [dbo].[CierreOperacionDescuento]
    @Cedula          VARCHAR(20),
    @Factura         VARCHAR(50),
    @Segmento        VARCHAR(100),
    @ValorDescuento  DECIMAL(18,4),
    @NumAutorizacion VARCHAR(100),
    @UsuarioAutoriza VARCHAR(50),
    @UsuarioEjecuta  VARCHAR(50),
    @Mensaje         VARCHAR(500) OUTPUT
AS
BEGIN
    SET NOCOUNT ON

    DECLARE
        @idCompra            INT,
        @SaldoOriginal       DECIMAL(18,2),
        @RucCompra           VARCHAR(20),
        @Cliente             VARCHAR(200),
        @idEntidadFinanciera INT,
        @Confirmado          BIT,
        @idCliente           INT,
        @idMotivo            INT,
        @NumeroDeOperacion   VARCHAR(20),
        @Tipo                INT,
        @Bodega              INT  = 57,
        @Numero              VARCHAR(20),
        @Zero                VARCHAR(10),
        @idAnticipo          INT,
        @Detalle             VARCHAR(200),
        @Saldo               DECIMAL(18,2),
        @idCre_TablaDeAmort  INT,
        @Capital             DECIMAL(18,2),
        @Interes             INT,
        @Estado              INT,
        @Fecha               DATE = GETDATE()

    SET @Mensaje = ''

    -- --------------------------------------------------------
    -- 1. Resolver segmento → idEntidadFinanciera
    -- --------------------------------------------------------
    DECLARE @idEF INT
    SELECT @idEF = idEntidadFinanciera
    FROM   Cre_EntidadFinanciera
    WHERE  LTRIM(RTRIM(UPPER(Nombre))) = LTRIM(RTRIM(UPPER(@Segmento)))

    IF ISNULL(@idEF, 0) = 0
    BEGIN
        SET @Mensaje = 'SEGMENTO "' + @Segmento + '" no existe en Cre_EntidadFinanciera'
        RETURN
    END

    -- --------------------------------------------------------
    -- 2. Buscar la operación en Compra filtrada por segmento
    --    Compra.Saldo puede estar en 0 aunque haya cuotas pendientes
    --    en Cre_TablaDeAmortizacion — NO filtrar por Saldo>0
    -- --------------------------------------------------------
    SELECT TOP 1
        @idCompra            = c.idCompra,
        @SaldoOriginal       = ISNULL(c.Saldo, 0),
        @RucCompra           = ISNULL(c.Ruc, ''),
        @Cliente             = ISNULL(c.Nombre, ''),
        @idEntidadFinanciera = ISNULL(c.idEntidadFinanciera, 0),
        @Confirmado          = ISNULL(c.Confirmado, 0),
        @idCliente           = ISNULL(c.idCliente, 0),
        @idMotivo            = ISNULL(c.idMotivo, 0),
        @NumeroDeOperacion   = ISNULL(c.NumeroDeOperacion, '')
    FROM Compra c
    WHERE c.Numero                = @Factura
      AND c.Ruc                   = @Cedula
      AND c.Borrar                = 0
      AND c.Estado               <> 6
      AND c.idTipoFactura        IN (1, 27)
      AND c.idFormaPago           = 9
      AND c.idEntidadFinanciera   = @idEF
    ORDER BY c.idCompra DESC

    IF @idCompra IS NULL
    BEGIN
        SET @Mensaje = 'FACTURA ' + @Factura + ' NO ENCONTRADA para cédula ' + @Cedula
            + ' con segmento "' + @Segmento + '" (CREDIPOINT)'
        RETURN
    END

    -- --------------------------------------------------------
    -- 3. Validar monto básico (> 0)
    --    NO comparar contra Compra.Saldo porque puede estar en 0
    --    La validación real del saldo se hace en paso 5 vs amortización
    -- --------------------------------------------------------
    IF @ValorDescuento <= 0
    BEGIN
        SET @Mensaje = 'VALOR INVALIDO: El valor con descuento debe ser mayor a cero.'
        RETURN
    END

    -- --------------------------------------------------------
    -- 4. Verificar que no haya sido procesado ya
    -- --------------------------------------------------------
    IF EXISTS (SELECT 1 FROM CierreOperacionDescuentoLog
               WHERE idCompra = @idCompra AND Estado = 'PROCESADO')
    BEGIN
        SET @Mensaje = 'YA PROCESADO: Factura ' + @Factura + ' ya tiene un cierre registrado.'
        RETURN
    END

    -- --------------------------------------------------------
    -- 5. Validar cuotas pendientes y saldo real en amortización
    -- --------------------------------------------------------
    DECLARE @CntCuotas   INT,
            @SaldoReal   DECIMAL(18,4),
            @AntPendient INT

    SELECT @CntCuotas = COUNT(1),
           @SaldoReal  = ISNULL(SUM(Capital - ABCapital), 0)
    FROM   Cre_TablaDeAmortizacion
    WHERE  Vendida  = 0
      AND  Estado   NOT IN (2, 4)
      AND  idCompra = @idCompra

    IF ISNULL(@CntCuotas, 0) = 0
    BEGIN
        SET @Mensaje = 'SIN CUOTAS PENDIENTES en la tabla de amortización para factura ' + @Factura
        RETURN
    END

    IF ISNULL(@SaldoReal, 0) <= 0
    BEGIN
        SET @Mensaje = 'SALDO REAL DE CUOTAS es cero — la operación ya está cerrada en amortización: ' + @Factura
        RETURN
    END

    -- Siempre usar el saldo real de cuotas, igual que el proceso original
    SET @ValorDescuento = @SaldoReal

    -- Anticipos pendientes sin aplicar
    SELECT @AntPendient = COUNT(1) FROM Anticipos WHERE idOrigen = @idCompra AND Estado = 1
    IF ISNULL(@AntPendient, 0) > 0
    BEGIN
        SET @Mensaje = 'HAY ' + CAST(@AntPendient AS VARCHAR(5)) + ' ANTICIPO(S) PENDIENTE(S) sin aplicar para factura ' + @Factura + ' — liquidar primero'
        RETURN
    END

    -- --------------------------------------------------------
    -- 6. Obtener @Tipo desde Cre_EntidadFinanciera
    --    Si Confirmado=0 → Tipo=2 (igual que proceso original)
    --    Excepción: POINT (EF=1) con idMotivo=39 → Tipo=17
    -- --------------------------------------------------------
    SELECT @Tipo = CASE WHEN @Confirmado = 0 THEN 2 ELSE ISNULL(Tipo, 2) END
    FROM   Cre_EntidadFinanciera
    WHERE  idEntidadFinanciera = @idEntidadFinanciera

    IF @Tipo IS NULL SET @Tipo = 2

    IF @idEntidadFinanciera = 1 AND ISNULL(@idMotivo, 0) = 39
        SET @Tipo = 17

    -- --------------------------------------------------------
    -- 6. Obtener siguiente número de anticipo (Bodega 57)
    --    Igual que proceso original
    -- --------------------------------------------------------
    SELECT @Numero = NumAnticipo FROM Bodega WHERE Bodega = @Bodega
    SET @Zero   = SUBSTRING('000000000', 1, 7 - LEN(@Numero))
    SET @Numero = @Zero + CAST(@Numero AS VARCHAR(20))
    UPDATE Bodega SET NumAnticipo = NumAnticipo + 1 WHERE Bodega = @Bodega

    SET @Detalle = 'CIERRE DESC. ESPECIAL - ' + UPPER(ISNULL(@Segmento, ''))
        + ' - FACT: ' + @Factura
        + ' - AUT: ' + @NumAutorizacion
        + ' / RESP: ' + @UsuarioAutoriza

    BEGIN TRANSACTION

    BEGIN TRY

        -- --------------------------------------------------------
        -- 7. Crear Anticipo por el valor con descuento
        --    Igual que proceso original pero con @ValorDescuento
        -- --------------------------------------------------------
        INSERT INTO Anticipos (
            ClienteProveedor, Fecha, Bodega, idCliente, Secuencial,
            idFormaPago, Detalle, Abono, Estado, idOrigen,
            Tipo, Pichincha, NumeroDeOperacion, NormalAtrasado, EstadoAbogados
        )
        VALUES (
            0, @Fecha, @Bodega, @idCliente, @Numero,
            27, @Detalle, @ValorDescuento, 3, @idCompra,
            @Tipo, CASE @Confirmado WHEN 1 THEN 1 ELSE 0 END,
            @NumeroDeOperacion, 0, 0
        )

        SET @idAnticipo = SCOPE_IDENTITY()

        INSERT INTO DetalleFormaPago (idCompra, idFormaPago, Origen, Valor)
        VALUES (@idAnticipo, 27, 3, @ValorDescuento)

        -- --------------------------------------------------------
        -- 8. Procesar cuotas con @ValorDescuento
        --    Igual que proceso original — distribuye el valor
        --    entre cuotas hasta agotarlo
        -- --------------------------------------------------------
        SET @Saldo  = @ValorDescuento
        SET @Estado = 2

        DECLARE cTabAmort CURSOR FOR
            SELECT idCre_TablaDeAmortizacion,
                   Capital - ABCapital,
                   Interes - ABInteres
            FROM   Cre_TablaDeAmortizacion
            WHERE  Vendida  = 0
              AND  Estado   NOT IN (2, 4)
              AND  idCompra = @idCompra
            ORDER BY NumeroCuota

        OPEN cTabAmort
        FETCH NEXT FROM cTabAmort INTO @idCre_TablaDeAmort, @Capital, @Interes

        WHILE (@@FETCH_STATUS = 0)
        BEGIN
            IF (@Saldo > 0)
            BEGIN
                IF (ROUND(@Saldo, 2) >= ROUND(@Capital, 2))
                BEGIN
                    UPDATE Cre_TablaDeAmortizacion SET Interes = 0
                    WHERE  idCre_TablaDeAmortizacion = @idCre_TablaDeAmort
                    SET @Saldo   = @Saldo - @Capital
                    SET @Interes = 0
                    SET @Estado  = 2
                END
                ELSE
                BEGIN
                    UPDATE Cre_TablaDeAmortizacion SET Interes = 0
                    WHERE  idCre_TablaDeAmortizacion = @idCre_TablaDeAmort
                    SET @Interes = 0
                    SET @Capital = @Saldo
                    SET @Estado  = 3
                    SET @Saldo   = 0
                END

                EXEC ActualizaCobranza  @idCre_TablaDeAmort, 0, 0, 0, @Capital, @Capital, 0, 0, @Fecha, @idAnticipo, 0, @Capital, 0, 0
                EXEC GuardaPagosCuotasCredito @idAnticipo, @idCre_TablaDeAmort, 0, 0, @Interes, @Capital, @Capital, 0, @Estado, 'CANCELACION CUOTA', @Fecha, 1, @Capital, 1, '', 0
            END

            FETCH NEXT FROM cTabAmort INTO @idCre_TablaDeAmort, @Capital, @Interes
        END

        CLOSE cTabAmort
        DEALLOCATE cTabAmort

        -- --------------------------------------------------------
        -- 9. Condonar cuotas restantes (parte del descuento)
        --    Las que quedaron pendientes se cierran con Capital=0
        -- --------------------------------------------------------
        DECLARE @DetalleCondona VARCHAR(200)
        SET @DetalleCondona = 'DESCUENTO ESPECIAL AUT: ' + @NumAutorizacion

        DECLARE cTabAmortDesc CURSOR FOR
            SELECT idCre_TablaDeAmortizacion,
                   Capital - ABCapital,
                   Interes - ABInteres
            FROM   Cre_TablaDeAmortizacion
            WHERE  Vendida  = 0
              AND  Estado   NOT IN (2, 4)
              AND  idCompra = @idCompra
            ORDER BY NumeroCuota

        OPEN cTabAmortDesc
        FETCH NEXT FROM cTabAmortDesc INTO @idCre_TablaDeAmort, @Capital, @Interes

        WHILE (@@FETCH_STATUS = 0)
        BEGIN
            UPDATE Cre_TablaDeAmortizacion SET Interes = 0
            WHERE  idCre_TablaDeAmortizacion = @idCre_TablaDeAmort

            EXEC ActualizaCobranza  @idCre_TablaDeAmort, 0, 0, 0, @Capital, @Capital, 0, 0, @Fecha, @idAnticipo, 0, @Capital, 0, 0
            EXEC GuardaPagosCuotasCredito @idAnticipo, @idCre_TablaDeAmort, 0, 0, 0, @Capital, @Capital, 0, 2, @DetalleCondona, @Fecha, 1, @Capital, 1, '', 0

            FETCH NEXT FROM cTabAmortDesc INTO @idCre_TablaDeAmort, @Capital, @Interes
        END

        CLOSE cTabAmortDesc
        DEALLOCATE cTabAmortDesc

        -- --------------------------------------------------------
        -- 10. Generar asiento contable
        --     Mismo SP que proceso original
        -- --------------------------------------------------------
        EXEC AsientoCobranzaCrediPoint @idAnticipo, 0, @idCompra

        -- --------------------------------------------------------
        -- 11. Registrar en log de auditoría
        -- --------------------------------------------------------
        INSERT INTO CierreOperacionDescuentoLog (
            idCompra, Cedula, Cliente, Factura, Segmento,
            SaldoOriginal, ValorDescuento, MontoAjuste,
            NumAutorizacion, UsuarioAutoriza, UsuarioEjecuta,
            idAsiento, Estado
        )
        VALUES (
            @idCompra, @Cedula, @Cliente, @Factura, @Segmento,
            @SaldoReal, @ValorDescuento, @SaldoReal - @ValorDescuento,
            @NumAutorizacion, @UsuarioAutoriza, @UsuarioEjecuta,
            @idAnticipo, 'PROCESADO'
        )

        COMMIT TRANSACTION

    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION
        SET @Mensaje = 'ERROR SISTEMA: ' + ERROR_MESSAGE()
    END CATCH
END
GO

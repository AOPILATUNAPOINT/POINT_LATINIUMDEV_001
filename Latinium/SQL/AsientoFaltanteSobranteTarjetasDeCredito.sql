ALTER PROCEDURE [dbo].[AsientoFaltanteSobranteTarjetasDeCredito]
    @idDepositoTarjeta Int, @Fecha Date, @Numero Varchar(50), @Bodega Int, @idCuenta Int
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @CodAsiento Varchar(20), @RetIdAsiento Int, @idProyecto Int,
            @idCuentaSobrante Int, @idCuentaTC Int, @idCuentaFaltante Int,
            @NumeroOrigen Varchar(20), @Concepto_Origen Varchar(60), @IdAsiento Int

    SET @idCuenta = 1239

    SELECT @idProyecto = idProyecto FROM Bodega WHERE Bodega = @Bodega;

    SELECT @NumeroOrigen = Numero, @Concepto_Origen = Concepto, @IdAsiento = idAsientoDiferencias
      FROM DepositosDeTarjetas WHERE idDepositoTarjeta = @idDepositoTarjeta

    IF EXISTS (SELECT 1 FROM Asiento WHERE idAsiento = @IdAsiento AND Borrar = 0)
    BEGIN
        UPDATE Asiento
           SET Fecha = @Fecha, Descripcion = 'DIFERENCIAS LIQUIDACION DE TARJETAS ' + @Concepto_Origen,
               Numero = @Numero, idOrigenAsiento = 31, idProyecto = @idProyecto
         WHERE idAsiento = @IdAsiento

        DELETE FROM AsientoDetalle WHERE idAsiento = @IdAsiento
        SET @RetIdAsiento = @IdAsiento
    END
    ELSE
    BEGIN
        EXEC ASIENTONUMERAR 2, @CodAsiento OUTPUT

        INSERT INTO Asiento (idTipoAsiento, Fecha, Descripcion, CodAsiento, Numero, idOrigenAsiento, idProyecto)
        VALUES (1, @Fecha, 'DIFERENCIAS LIQUIDACION DE TARJETAS ' + @Concepto_Origen, @CodAsiento, @Numero, 31, @idProyecto)

        SET @RetIdAsiento = SCOPE_IDENTITY();
        UPDATE DepositosDeTarjetas SET idAsientoDiferencias = @RetIdAsiento
         WHERE idDepositoTarjeta = @idDepositoTarjeta
    END

    SELECT @idCuentaSobrante  = idCuenta FROM AsignaCuenta WHERE idAsignacion = 108
    SELECT @idCuentaTC        = idCuenta FROM AsignaCuenta WHERE idAsignacion = 92
    SELECT @idCuentaFaltante  = idCuenta FROM AsignaCuenta WHERE idAsignacion = 107

    -- SOBRANTE (solo TC/Débito)
    INSERT INTO AsientoDetalle (idAsiento, idCuenta, Detalle, Debe, idProyecto)
    SELECT @RetIdAsiento, @idCuenta,
           'SOBRANTE TARJETAS DE CREDITO LOTE ' + dfp.Lote,
           SUM(dfp.Diferencia), @idProyecto
    FROM DetalleDepositoDeTarjetas ddt
    INNER JOIN DetalleFormaPago dfp ON dfp.idDetFormaPago = ddt.idDetFormaPago
    WHERE ddt.idDepositoTarjeta = @idDepositoTarjeta
      AND dfp.Diferencia > 0
      AND dfp.idFormaPago IN (3, 11)
    GROUP BY dfp.Lote;

    INSERT INTO AsientoDetalle (idAsiento, idCuenta, Detalle, Haber, idProyecto)
    SELECT @RetIdAsiento, @idCuentaSobrante,
           'SOBRANTE TARJETAS DE CREDITO LOTE ' + dfp.Lote,
           SUM(dfp.Diferencia), @idProyecto
    FROM DetalleDepositoDeTarjetas ddt
    INNER JOIN DetalleFormaPago dfp ON dfp.idDetFormaPago = ddt.idDetFormaPago
    WHERE ddt.idDepositoTarjeta = @idDepositoTarjeta
      AND dfp.Diferencia > 0
      AND dfp.idFormaPago IN (3, 11)
    GROUP BY dfp.Lote;

    -- FALTANTE (solo TC/Débito)
    INSERT INTO AsientoDetalle (idAsiento, idCuenta, Detalle, Debe, idProyecto)
    SELECT @RetIdAsiento, @idCuentaFaltante,
           'FALTANTE TARJETAS DE CREDITO LOTE ' + dfp.Lote,
           ABS(SUM(dfp.Diferencia)), @idProyecto
    FROM DetalleDepositoDeTarjetas ddt
    INNER JOIN DetalleFormaPago dfp ON dfp.idDetFormaPago = ddt.idDetFormaPago
    WHERE ddt.idDepositoTarjeta = @idDepositoTarjeta
      AND dfp.Diferencia < 0
      AND dfp.idFormaPago IN (3, 11)
    GROUP BY dfp.Lote;

    INSERT INTO AsientoDetalle (idAsiento, idCuenta, Detalle, Haber, idProyecto)
    SELECT @RetIdAsiento, @idCuentaTC,
           'FALTANTE TARJETAS DE CREDITO LOTE ' + dfp.Lote,
           ABS(SUM(dfp.Diferencia)), @idProyecto
    FROM DetalleDepositoDeTarjetas ddt
    INNER JOIN DetalleFormaPago dfp ON dfp.idDetFormaPago = ddt.idDetFormaPago
    WHERE ddt.idDepositoTarjeta = @idDepositoTarjeta
      AND dfp.Diferencia < 0
      AND dfp.idFormaPago IN (3, 11)
    GROUP BY dfp.Lote;

    EXEC AsientoTotal @RetIdAsiento;
END

CREATE PROCEDURE [dbo].[ValidaCambioFormaPago]
    @idDetFormaPago    Int,
    @idFormaPagoActual Int,
    @idFormaPagoNuevo  Int = 0,
    @idOrigen          Int = 0
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @Mensaje Varchar(500) = ''

    -- Nunca se pueden tener como forma de pago actual modificable
    IF @idFormaPagoActual IN (6, 8, 29, 30, 39)
    BEGIN
        SET @Mensaje = 'No se puede modificar esta forma de pago.'
        SELECT @Mensaje
        RETURN
    END

    -- Nunca se puede cambiar A estas formas de pago
    IF @idFormaPagoNuevo IN (6, 8, 10, 12, 29, 30, 37, 39)
    BEGIN
        SET @Mensaje = 'No se puede seleccionar esta forma de pago.'
        SELECT @Mensaje
        RETURN
    END

    -- Anticipo (idOrigen=2) con TC: no puede cambiar a forma no-TC
    IF @idOrigen = 2 AND @idFormaPagoActual IN (3, 11, 45)
        AND @idFormaPagoNuevo NOT IN (3, 11, 45) AND @idFormaPagoNuevo > 0
    BEGIN
        SET @Mensaje = 'El anticipo fue registrado con Tarjeta. No se puede cambiar a otra forma de pago.'
        SELECT @Mensaje
        RETURN
    END

    -- Solo aplica si el registro ya existe en BD
    IF @idDetFormaPago > 0
    BEGIN
        -- TC/Debito/De Una: bloquear si esta en liquidacion
        IF @idFormaPagoActual IN (3, 11, 45)
        BEGIN
            IF EXISTS (SELECT 1 FROM DetalleDepositoDeTarjetas WHERE idDetFormaPago = @idDetFormaPago)
            BEGIN
                DECLARE @NumDep Varchar(50), @FechaDep Varchar(20)
                SELECT TOP 1
                    @NumDep   = dt.Numero,
                    @FechaDep = CONVERT(Varchar(20), dt.Fecha, 103)
                FROM DetalleDepositoDeTarjetas ddt
                INNER JOIN DepositosDeTarjetas dt ON dt.idDepositoTarjeta = ddt.idDepositoTarjeta
                WHERE ddt.idDetFormaPago = @idDetFormaPago

                SET @Mensaje = 'No se puede modificar. El registro esta en la liquidacion N: ' +
                               ISNULL(@NumDep, '') + ' del ' + ISNULL(@FechaDep, '')
            END
        END

        -- Transferencia/Deposito: bloquear si esta conciliado
        IF @idFormaPagoActual IN (7, 13) AND @Mensaje = ''
        BEGIN
            IF EXISTS (
                SELECT 1
                FROM AsientoDetalle ad
                INNER JOIN DetalleConciliacionBancaria dcb ON dcb.idDetAsiento = ad.idDetAsiento
                INNER JOIN ConciliacionBancaria cb ON cb.idConciliacionBancaria = dcb.idConciliacionBancaria
                WHERE ad.idAsiento = (SELECT idAsiento FROM DetalleFormaPago WHERE idDetFormaPago = @idDetFormaPago)
                  AND cb.Estado < 2
            )
                SET @Mensaje = 'No se puede modificar. El registro ya esta conciliado.'
        END
    END

    SELECT @Mensaje
END

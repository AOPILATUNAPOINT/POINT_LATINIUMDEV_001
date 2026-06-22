-- Tabla de configuracion de limites de campos en DetalleFormaPago
-- Permite cambiar MaxLength de los campos sin recompilar la aplicacion

IF NOT EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_NAME = 'DetalleFormaPagoConfiguracion'
)
BEGIN
    CREATE TABLE DetalleFormaPagoConfiguracion (
        idDetalleFormaPagoConfiguracion INT           IDENTITY(1,1) PRIMARY KEY,
        Clave                           VARCHAR(100)  NOT NULL,
        Descripcion                     VARCHAR(500)  NULL,
        Valor                           INT           NOT NULL,
        Activo                          BIT           NOT NULL DEFAULT 1,
        FechaCreacion                   DATETIME      NOT NULL DEFAULT GETDATE()
    )
END
GO

-- MaxLength campo "N. Transaccion" (txtNumVoucher) - antes: 4
IF NOT EXISTS (SELECT 1 FROM DetalleFormaPagoConfiguracion WHERE Clave = 'MaxLengthNumVoucher')
    INSERT INTO DetalleFormaPagoConfiguracion (Clave, Descripcion, Valor)
    VALUES ('MaxLengthNumVoucher', 'Longitud maxima campo N. Transaccion (voucher)', 20)
GO

-- MaxLength campo "Lote" (txtLote) - antes: 4
IF NOT EXISTS (SELECT 1 FROM DetalleFormaPagoConfiguracion WHERE Clave = 'MaxLengthLote')
    INSERT INTO DetalleFormaPagoConfiguracion (Clave, Descripcion, Valor)
    VALUES ('MaxLengthLote', 'Longitud maxima campo Lote de Voucher', 4)
GO

-- MaxLength campo "Numero De TRX" (txtnumeroTRX) - antes: 10
IF NOT EXISTS (SELECT 1 FROM DetalleFormaPagoConfiguracion WHERE Clave = 'MaxLengthNumeroTRX')
    INSERT INTO DetalleFormaPagoConfiguracion (Clave, Descripcion, Valor)
    VALUES ('MaxLengthNumeroTRX', 'Longitud maxima campo Numero de TRX', 10)
GO

-- Formas de pago que validan numero duplicado en uGridDepositos (cierre de caja)
-- Cada registro = un idFormaPago que entra a la validacion
-- Definicion de duplicado: misma forma de pago + misma cuenta + mismo numero en filas nuevas del grid
-- Para desactivar una forma: UPDATE DetalleFormaPagoConfiguracion SET Activo=0 WHERE Clave='ValidaDepDuplicadoFP' AND Valor=<idFormaPago>
-- Para agregar una forma:   INSERT INTO DetalleFormaPagoConfiguracion (Clave, Descripcion, Valor) VALUES ('ValidaDepDuplicadoFP', 'descripcion', <idFormaPago>)
-- Para ver config actual:   SELECT dc.Descripcion, dc.Valor, fp.FormaPago, dc.Activo FROM DetalleFormaPagoConfiguracion dc INNER JOIN CompraFormaPago fp ON fp.idFormaPago = dc.Valor WHERE dc.Clave = 'ValidaDepDuplicadoFP'

-- EFECTIVO (1)
IF NOT EXISTS (SELECT 1 FROM DetalleFormaPagoConfiguracion WHERE Clave = 'ValidaDepDuplicadoFP' AND Valor = 1)
    INSERT INTO DetalleFormaPagoConfiguracion (Clave, Descripcion, Valor)
    VALUES ('ValidaDepDuplicadoFP', 'EFECTIVO - valida numero duplicado en cierre de caja', 1)
GO

-- SOBRANTE EFECTIVO (32) - desactivado: los numeros son fechas y pueden repetirse entre cajeros
IF NOT EXISTS (SELECT 1 FROM DetalleFormaPagoConfiguracion WHERE Clave = 'ValidaDepDuplicadoFP' AND Valor = 32)
    INSERT INTO DetalleFormaPagoConfiguracion (Clave, Descripcion, Valor, Activo)
    VALUES ('ValidaDepDuplicadoFP', 'SOBRANTE EFECTIVO - valida numero duplicado en cierre de caja', 32, 0)
GO

-- DE UNA (45)
IF NOT EXISTS (SELECT 1 FROM DetalleFormaPagoConfiguracion WHERE Clave = 'ValidaDepDuplicadoFP' AND Valor = 45)
    INSERT INTO DetalleFormaPagoConfiguracion (Clave, Descripcion, Valor)
    VALUES ('ValidaDepDuplicadoFP', 'DE UNA - valida numero duplicado en cierre de caja', 45)
GO

-- TRANSFERENCIA (7) - por si se habilita en el futuro
IF NOT EXISTS (SELECT 1 FROM DetalleFormaPagoConfiguracion WHERE Clave = 'ValidaDepDuplicadoFP' AND Valor = 7)
    INSERT INTO DetalleFormaPagoConfiguracion (Clave, Descripcion, Valor)
    VALUES ('ValidaDepDuplicadoFP', 'TRANSFERENCIA - valida numero duplicado en cierre de caja', 7)
GO

-- DEPOSITO (13) - por si se habilita en el futuro
IF NOT EXISTS (SELECT 1 FROM DetalleFormaPagoConfiguracion WHERE Clave = 'ValidaDepDuplicadoFP' AND Valor = 13)
    INSERT INTO DetalleFormaPagoConfiguracion (Clave, Descripcion, Valor)
    VALUES ('ValidaDepDuplicadoFP', 'DEPOSITO - valida numero duplicado en cierre de caja', 13)
GO

-- Para consultar la configuracion actual:
-- SELECT * FROM DetalleFormaPagoConfiguracion

-- Para cambiar un limite sin recompilar:
-- UPDATE DetalleFormaPagoConfiguracion SET Valor = 30 WHERE Clave = 'MaxLengthNumVoucher'

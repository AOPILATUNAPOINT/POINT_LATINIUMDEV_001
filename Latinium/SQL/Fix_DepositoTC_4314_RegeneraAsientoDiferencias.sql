-- CREQ-4314: Regenera asiento de diferencias del deposito 4314
-- Problema: registro EFECTIVO (idFormaPago=1) entro a DetalleDepositoDeTarjetas
-- y el SP genero un FALTANTE falso de $10.40 en cuenta 1.01.04.02
-- Solucion: SP actualizado con filtro IN (3,11) y asiento regenerado

-- Verificar asiento antes de ejecutar:
-- SELECT ad.idDetAsiento, ad.Detalle, ad.Debe, ad.Haber, c.Codigo, c.Descripcion
-- FROM AsientoDetalle ad
-- INNER JOIN Cuenta c ON c.idCuenta = ad.idCuenta
-- WHERE ad.idAsiento = 466227

EXEC dbo.AsientoFaltanteSobranteTarjetasDeCredito
    4314,        -- idDepositoTarjeta
    '20260508',  -- Fecha
    '261095566', -- Numero
    6,           -- Bodega (TANTRA PLATINUM)
    1239         -- idCuenta

-- Verificar resultado: debe quedar solo SOBRANTE $3.00, sin FALTANTE $10.40
-- SELECT ad.idDetAsiento, ad.Detalle, ad.Debe, ad.Haber, c.Codigo, c.Descripcion
-- FROM AsientoDetalle ad
-- INNER JOIN Cuenta c ON c.idCuenta = ad.idCuenta
-- WHERE ad.idAsiento = 466227

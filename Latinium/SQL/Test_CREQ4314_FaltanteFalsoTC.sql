-- ============================================================
-- CREQ-4314: Script de prueba en BD de desarrollo
-- Recrea el escenario: anticipo TC cambiado a EFECTIVO
-- genera faltante falso en liquidacion de tarjetas
-- ============================================================
-- PREREQUISITO: Ejecutar en BD de PRUEBAS, no produccion
-- PREREQUISITO: Tener el SP ValidaCambioFormaPago ya creado
-- ============================================================

-- ============================================================
-- PARTE 1: DATOS BASE (ajusta los IDs a tu BD de pruebas)
-- ============================================================

-- Busca una bodega valida en tu BD de pruebas
SELECT TOP 5 Bodega, Nombre, idProyecto FROM Bodega WHERE Activo = 1

-- Busca un cliente valido
-- SELECT TOP 5 idCliente, RazonSocial FROM Cliente WHERE Activo = 1

-- ============================================================
-- PARTE 2: SIMULAR ESTADO CORRUPTO (como quedó en produccion)
-- Insertar un DetalleFormaPago con idFormaPago=1 (EFECTIVO)
-- que tenga Diferencia negativa y Lote vacio,
-- y vincularlo a un deposito TC
-- ============================================================

-- PASO 2A: Crear el registro de pago corrupto (EFECTIVO, sin Lote)
-- (En produccion esto ocurrio porque el usuario cambio TC→EFECTIVO en VentaFPagoLot)
DECLARE @idDetFormaPagoTest INT
DECLARE @idDepositoTest     INT
DECLARE @idAsientoTest      INT

-- Inserta un DetalleFormaPago falso con EFECTIVO y diferencia negativa
INSERT INTO DetalleFormaPago (
    idFormaPago, Valor, ValorCobrar, Diferencia, Lote, Origen,
    Capturado, Usuario, FechaSistema
)
VALUES (
    1,        -- EFECTIVO (el bug: debería ser 3=TC o 11=Débito)
    10.40,    -- Valor
    10.40,    -- ValorCobrar
    -10.40,   -- Diferencia negativa = FALTANTE falso
    '',       -- Lote vacío (TC tendría número de lote)
    0,        -- Origen (int)
    0,
    'TEST',
    GETDATE()
)
SET @idDetFormaPagoTest = SCOPE_IDENTITY()
PRINT 'idDetFormaPago creado: ' + CAST(@idDetFormaPagoTest AS VARCHAR)

-- PASO 2B: Crear un deposito de tarjetas de prueba
INSERT INTO DepositosDeTarjetas (
    Numero, Fecha, Concepto, Estado, Usuario, FechaSistema
)
VALUES (
    'TEST-9999', GETDATE(), 'PRUEBA CREQ4314', 1, 'TEST', GETDATE()
)
SET @idDepositoTest = SCOPE_IDENTITY()
PRINT 'idDepositoTarjeta creado: ' + CAST(@idDepositoTest AS VARCHAR)

-- PASO 2C: Vincular el pago EFECTIVO al deposito TC (como ocurrio en produccion)
INSERT INTO DetalleDepositoDeTarjetas (idDepositoTarjeta, idDetFormaPago)
VALUES (@idDepositoTest, @idDetFormaPagoTest)
PRINT 'Registro vinculado al deposito'

-- ============================================================
-- PARTE 3: VERIFICAR ESTADO INICIAL (debe mostrar el problema)
-- ============================================================
PRINT ''
PRINT '=== ESTADO INICIAL (muestra el registro no-TC en deposito) ==='
SELECT
    ddt.idDetDepositoTarjeta,
    dfp.idDetFormaPago,
    dfp.idFormaPago,
    fp.FormaPago,
    dfp.Valor,
    dfp.Diferencia,
    dfp.Lote,
    dfp.Origen
FROM DetalleDepositoDeTarjetas ddt
INNER JOIN DetalleFormaPago dfp ON dfp.idDetFormaPago = ddt.idDetFormaPago
LEFT JOIN  FormaPago fp           ON fp.idFormaPago   = dfp.idFormaPago
WHERE ddt.idDepositoTarjeta = @idDepositoTest

-- ============================================================
-- PARTE 4: EJECUTAR SP (ANTES del fix — debe generar FALTANTE falso)
-- Comenta esta parte si el SP ya tiene el fix aplicado
-- ============================================================
PRINT ''
PRINT '=== EJECUTANDO SP AsientoFaltanteSobranteTarjetasDeCredito ==='
EXEC dbo.AsientoFaltanteSobranteTarjetasDeCredito
    @idDepositoTest, GETDATE(), 'TEST-9999', 6, 1239

-- Ver el asiento generado
SELECT @idAsientoTest = idAsientoDiferencias
FROM DepositosDeTarjetas WHERE idDepositoTarjeta = @idDepositoTest

PRINT 'idAsiento generado: ' + CAST(@idAsientoTest AS VARCHAR)

PRINT ''
PRINT '=== ASIENTO GENERADO ==='
SELECT
    ad.idDetAsiento,
    ad.Detalle,
    ad.Debe,
    ad.Haber,
    c.Codigo,
    c.Descripcion AS Cuenta
FROM AsientoDetalle ad
INNER JOIN Cuenta c ON c.idCuenta = ad.idCuenta
WHERE ad.idAsiento = @idAsientoTest
ORDER BY ad.idDetAsiento

-- CON EL BUG:    aparece FALTANTE con idCuenta de anticipo empleados
-- CON EL FIX:    no aparece nada (registro EFECTIVO ignorado) o solo SOBRANTE TC real

-- ============================================================
-- PARTE 5: PROBAR ValidaCambioFormaPago (prevención)
-- ============================================================
PRINT ''
PRINT '=== PRUEBA ValidaCambioFormaPago ==='

-- CASO A: Anticipo nuevo (idDetFormaPago=0), TC → EFECTIVO con idOrigen=2
-- ESPERADO: "El anticipo fue registrado con Tarjeta. No se puede cambiar..."
PRINT 'CASO A - Anticipo nuevo TC→EFECTIVO (debe bloquear):'
EXEC ValidaCambioFormaPago 0, 3, 1, 2   -- idDetFP=0, actual=TC(3), nuevo=EFECTIVO(1), origen=anticipo(2)

-- CASO B: Registro TC ya en un deposito → cambio de forma
-- ESPERADO: "No se puede modificar. El registro esta en la liquidacion N: TEST-9999..."
PRINT 'CASO B - TC en liquidacion (debe bloquear):'
EXEC ValidaCambioFormaPago @idDetFormaPagoTest, 3, 1, 0

-- CASO C: Cambio a forma tipo N (idFormaPago=10)
-- ESPERADO: "No se puede seleccionar esta forma de pago."
PRINT 'CASO C - Cambio a forma N (debe bloquear):'
EXEC ValidaCambioFormaPago 0, 3, 10, 0

-- CASO D: Cambio valido (TC → Débito, sin liquidacion)
-- ESPERADO: '' (mensaje vacío = permitido)
PRINT 'CASO D - Cambio TC→Debito sin liquidacion (debe permitir):'
EXEC ValidaCambioFormaPago 0, 3, 11, 0

-- ============================================================
-- PARTE 6: LIMPIEZA (ejecutar al terminar las pruebas)
-- ============================================================
-- DESCOMENTA cuando termines de revisar los resultados:
/*
DELETE FROM AsientoDetalle WHERE idAsiento = @idAsientoTest
DELETE FROM Asiento         WHERE idAsiento = @idAsientoTest
DELETE FROM DetalleDepositoDeTarjetas WHERE idDepositoTarjeta = @idDepositoTest
DELETE FROM DepositosDeTarjetas       WHERE idDepositoTarjeta = @idDepositoTest
DELETE FROM DetalleFormaPago          WHERE idDetFormaPago     = @idDetFormaPagoTest
PRINT 'Datos de prueba eliminados'
*/

-- ============================================================
-- FIX: OrdenDeCompraRevertirFacturacionODC
--
-- PROBLEMA 1: Cuando hay ítems pendientes (CantidadFacturada < Cantidad),
--   el SP pone la ODC en Estado = 5 (APROBADO). Esto saca la ODC
--   del flujo de facturación y bloquea la creación de una segunda factura
--   para ítems que YA FUERON RECIBIDOS.
--
-- PROBLEMA 2 (nuevo): Si una ODC tiene recepciones parciales (ej. 50 de
--   100 cilindros), una vez que lo recibido está completamente facturado
--   la ODC debe volver a Estado = 5 (APROBADO) para poder recepcionar
--   el resto. Si se dejara en Estado = 20 (POR FACTURAR), el botón
--   "Recepción Bodega" de frmEntregaRecepcionMercaderia nunca se habilitaría.
--
-- SOLUCIÓN: Lógica de 3 vías basada en CantidadRecibida:
--
--   CASO 1 → Estado 9 (FACTURADO)
--     No existe ningún item con CantidadFacturada < Cantidad
--     → ODC completamente cerrada
--
--   CASO 2 → Estado 20 (POR FACTURAR)
--     Existe al menos un item con CantidadFacturada < CantidadRecibida
--     → Hay mercadería recibida en bodega que aún no se ha facturado
--
--   CASO 3 → Estado 5 (APROBADO)
--     Todo lo recibido está facturado (CantidadFacturada >= CantidadRecibida)
--     pero quedan items por recibir (CantidadRecibida < Cantidad)
--     → Debe regresar a APROBADO para permitir una nueva Recepción Bodega
-- ============================================================

ALTER PROCEDURE [dbo].[OrdenDeCompraRevertirFacturacionODC]
(
    @idCompra INT,
    @Estado   INT,
    @Usuario  VARCHAR(50)
)
AS
BEGIN

    SET NOCOUNT ON;

    PRINT '=============================================';
    PRINT 'INICIO REVERTIR POR FACTURAR';
    PRINT 'idCompra: ' + CAST(@idCompra AS VARCHAR);
    PRINT 'Estado: '   + CAST(@Estado   AS VARCHAR);
    PRINT 'Usuario: '  + ISNULL(@Usuario, '');
    PRINT '=============================================';

    IF (@Estado <> 20)
    BEGIN
        PRINT 'La compra no está en estado 20.';
        RETURN;
    END

    IF NOT EXISTS (SELECT 1 FROM Compra WHERE idCompra = @idCompra)
    BEGIN
        PRINT 'La compra no existe.';
        RETURN;
    END

    PRINT 'Compra válida.';

    IF NOT EXISTS
    (
        SELECT 1
        FROM DetCompra
        WHERE idCompra = @idCompra
          AND ISNULL(CantidadPorFacturar, 0) > 0
    )
    BEGIN
        PRINT 'No existen cantidades por facturar.';
        RETURN;
    END

    PRINT 'Existen cantidades reservadas.';

    --------------------------------------------------------------------
    -- 1. CONSOLIDAR FACTURADO
    --------------------------------------------------------------------
    PRINT 'Consolidando cantidades...';

    UPDATE DetCompra
    SET
        CantidadFacturada   = ISNULL(CantidadFacturada,   0)
                            + ISNULL(CantidadPorFacturar, 0),
        CantidadPorFacturar = 0,

        FechaFacturacion       = GETDATE(),
        UsuarioFacturacion     = @Usuario,

        FechaReservaFactura    = NULL,
        UsuarioReservaFactura  = NULL
    WHERE idCompra = @idCompra;

    PRINT 'DetCompra actualizado: ' + CAST(@@ROWCOUNT AS VARCHAR);

    --------------------------------------------------------------------
    -- 2. ESTADO FACTURACIÓN DETCOMPRA
    --------------------------------------------------------------------
    UPDATE DetCompra
    SET EstadoFacturacion =
        CASE
            WHEN ISNULL(CantidadFacturada, 0) >= ISNULL(Cantidad, 0) THEN 2
            WHEN ISNULL(CantidadFacturada, 0) >  0                   THEN 1
            ELSE 0
        END
    WHERE idCompra = @idCompra;

    --------------------------------------------------------------------
    -- 3. SINCRONIZAR DETALLE SERIALES: FAC_SEL (4) → FAC_FACT (2)
    --------------------------------------------------------------------
    PRINT 'Actualizando DetalleSeriales a FACTURADO...';

    UPDATE DS
    SET DS.EstadoFacturacion = 2
    FROM DetalleSeriales DS
    INNER JOIN DetCompra DC ON DC.idDetCompra = DS.idDetCompra
    WHERE DC.idCompra = @idCompra
      AND DS.EstadoFacturacion = 4;

    PRINT 'DetalleSeriales facturados: ' + CAST(@@ROWCOUNT AS VARCHAR);

    --------------------------------------------------------------------
    -- 4. LIMPIAR DETALLE SERIALES PENDIENTES (reset FAC_PEND = 0 → 0)
    --------------------------------------------------------------------
    UPDATE DS
    SET DS.EstadoFacturacion = 0
    FROM DetalleSeriales DS
    INNER JOIN DetCompra DC ON DC.idDetCompra = DS.idDetCompra
    WHERE DC.idCompra = @idCompra
      AND ISNULL(DS.EstadoFacturacion, 0) = 0;

    PRINT 'DetalleSeriales limpiados: ' + CAST(@@ROWCOUNT AS VARCHAR);

    --------------------------------------------------------------------
    -- 5. ESTADO FINAL COMPRA — LÓGICA DE 3 VÍAS
    --
    --   CASO 1 → Estado 9  : todo completamente facturado
    --   CASO 2 → Estado 20 : hay mercadería recibida sin facturar
    --   CASO 3 → Estado 5  : todo lo recibido está facturado,
    --                         pero quedan items por recibir
    --------------------------------------------------------------------
    PRINT 'Evaluando estado final...';

    IF NOT EXISTS
    (
        -- ¿Queda algo sin facturar del total pedido?
        SELECT 1
        FROM DetCompra
        WHERE idCompra = @idCompra
          AND ISNULL(CantidadFacturada, 0) < ISNULL(Cantidad, 0)
    )
    BEGIN
        -- CASO 1: ODC completamente facturada
        UPDATE Compra
        SET Estado       = 9,
            FechaEditado = GETDATE(),
            Usuario      = @Usuario
        WHERE idCompra = @idCompra;

        UPDATE DetCompra
        SET EstadoSerial = 1
        WHERE idCompra = @idCompra;

        PRINT 'RESULTADO: COMPLETO → ESTADO 9 (FACTURADO)';
    END
    ELSE IF EXISTS
    (
        -- ¿Hay items recibidos en bodega que aún no se facturaron?
        SELECT 1
        FROM DetCompra
        WHERE idCompra = @idCompra
          AND ISNULL(CantidadFacturada, 0) < ISNULL(CantidadRecibida, 0)
    )
    BEGIN
        -- CASO 2: Quedan items recibidos por facturar
        UPDATE Compra
        SET Estado       = 20,
            FechaEditado = GETDATE(),
            Usuario      = @Usuario
        WHERE idCompra = @idCompra;

        PRINT 'RESULTADO: RECIBIDOS PENDIENTES → ESTADO 20 (POR FACTURAR)';
    END
    ELSE
    BEGIN
        -- CASO 3: Todo lo recibido está facturado, pero faltan items
        --         por recibir. Regresa a APROBADO para permitir
        --         una nueva Recepción Bodega.
        UPDATE Compra
        SET Estado       = 5,
            FechaEditado = GETDATE(),
            Usuario      = @Usuario
        WHERE idCompra = @idCompra;

        PRINT 'RESULTADO: RECIBIDOS COMPLETOS, PENDIENTE RECEPCIÓN → ESTADO 5 (APROBADO)';
    END

    PRINT '=============================================';
    PRINT 'FIN PROCESO';
    PRINT '=============================================';

END

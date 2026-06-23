-- ============================================================
-- FIX: OrdenDeCompraEnviarRecibidoBodega
--
-- PROBLEMA: El SP tiene un bloque que lanza RAISERROR cuando
--   la ODC ya está en estado 28 (RECIBIDO EN BODEGA). Esto
--   bloquea el Escenario A:
--
--   Escenario A (múltiples recepciones antes de facturar):
--     Recibir 2 TVs → ODC estado 28
--     Llega 1 TV más → usuario quiere recibirla y facturar las 3 juntas
--     → BLOQUEADO: "La compra ya se encuentra en estado RECIBIDO EN BODEGA"
--
-- SOLUCIÓN: Eliminar el bloque de validación que impide re-recepcionar.
--   Cuando la ODC ya está en 28:
--     - La cabecera queda en 28 (no hay cambio)
--     - Los ítems con CantidadRecibida > 0 y EstadoFacturacion = 0
--       se actualizan a FAC_REC (3)
--     - Los seriales nuevos (Disponible=1, Estado != FAC_REC) se actualizan
--   El UPDATE de EstadoFacturacion solo afecta ítems que aún no tienen
--   estado de facturación (0=FAC_PEND), para no pisar ítems que ya
--   tienen facturación parcial (1=FAC_RES) o total (2=FAC_FACT).
--
-- ESCENARIOS SOPORTADOS TRAS ESTE FIX:
--   A. Recibir 2 → Recibir 1 más (antes de facturar) → Facturar 3 juntas ✓
--   B. Recibir 2 → Facturar 2 → Recibir 1 → Facturar 1 (flujo Revertir) ✓
--   C. Recibir 2 → Facturar 1 → Recibir 1 → Facturar 2 (combinado) ✓
-- ============================================================

ALTER PROCEDURE [dbo].[OrdenDeCompraEnviarRecibidoBodega]
(
    @idCompra INT,
    @Usuario  VARCHAR(50)
)
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        BEGIN TRAN;

        -------------------------------------------------------------------
        -- ESTADO COMPRA RECIBIDO
        -------------------------------------------------------------------
        DECLARE @EstadoCompraRecibido INT;

        SELECT TOP 1
            @EstadoCompraRecibido = ValorLegacy
        FROM DetCompraEstado
        WHERE Codigo     = 'COMP_REC'
          AND TipoEstado = 'COMPRA'
          AND Activo     = 1;

        -------------------------------------------------------------------
        -- ESTADO FACTURACION RECIBIDO
        -------------------------------------------------------------------
        DECLARE @EstadoFacturacionRecibido INT;

        SELECT TOP 1
            @EstadoFacturacionRecibido = ValorLegacy
        FROM DetCompraEstado
        WHERE Codigo     = 'FAC_REC'
          AND TipoEstado = 'FACTURACION'
          AND Activo     = 1;

        -------------------------------------------------------------------
        -- VALIDAR ESTADOS CATÁLOGO
        -------------------------------------------------------------------
        IF @EstadoCompraRecibido IS NULL
        BEGIN
            RAISERROR('No existe estado COMP_REC.', 16, 1);
            ROLLBACK TRAN;
            RETURN;
        END

        IF @EstadoFacturacionRecibido IS NULL
        BEGIN
            RAISERROR('No existe estado FAC_REC.', 16, 1);
            ROLLBACK TRAN;
            RETURN;
        END

        -------------------------------------------------------------------
        -- VALIDAR QUE EXISTA LA COMPRA
        -------------------------------------------------------------------
        IF NOT EXISTS (SELECT 1 FROM Compra WHERE idCompra = @idCompra)
        BEGIN
            RAISERROR('La compra no existe.', 16, 1);
            ROLLBACK TRAN;
            RETURN;
        END

        -------------------------------------------------------------------
        -- [FIX] BLOQUE ELIMINADO:
        --   "VALIDAR SI YA ESTÁ EN RECIBIDO"
        --   IF EXISTS (...Estado = @EstadoCompraRecibido) → RAISERROR
        --
        --   Se eliminó para permitir múltiples recepciones parciales
        --   antes de facturar (Escenario A). El SP es idempotente:
        --   si ya está en 28, actualizará los nuevos ítems/seriales
        --   y mantendrá el estado 28 sin problema.
        -------------------------------------------------------------------

        -------------------------------------------------------------------
        -- VALIDAR QUE EXISTAN CANTIDADES RECIBIDAS
        -------------------------------------------------------------------
        IF NOT EXISTS
        (
            SELECT 1
            FROM DetCompra
            WHERE idCompra = @idCompra
              AND ISNULL(CantidadRecibida, 0) > 0
        )
        BEGIN
            RAISERROR('No existen cantidades recibidas en bodega.', 16, 1);
            ROLLBACK TRAN;
            RETURN;
        END

        -------------------------------------------------------------------
        -- ACTUALIZA CABECERA
        -- (Si ya está en 28 el UPDATE es inocuo)
        -------------------------------------------------------------------
        UPDATE Compra
        SET
            Estado       = @EstadoCompraRecibido,
            FechaEditado = GETDATE(),
            Usuario      = @Usuario
        WHERE idCompra = @idCompra;

        -------------------------------------------------------------------
        -- ACTUALIZA DETALLE
        -- [FIX] Solo actualiza ítems que NO tienen aún facturación
        --   (EstadoFacturacion = 0 = FAC_PEND). No pisa ítems que ya
        --   tienen facturación parcial (1) o total (2).
        -------------------------------------------------------------------
        UPDATE DetCompra
        SET
            EstadoFacturacion = @EstadoFacturacionRecibido
        WHERE idCompra = @idCompra
          AND ISNULL(CantidadRecibida, 0) > 0
          AND ISNULL(EstadoFacturacion,  0) = 0;   -- solo FAC_PEND

        -------------------------------------------------------------------
        -- ACTUALIZA SERIALES
        -- (mismo comportamiento original: seriales nuevos con Disponible=1)
        -------------------------------------------------------------------
        UPDATE S
        SET S.Estado = @EstadoFacturacionRecibido
        FROM Seriales S
        WHERE EXISTS
        (
            SELECT 1
            FROM DetalleSeriales DS
            INNER JOIN DetCompra DC
                ON DC.idDetCompra = DS.idDetCompra
            WHERE DS.idSerial  = S.idSerial
              AND DC.idCompra  = @idCompra
              AND ISNULL(DS.Disponible, 0) = 1
        )
        AND ISNULL(S.Estado, 0) <> @EstadoFacturacionRecibido;

        -------------------------------------------------------------------
        -- COMMIT
        -------------------------------------------------------------------
        COMMIT TRAN;

        PRINT 'Compra enviada correctamente a RECIBIDO EN BODEGA';

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
            ROLLBACK TRAN;

        DECLARE @Error VARCHAR(MAX);
        SET @Error = ERROR_MESSAGE();
        RAISERROR(@Error, 16, 1);

    END CATCH

END

-- ============================================================
-- FIX: OrdenDeCompraEnviarPorFacturar
-- PROBLEMA: Cuando se crea la segunda factura para una ODC,
--   el SP lanza RAISERROR porque la ODC ya está en estado 20
--   (POR FACTURAR). Esto causa ROLLBACK y CantidadPorFacturar
--   vuelve a 0, mostrando cantidades en 0 en la segunda factura.
-- SOLUCIÓN: Eliminar el bloque "VALIDAR SI YA ESTÁ EN POR FACTURAR".
--   El SP ya tiene suficiente control: valida que existan
--   CantidadPorFacturar > 0 antes de proceder.
-- ============================================================

ALTER PROCEDURE [dbo].[OrdenDeCompraEnviarPorFacturar]
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
        -- ESTADOS DESDE CATÁLOGO
        -------------------------------------------------------------------
        DECLARE @EstadoCompraPorFacturar INT;

        SELECT TOP 1
            @EstadoCompraPorFacturar = ValorLegacy
        FROM DetCompraEstado
        WHERE Codigo    = 'COMP_POR_FACT'
          AND TipoEstado = 'COMPRA'
          AND Activo     = 1;

        DECLARE @EstadoFacturacionPorFacturar INT;
        DECLARE @EstadoFacturacionTotal       INT;

        IF @EstadoCompraPorFacturar IS NULL
        BEGIN
            RAISERROR('No existe estado POR FACTURAR en COMPRA.', 16, 1);
            ROLLBACK TRAN;
            RETURN;
        END

        -------------------------------------------------------------------
        -- FAC_RES
        -------------------------------------------------------------------
        SELECT TOP 1
            @EstadoFacturacionPorFacturar = ValorLegacy
        FROM DetCompraEstado
        WHERE Codigo    = 'FAC_RES'
          AND TipoEstado = 'FACTURACION'
          AND Activo     = 1;

        -------------------------------------------------------------------
        -- FAC_FACT
        -------------------------------------------------------------------
        SELECT TOP 1
            @EstadoFacturacionTotal = ValorLegacy
        FROM DetCompraEstado
        WHERE Codigo    = 'FAC_FACT'
          AND TipoEstado = 'FACTURACION'
          AND Activo     = 1;

        -------------------------------------------------------------------
        -- VALIDAR ESTADOS
        -------------------------------------------------------------------
        IF @EstadoFacturacionPorFacturar IS NULL
        BEGIN
            RAISERROR('No existe estado FAC_RES.', 16, 1);
            ROLLBACK TRAN;
            RETURN;
        END

        IF @EstadoFacturacionTotal IS NULL
        BEGIN
            RAISERROR('No existe estado FAC_FACT.', 16, 1);
            ROLLBACK TRAN;
            RETURN;
        END

        -------------------------------------------------------------------
        -- VALIDAR COMPRA
        -------------------------------------------------------------------
        IF NOT EXISTS (SELECT 1 FROM Compra WHERE idCompra = @idCompra)
        BEGIN
            RAISERROR('La compra no existe.', 16, 1);
            ROLLBACK TRAN;
            RETURN;
        END

        -- [BLOQUE ELIMINADO] "VALIDAR SI YA ESTÁ EN POR FACTURAR"
        -- Se eliminó el IF EXISTS que lanzaba RAISERROR cuando Estado = 20.
        -- Permite crear múltiples facturas para la misma ODC; la validación
        -- real es que existan CantidadPorFacturar > 0 (bloque siguiente).

        -------------------------------------------------------------------
        -- VALIDAR QUE EXISTAN SERIALES PREPARADOS
        -------------------------------------------------------------------
        IF NOT EXISTS
        (
            SELECT 1
            FROM DetCompra
            WHERE idCompra = @idCompra
              AND ISNULL(CantidadPorFacturar, 0) > 0
        )
        BEGIN
            RAISERROR('No existen seriales preparados para facturación.', 16, 1);
            ROLLBACK TRAN;
            RETURN;
        END

        -------------------------------------------------------------------
        -- ACTUALIZA CABECERA
        -------------------------------------------------------------------
        UPDATE Compra
        SET
            Estado       = @EstadoCompraPorFacturar,
            FechaEditado = GETDATE(),
            Usuario      = @Usuario
        WHERE idCompra = @idCompra;

        -------------------------------------------------------------------
        -- ACTUALIZA DETALLE  (NO modifica CantidadPorFacturar)
        -------------------------------------------------------------------
        UPDATE DetCompra
        SET
            EstadoFacturacion =
                CASE
                    WHEN ISNULL(CantidadFacturada, 0) >= Cantidad
                    THEN @EstadoFacturacionTotal
                    ELSE @EstadoFacturacionPorFacturar
                END
        WHERE idCompra = @idCompra
          AND ISNULL(CantidadPorFacturar, 0) > 0;

        -------------------------------------------------------------------
        -- ACTUALIZA SERIALES
        -------------------------------------------------------------------
        UPDATE S
        SET S.Estado = @EstadoFacturacionPorFacturar
        FROM Seriales S
        WHERE EXISTS
        (
            SELECT 1
            FROM DetalleSeriales DS
            INNER JOIN DetCompra DC ON DC.idDetCompra = DS.idDetCompra
            WHERE DS.idSerial  = S.idSerial
              AND DC.idCompra  = @idCompra
              AND ISNULL(DS.Disponible, 0) = 1
        )
          AND ISNULL(S.Estado, 0) <> @EstadoFacturacionTotal
          AND ISNULL(S.Estado, 0) <> @EstadoFacturacionPorFacturar;

        -------------------------------------------------------------------
        -- VALIDACIÓN FINAL
        -------------------------------------------------------------------
        IF NOT EXISTS
        (
            SELECT 1
            FROM DetCompra
            WHERE idCompra = @idCompra
              AND ISNULL(CantidadPorFacturar, 0) > 0
        )
        BEGIN
            RAISERROR('No existen cantidades preparadas para facturar.', 16, 1);
            ROLLBACK TRAN;
            RETURN;
        END

        -------------------------------------------------------------------
        -- COMMIT
        -------------------------------------------------------------------
        COMMIT TRAN;

        PRINT 'Compra enviada correctamente a POR FACTURAR';

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
            ROLLBACK TRAN;

        DECLARE @Error VARCHAR(MAX);
        SET @Error = ERROR_MESSAGE();
        RAISERROR(@Error, 16, 1);

    END CATCH

END

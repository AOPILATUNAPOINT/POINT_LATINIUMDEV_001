GO
/****** Object:  StoredProcedure [dbo].[OrdenDeCompraDetalleRecibidoCon]    Script Date: 23/6/2026 9:00:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER   PROCEDURE [dbo].[OrdenDeCompraDetalleRecibidoCon]
(
    @idCompra INT = 0
)
AS
BEGIN

    SET NOCOUNT ON;

    -----------------------------------------------------------------------
    -- ESTADO COMPRA RECIBIDO
    -----------------------------------------------------------------------
    DECLARE @EstadoCompraRecibido INT;
	DECLARE @EstadoCompraPorFactura INT;

    SELECT TOP 1
        @EstadoCompraRecibido = ValorLegacy
    FROM DetCompraEstado
    WHERE Codigo = 'COMP_REC'
      AND TipoEstado = 'COMPRA'
      AND Activo = 1;

	      SELECT TOP 1
        @EstadoCompraPorFactura = ValorLegacy
    FROM DetCompraEstado
    WHERE Codigo = 'COMP_POR_FACT'
      AND TipoEstado = 'COMPRA'
      AND Activo = 1;


    -----------------------------------------------------------------------
    -- VALIDAR ESTADO
    -----------------------------------------------------------------------
    IF @EstadoCompraRecibido IS NULL
    BEGIN
        RAISERROR(
            'No existe estado COMP_REC.',
            16,
            1
        );
        RETURN;
    END

	    IF @EstadoCompraRecibido IS NULL
    BEGIN
        RAISERROR(
            'No existe estado COMP_POR_FACT.',
            16,
            1
        );
        RETURN;
    END

    -----------------------------------------------------------------------
    -- CONSULTA
    -----------------------------------------------------------------------
    SELECT

        -------------------------------------------------------------------
        -- COMPRA
        -------------------------------------------------------------------
        C.idCompra,
        C.Numero,

        -------------------------------------------------------------------
        -- DETALLE
        -------------------------------------------------------------------
        DC.idDetCompra,

        -------------------------------------------------------------------
        -- PRODUCTO
        -------------------------------------------------------------------
        A.idArticulo,
        A.Codigo,
        A.Articulo,
		ISNULL(A.ManejaSerial,0) AS ManejaSerial,

        -------------------------------------------------------------------
        -- CANTIDADES
        -------------------------------------------------------------------
        DC.Cantidad,
        ISNULL(DC.CantidadRecibida,0)    AS CantidadRecibida,
        ISNULL(DC.CantidadPorFacturar,0) AS CantidadPorFacturar,
        ISNULL(DC.CantidadFacturada,0)   AS CantidadFacturada,

        -------------------------------------------------------------------
        -- ESTADO FACTURACION
        -------------------------------------------------------------------
        ISNULL(DTC.Descripcion,'PENDIENTE') AS EstadoFacturacion

    FROM Compra C

        INNER JOIN DetCompra DC
            ON C.idCompra = DC.idCompra

        INNER JOIN Articulo A
            ON A.idArticulo = DC.idArticulo

        -------------------------------------------------------------------
        -- SOLO ESTADOS FACTURACION
        -------------------------------------------------------------------
        LEFT JOIN DetCompraEstado DTC
            ON DTC.ValorLegacy = DC.EstadoFacturacion
            AND DTC.TipoEstado = 'FACTURACION'
            AND DTC.Activo = 1

    -----------------------------------------------------------------------
    -- SOLO COMPRAS RECIBIDAS
    -----------------------------------------------------------------------
    WHERE C.Estado in (@EstadoCompraRecibido,@EstadoCompraPorFactura)

    -----------------------------------------------------------------------
    -- FILTRO OPCIONAL
    -----------------------------------------------------------------------
    AND
    (
        @idCompra = 0
        OR C.idCompra = @idCompra
    )

    ORDER BY
        C.idCompra DESC,
        A.Articulo;

END

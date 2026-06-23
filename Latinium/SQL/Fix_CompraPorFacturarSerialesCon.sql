-- ============================================================
-- FIX: CompraPorFacturarSerialesCon
-- PROBLEMA: Al guardar la primera factura parcial, el proceso
--   de guardado marca Disponible = 0 en TODOS los DetalleSeriales
--   del idDetCompra, incluyendo los seriales que NO fueron parte
--   de esa factura. El SP filtraba "AND ds.Disponible = 1", por
--   lo que esos seriales pendientes quedan invisibles para la
--   segunda factura.
-- SOLUCIÓN: Eliminar el filtro Disponible = 1.
--   El EstadoFacturacion ya cumple esa función:
--     FAC_FACT (2) excluye los ya facturados
--     FAC_ANU  (5) excluye los anulados
--   Solo se muestran: FAC_PEND(0), FAC_RES(1), FAC_REC(3), FAC_SEL(4)
-- ============================================================

ALTER PROCEDURE [dbo].[CompraPorFacturarSerialesCon]
(
	@idDetCompra INT
)
AS
BEGIN

	SET NOCOUNT ON;

	--------------------------------------------------------------------
	-- ESTADOS FACTURACION
	--------------------------------------------------------------------

	DECLARE @FAC_PEND INT;
	DECLARE @FAC_RES  INT;
	DECLARE @FAC_FACT INT;
	DECLARE @FAC_REC  INT;
	DECLARE @FAC_SEL  INT;
	DECLARE @FAC_ANU  INT;

	SELECT @FAC_PEND = ValorLegacy FROM DetCompraEstado
	WHERE Codigo = 'FAC_PEND' AND TipoEstado = 'FACTURACION' AND Activo = 1;

	SELECT @FAC_RES = ValorLegacy FROM DetCompraEstado
	WHERE Codigo = 'FAC_RES'  AND TipoEstado = 'FACTURACION' AND Activo = 1;

	SELECT @FAC_FACT = ValorLegacy FROM DetCompraEstado
	WHERE Codigo = 'FAC_FACT' AND TipoEstado = 'FACTURACION' AND Activo = 1;

	SELECT @FAC_REC = ValorLegacy FROM DetCompraEstado
	WHERE Codigo = 'FAC_REC'  AND TipoEstado = 'FACTURACION' AND Activo = 1;

	SELECT @FAC_SEL = ValorLegacy FROM DetCompraEstado
	WHERE Codigo = 'FAC_SEL'  AND TipoEstado = 'FACTURACION' AND Activo = 1;

	SELECT @FAC_ANU = ValorLegacy FROM DetCompraEstado
	WHERE Codigo = 'FAC_ANU'  AND TipoEstado = 'FACTURACION' AND Activo = 1;

	--------------------------------------------------------------------
	-- CONSULTA
	--------------------------------------------------------------------

	SELECT

		----------------------------------------------------------------
		-- SERIAL
		----------------------------------------------------------------
		ds.idDetalleSerial,
		ds.idSerial,
		s.Serial,

		----------------------------------------------------------------
		-- DETALLE COMPRA
		----------------------------------------------------------------
		dc.idDetCompra,
		dc.idArticulo,

		----------------------------------------------------------------
		-- ARTICULO
		----------------------------------------------------------------
		a.Codigo,
		a.Articulo,

		----------------------------------------------------------------
		-- INVENTARIO
		----------------------------------------------------------------
		ds.Disponible,

		----------------------------------------------------------------
		-- FACTURACION
		----------------------------------------------------------------
		ds.EstadoFacturacion,
		dce.Descripcion AS EstadoFacturacionDescripcion,

		----------------------------------------------------------------
		-- UI: pre-seleccionado si ya fue marcado FAC_SEL
		----------------------------------------------------------------
		CAST(
			CASE
				WHEN ds.EstadoFacturacion = @FAC_SEL
				THEN 1
				ELSE 0
			END
			AS BIT
		) AS Sel

	FROM DetalleSeriales ds

		INNER JOIN Seriales s
			ON s.idSerial = ds.idSerial

		INNER JOIN DetCompra dc
			ON dc.idDetCompra = ds.idDetCompra

		INNER JOIN Articulo a
			ON a.idArticulo = dc.idArticulo

		LEFT JOIN DetCompraEstado dce
			ON dce.ValorLegacy = ds.EstadoFacturacion
			AND dce.TipoEstado = 'FACTURACION'

	WHERE ds.idDetCompra = @idDetCompra

	--------------------------------------------------------------------
	-- [FIX] Se eliminó el filtro "AND ds.Disponible = 1"
	--   porque el proceso de guardado de factura marca Disponible = 0
	--   en todos los DetalleSeriales del item, incluyendo los que no
	--   fueron facturados. El EstadoFacturacion es el control real:
	--------------------------------------------------------------------

	--------------------------------------------------------------------
	-- SOLO FACTURABLES (excluye FAC_FACT=2 y FAC_ANU=5)
	--------------------------------------------------------------------
	AND ds.EstadoFacturacion IN
	(
		@FAC_PEND,   -- 0  pendiente
		@FAC_RES,    -- 1  reservado
		@FAC_REC,    -- 3  recibido en bodega
		@FAC_SEL     -- 4  seleccionado para esta factura
	)

	ORDER BY
		s.Serial;

END

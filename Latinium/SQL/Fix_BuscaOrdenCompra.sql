-- ============================================================
-- FIX: BuscaOrdenCompra
-- PROBLEMA: El SP no devuelve ODCs en estado 28 (RECEPCIÓN BODEGA),
--   por lo que al buscar en frmCompra no aparece la ODC luego de
--   hacer clic en "Recepción Bodega".
-- SOLUCIÓN: Agregar OR c.Estado = 28 en los estados válidos.
-- ============================================================

ALTER PROCEDURE [dbo].[BuscaOrdenCompra]
	@Criterio VARCHAR(100),
	@Filtro   VARCHAR(200)
AS
BEGIN
	SET NOCOUNT ON;

	DECLARE @Select   VARCHAR(8000),
	        @Caracter VARCHAR(10) = ''''

	------------------------------------------------------------
	-- CONSULTA BASE
	------------------------------------------------------------
	SET @Select = '
	DECLARE @Sel AS BIT = 0

	SELECT
		@Sel AS Sel,
		c.idCompra,
		c.idComprobante,
		c.Consignacion,
		c.Estado,
		c.ContadoCredito,
		c.Fecha,
		ISNULL(c.FechaRecepcion, GETDATE()) AS FechaRecepcion,
		c.idCliente,
		cl.idTipoRuc,
		cl.idGrupoCliente,
		c.idTipoDePago,
		cl.Ruc,
		cl.Nombre,
		c.Bodega,
		cl.DiasCredito,
		cl.DiasCredito1,
		cl.DiasCredito2,
		c.idImportacionTipo,
		c.Numero,
		c.Notas,
		c.SubtotalExcento,
		c.Descuento0,
		c.SubtotalIva,
		c.Descuento,
		c.Iva,
		c.Total
	FROM Compra c
	INNER JOIN Cliente cl ON cl.idCliente = c.idCliente
	WHERE c.idTipoFactura = 2
	AND c.Borrar = 0
	AND CONVERT(DATE, c.Fecha) > ''20141231''
	AND
	(
		--------------------------------------------------------
		-- ESTADOS VÁLIDOS DE NEGOCIO
		--------------------------------------------------------
		c.Estado = 3
		OR (c.Estado = 1 AND c.idImportacionTipo IS NOT NULL AND c.SaldoAnticipo > 0)

		--------------------------------------------------------
		-- NUEVO ESTADO: POR FACTURAR
		--------------------------------------------------------
		OR c.Estado = 20

		--------------------------------------------------------
		-- [FIX] RECEPCIÓN BODEGA: ODC lista para seleccionar
		--       seriales y crear la primera factura parcial
		--------------------------------------------------------
		OR c.Estado = 28
	)'

	------------------------------------------------------------
	-- FILTROS DINÁMICOS
	------------------------------------------------------------
	IF (@Criterio = 'RUC')
		SET @Select = @Select + ' AND cl.RUC LIKE ' + @Caracter + '%' + @Filtro + '%' + @Caracter

	IF (@Criterio = 'NOMBRE')
		SET @Select = @Select + ' AND cl.Nombre LIKE ' + @Caracter + '%' + @Filtro + '%' + @Caracter

	IF (@Criterio = 'NUMERO')
		SET @Select = @Select + ' AND c.Numero LIKE ' + @Caracter + '%' + @Filtro + '%' + @Caracter

	------------------------------------------------------------
	-- ORDENAMIENTO
	------------------------------------------------------------
	SET @Select = @Select + ' ORDER BY c.Fecha ASC'

	------------------------------------------------------------
	-- DEBUG
	------------------------------------------------------------
	PRINT(@Select)

	EXEC(@Select)
END

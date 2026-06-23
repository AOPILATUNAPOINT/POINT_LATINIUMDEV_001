-- ============================================================
-- FIX: MovimientoSerialFacturacion
--
-- PROBLEMA (facturación parcial de ODC con seriales):
--   Al guardar Invoice 1 (solo 1 serial seleccionado de N),
--   el SP hacía dos cosas incorrectas:
--
--   1. INSERT: copiaba TODOS los seriales de @idOrigen al
--      nuevo invoice, aunque el usuario solo seleccionó algunos.
--      Ej: TELE001 (FAC_SEL=4) y TELE002 (FAC_PEND=0) ambos
--      quedaban en Invoice 1, bloqueando a TELE002 para Invoice 2.
--
--   2. UPDATE Disponible=0: marcaba TODOS los seriales de
--      @idOrigen como Disponible=0, incluyendo los no facturados.
--      Ej: TELE002 quedaba Disponible=0 aunque no fue en Invoice 1.
--
-- SOLUCIÓN:
--   Aplicar lógica condicional:
--   - Si EXISTE al menos un serial con EstadoFacturacion=4
--     (FAC_SEL, flujo nuevo frmCompraPorFacturarSeriales):
--       → Solo afectar seriales con EstadoFacturacion=4
--   - Si NO EXISTE ninguno con EstadoFacturacion=4
--     (flujo antiguo, todos los seriales se asignan juntos):
--       → Comportamiento original: afectar todos (Disponible=1)
--
-- ARCHIVOS RELACIONADOS:
--   - frmCompraPorFacturar.cs     : validación restaurada (AND Disponible=1)
--   - Fix_CompraPorFacturarSerialesCon.sql : DESCARTAR, no es necesario
-- ============================================================

ALTER PROCEDURE [dbo].[MovimientoSerialFacturacion]
	@idOrigen       INT,
	@IdDetCompra    INT,
	@IdArticulo     INT,
	@BodegaDestino  INT,
	@Posicion       INT = 0
AS
BEGIN
	SET NOCOUNT ON;

	--------------------------------------------------------------------
	-- DETECCIÓN DEL MODO: ¿flujo nuevo (FAC_SEL) o flujo antiguo?
	--
	-- Flujo nuevo  = frmCompraPorFacturarSeriales marcó FAC_SEL (4)
	-- Flujo antiguo = todos los seriales en FAC_PEND (0) / Disponible=1
	--------------------------------------------------------------------
	DECLARE @bModoParcial BIT = 0;

	IF EXISTS
	(
		SELECT 1
		FROM DetalleSeriales
		WHERE idDetCompra = @idOrigen
		  AND EstadoFacturacion = 4   -- FAC_SEL
	)
		SET @bModoParcial = 1;

	--------------------------------------------------------------------
	-- 1. COPIAR SERIALES AL NUEVO DOCUMENTO (INSERT)
	--
	-- Flujo nuevo  → solo copia los FAC_SEL (4)
	-- Flujo antiguo → copia todos (comportamiento original)
	--------------------------------------------------------------------
	INSERT INTO dbo.DetalleSeriales
	(
		idSerial,
		idDetCompra,
		Bodega,
		Fecha,
		Observaciones,
		Disponible,
		idOrigen,
		Estado,
		Posicion
	)
	SELECT
		ds.idSerial,
		@IdDetCompra,

		CASE c.idTipoFactura
			WHEN 5 THEN @BodegaDestino
			ELSE ds.Bodega
		END,

		GETDATE(),
		ds.Observaciones,
		1,                              -- Disponible = 1 en el nuevo documento
		ds.idDetalleSerial,             -- idOrigen → apunta al registro del ODC

		CASE c.idTipoFactura
			WHEN 1  THEN 1
			WHEN 2  THEN 1
			WHEN 11 THEN 4
			WHEN 13 THEN 4
			WHEN 43 THEN 1
		END,

		@Posicion

	FROM dbo.DetalleSeriales ds
		INNER JOIN dbo.DetCompra dc
			ON dc.idDetCompra = ds.idDetCompra
		INNER JOIN dbo.Compra c
			ON c.idCompra = dc.idCompra

	WHERE dc.idDetCompra = @idOrigen
	  AND dc.idArticulo  = @IdArticulo

	  ------------------------------------------------------------
	  -- [FIX] FILTRO DE SELECCIÓN PARCIAL
	  -- Flujo nuevo : solo FAC_SEL (4)
	  -- Flujo antiguo: todos disponibles (comportamiento previo)
	  ------------------------------------------------------------
	  AND
	  (
		  (@bModoParcial = 1 AND ds.EstadoFacturacion = 4)
		  OR
		  (@bModoParcial = 0 AND ds.Disponible = 1)
	  )

	  /* ============================================
	     ORIGINAL: bloquear duplicados ya facturados
	  ============================================ */
	  AND NOT EXISTS
	  (
		  SELECT 1
		  FROM dbo.DetalleSeriales x
		  WHERE x.idSerial         = ds.idSerial
		    AND x.Estado           = 1
		    AND x.idDetalleSerial <> ds.idDetalleSerial
	  );

	--------------------------------------------------------------------
	-- 2. MARCAR COMO NO DISPONIBLES EN EL DOCUMENTO ORIGEN
	--
	-- Flujo nuevo  → solo los FAC_SEL (4) pasan a Disponible=0
	-- Flujo antiguo → todos (comportamiento original)
	--------------------------------------------------------------------
	IF (ISNULL(@idOrigen, 0) > 0)
	BEGIN
		UPDATE dbo.DetalleSeriales
		SET    Disponible = 0
		WHERE  idDetCompra = @idOrigen
		  AND  Estado NOT IN (100)

		  ------------------------------------------------------------
		  -- [FIX] FILTRO DE SELECCIÓN PARCIAL (igual que INSERT)
		  ------------------------------------------------------------
		  AND
		  (
			  (@bModoParcial = 1 AND EstadoFacturacion = 4)
			  OR
			  (@bModoParcial = 0)
		  );
	END

	--------------------------------------------------------------------
	-- 3. ACTUALIZAR ESTADO EN LA TABLA MAESTRA Seriales
	--    (sin cambio respecto al original)
	--------------------------------------------------------------------
	UPDATE dbo.Seriales
	SET
		Estado =
			CASE c.idTipoFactura
				WHEN 2  THEN 1
				WHEN 1  THEN 17
				WHEN 11 THEN 4
				WHEN 13 THEN 4
				WHEN 36 THEN 4
				WHEN 43 THEN 4
			END,

		Remate =
			CASE c.idTipoFactura
				WHEN 1 THEN 1
				ELSE 0
			END,

		Bodega =
			CASE c.idTipoFactura
				WHEN 5 THEN @BodegaDestino
				ELSE ds.Bodega
			END

	FROM dbo.Seriales s
		INNER JOIN dbo.DetalleSeriales ds
			ON s.idSerial = ds.idSerial
		INNER JOIN dbo.DetCompra dc
			ON dc.idDetCompra = ds.idDetCompra
		INNER JOIN dbo.Compra c
			ON c.idCompra = dc.idCompra

	WHERE dc.idDetCompra = @idOrigen;

END

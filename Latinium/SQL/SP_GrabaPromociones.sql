-- BASE DE DATOS: PROSERVICIOS / PROSERVICIOS_PRUEBAS ÚNICAMENTE
-- Este SP es exclusivo de ProServicios (habitaciones/Tantra).
-- NO aplicar en POINT general ni en otras bases.
USE [PROSERVICIOS]
GO
/****** Object:  StoredProcedure [dbo].[GrabaPromociones]    Script Date: 1/7/2026 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Graba encabezado de una promoción en CompraPromocion.
-- INSERT solo en primer ingreso; UPDATE en modificaciones posteriores.
-- Exec GrabaPromociones 1039, 'PROMOCIÓN', '20181114 00:00', '20181129 23:00', 'GRATIS CILINDRO DE GAS EN COCINAS', 0, 0, 0, 0, 0, 5, 'ACTIVO', 0, 0, 0, 0, 1, 0, 0, 0, 0, '20000101', '20000101'
--
-- PENDIENTE — Segunda Visita 50%:
--   Agregar parámetros opcionales al final (retrocompatibles):
--     @ValidezCuponDias     Int = 0      -- días de vigencia del cupón (ej. 30)
--     @PorcentajeDescuento  Decimal(5,2) = 0  -- descuento de la segunda visita (ej. 50.00)
--   Y columnas correspondientes en CompraPromocion:
--     ALTER TABLE CompraPromocion ADD ValidezCuponDias INT NULL DEFAULT(0)
--     ALTER TABLE CompraPromocion ADD PorcentajeDescuento DECIMAL(5,2) NULL DEFAULT(0)
-- =============================================
ALTER PROCEDURE [dbo].[GrabaPromociones]
	@idPromocion Int, @Tipo Varchar(20), @Desde DateTime = Null, @Hasta DateTime = Null, @Descripcion Varchar(200) = Null,
	@PrecioMinimo Decimal(18, 2) = Null, @PrecioMaximo Decimal(18, 2) = Null,
	@Cuotas Int = Null, @CuotaReal Decimal(18, 2) = Null, @CuotaCalculada Decimal(18, 2) = Null,
	@NumIngreso Int = Null, @Estado Varchar(20) = Null, @idArticulo Int = Null,
	@Variable Bit = 0, @SumaDescuento Bit = 0, @Valor Bit = 0, @Adicional Bit = 0, @Opcional Bit = 0, @Cupon Bit = 0, @Condicion Bit = 0, @ValidaCargaProducto Bit = 0,
	@HoraInicio Datetime = '20000101', @HoraFinaliza Datetime = '20000101', @PuntosFinal Int = 0,
	@TiempoEstadia Int = 0, @ToqueQueda Bit = 0, @Encuesta Bit = 0, @Preventa Bit = 0
	-- PENDIENTE: @ValidezCuponDias Int = 0, @PorcentajeDescuento Decimal(5,2) = 0
AS
BEGIN
	SET NOCOUNT ON;

    Declare @CodPromocion Int

    If Not Exists(Select idArticulo From CompraPromocion Where idPromocion = @idPromocion)
		Begin
			PRINT ('NUEVO')
			Insert Into CompraPromocion (idArticulo, Codigo, NumIngreso, Estado, Valor, Variable, SumaDescuento, Opcional, ValidaCargaProducto, TiempoEstadia, ToqueQueda, Encuesta, Preventa)
				Values (@idArticulo, (Select Count(*) From CompraPromocion) + 1, 1, @Estado, @Valor, @Variable, @SumaDescuento, @Opcional, @ValidaCargaProducto, @TiempoEstadia, @ToqueQueda, @Encuesta, @Preventa)

			Set @CodPromocion = SCOPE_IDENTITY()
		End
	Else
		Begin
			Print ('UPDATE VALCP : ' + CAST(@ValidaCargaProducto AS VARCHAR(10)))

			Update CompraPromocion
			Set Tipo = @Tipo, FechaDesde = @Desde, FechaHasta = @Hasta, Descripcion = @Descripcion,
				PrecioMinimo = @PrecioMinimo, PrecioMaximo = @PrecioMaximo,
				PuntosInicio = @Cuotas, CuotaReal = @CuotaReal, CuotaCalculada = @CuotaCalculada,
				NumIngreso = @NumIngreso, Estado = @Estado, idArticulo = @idArticulo,
				Variable = @Variable, SumaDescuento = @SumaDescuento, Valor = @Valor, Adicional = @Adicional, Opcional = @Opcional, Cupon = @Cupon,
				HoraInicia = @HoraInicio, HoraFinaliza = @HoraFinaliza, Condicion = @Condicion, ValidaCargaProducto = @ValidaCargaProducto, PuntosFinal = @PuntosFinal,
				TiempoEstadia = @TiempoEstadia, ToqueQueda = @ToqueQueda, Encuesta = @Encuesta, Preventa = @Preventa
				-- PENDIENTE: , ValidezCuponDias = @ValidezCuponDias, PorcentajeDescuento = @PorcentajeDescuento
			Where idPromocion = @idPromocion

			Set @CodPromocion = @idPromocion
		End

    Select @CodPromocion
END

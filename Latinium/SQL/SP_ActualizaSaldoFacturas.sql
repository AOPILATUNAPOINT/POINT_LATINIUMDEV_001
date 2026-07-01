USE [POINT]
GO
/****** Object:  StoredProcedure [dbo].[ActualizaSaldoFacturas]    Script Date: 1/7/2026 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Recalcula el saldo pendiente de un documento y actualiza su Estado (7=parcial, 8=pagado)
-- Se llama desde múltiples flujos de cobro con distintos @idTipoFactura:
--   1  = Factura de venta         → actualiza Compra.SaldoAnticipo
--   4  = Orden de compra          → actualiza Compra.SaldoAnticipo + Estado
--   5  = Nota de Crédito          → actualiza Compra.SaldoAnticipo + Estado
--   6  = Guía de remisión         → actualiza Compra.SaldoAnticipo + Estado
--  24  = Anticipo (Return Pass, Gift Card, etc.) → actualiza Anticipos.Saldo + Estado
--  25  = Anticipo especial (AP con caja chica)   → actualiza Anticipos.Saldo + Estado
--  27  = Tipo 27 (liquidación)    → actualiza Compra.SaldoAnticipo (sin cambio de Estado)
--  28  = Tipo 28                  → actualiza Compra.SaldoAnticipo + Estado

ALTER PROCEDURE [dbo].[ActualizaSaldoFacturas]
	@idCompra Int, @idTipoFactura Int, @idFormaDePago Int
AS
BEGIN
	SET NOCOUNT ON;

	Declare @Saldo Decimal(18, 4) = 0, @Pagado Decimal(18, 4) = 0, @Entrada Decimal(18, 4) = 0,
			@AbonosCredito Decimal(18, 2), @NotasDeCredito Decimal(18, 4) = 0,
			@AbonosCP Decimal(18, 2), @Cartera Decimal(18, 2)

	-- =============================================
	-- idTipoFactura = 1 → Factura de Venta
	-- =============================================
	If (@idTipoFactura = 1)
		Begin
			Select @Entrada = ISNULL(SUM(ISNULL(Valor, 0)), 0)
			From DetalleFormaPago
			Where idCompra = @idCompra And Origen = 1 And idFormaPago Not In (6, 9, 10)

			Select @AbonosCredito = ISNULL(SUM(ISNULL(VALOR, 0)), 0)
			From AnticiposCompras
			Where Estado <> 1 And idCompra = @idCompra

			Select @AbonosCP = ISNULL(SUM(ISNULL(ROUND(pcc.ABCapital, 2), 0)), 0)
			From PagosCuotasCredito pcc
				Inner Join Cre_TablaDeAmortizacion ct On pcc.idCre_TablaDeAmortizacion = ct.idCre_TablaDeAmortizacion
			Where pcc.Borrar = 0 And ISNULL(ct.Vendida, 0) = 0 And idCompra = @idCompra And ISNULL(ct.NumeroLote, 0) = 0

			Select @Cartera = ISNULL(ad.Haber, 0)
			From AsientoDetalle ad
				Inner Join Asiento a On a.idAsiento = ad.idAsiento
			Where ad.idFactura = @idCompra And a.Borrar = 0

			Select @NotasDeCredito = ISNULL(SUM(ISNULL(VALOR, 0)), 0)
			From FacturasNotasDeCredito
			Where Estado = 0 And idCompra = @idCompra

			Set @Pagado = ISNULL(@Entrada, 0) + ISNULL(@AbonosCredito, 0) + ISNULL(@AbonosCP, 0) + ISNULL(@Cartera, 0) + ISNULL(@NotasDeCredito, 0)

			Update Compra
			Set SaldoAnticipo = ROUND(Total, 2) - ROUND(Retenido, 2) - ROUND(@Pagado, 2)
			Where idCompra = @idCompra
		End

	-- =============================================
	-- idTipoFactura = 4 → Orden de Compra
	-- =============================================
	If (@idTipoFactura = 4)
		Begin
			If (Select Origen From Compra Where idCompra = @idCompra) = 0
				Begin
					Select @Saldo = ISNULL(SUM(ISNULL(h.Valor, 0)), 0)
					From HistorialAutorizacionDePago h
						Inner Join AutorizacionDePagos a On a.idAutorizacionPago = h.idAutorizacion
					Where a.idTipo = 4 And h.Estado = 0 And a.idOrigen = @idCompra

					Update Compra
					Set SaldoAnticipo = ROUND(Total, 2) - ISNULL(ROUND(Retenido, 2), 0) - ISNULL(ROUND(@Saldo, 2), 0)
					From Compra c
						Inner Join Cliente cl On cl.idCliente = c.idCliente
					Where idCompra = @idCompra

					If ((Select ROUND(ISNULL(SaldoAnticipo, 0), 2) From Compra Where idCompra = @idCompra) = 0)
						Update Compra Set Estado = 8 Where idCompra = @idCompra And idTipoFactura = 4
					If ((Select ROUND(ISNULL(SaldoAnticipo, 0), 2) From Compra Where idCompra = @idCompra) != 0)
						Update Compra Set Estado = 7 Where idCompra = @idCompra And idTipoFactura = 4
				End
			Else
				Begin
					If Exists(Select ad.idDetAsiento From AsientoDetalle ad Inner Join Asiento a On a.idAsiento = ad.idAsiento Where ad.idFactura = @idCompra And a.Borrar = 0)
						Update Compra Set Estado = 8, SaldoAnticipo = 0 Where idCompra = @idCompra And idTipoFactura = 4
					Else
						Update Compra Set Estado = 7, SaldoAnticipo = (Total - ISNULL(Retenido, 0)) Where idCompra = @idCompra And idTipoFactura = 4
				End
		End

	-- =============================================
	-- idTipoFactura = 5 → Nota de Crédito
	-- =============================================
	If (@idTipoFactura = 5)
		Begin
			Declare @ValorDevDinero Decimal(18, 4) = 0, @CruceCrediPoint Decimal(18, 2) = 0, @CruceSolidario Decimal(18, 2) = 0

			Select @Saldo = ISNULL(SUM(ISNULL(Valor, 0)), 0)
			From DetalleFormaPago
			Where idDocumento = @idCompra And Borrar = 0

			Select @NotasDeCredito = ISNULL(SUM(ISNULL(VALOR, 0)), 0)
			From FacturasNotasDeCredito
			Where Estado = 0 And idNotaDeCredito = @idCompra

			Select @CruceCrediPoint = SUM(Abono)
			From PagosCuotasCredito
			Where idAnticipo = @idCompra And Tipo = 1

			Select @ValorDevDinero = SUM(hap.Valor)
			From HistorialAutorizacionDePago hap
				Inner Join AutorizacionDePagos ap On ap.idAutorizacionPago = hap.idAutorizacion
			Where ap.idOrigen = @idCompra And ap.idTipo = 14 And hap.Estado = 0

			Select @CruceSolidario = SUM(Valor)
			From CruceAnticiposSolidario
			Where idAnticipo = @idCompra And AnticipoNC = 1

			Set @Saldo = ISNULL(@Saldo, 0) + ISNULL(@NotasDeCredito, 0) + ISNULL(@ValorDevDinero, 0) + ISNULL(@CruceCrediPoint, 0) + ISNULL(@CruceSolidario, 0)

			Update Compra Set SaldoAnticipo = Total - @Saldo Where idCompra = @idCompra

			Update Compra
			Set Estado = Case When ROUND(SaldoAnticipo, 2) = 0.00 Then 8 Else 7 End
			Where idCompra = @idCompra
		End

	-- =============================================
	-- idTipoFactura = 6 → Guía de Remisión
	-- =============================================
	If (@idTipoFactura = 6)
		Begin
			Select @Saldo = SUM(Valor)
			From HistorialAutorizacionDePago
			Where idAsiento = @idCompra And Estado = 0 And Pago = 2

			Update Compra Set SaldoAnticipo = Total - ISNULL(@Saldo, 0) Where idCompra = @idCompra

			If ((Select ROUND(ISNULL(SaldoAnticipo, 0), 2) From Compra Where idCompra = @idCompra) = 0)
				Update Compra Set Estado = 8 Where idCompra = @idCompra And idTipoFactura = 6

			If ((Select ROUND(ISNULL(SaldoAnticipo, 0), 2) From Compra Where idCompra = @idCompra) > 0)
				Update Compra Set Estado = 7 Where idCompra = @idCompra And idTipoFactura = 6
		End

	-- =============================================
	-- idTipoFactura = 24 → Anticipo (Return Pass, Gift Card, etc.)
	-- Recalcula Anticipos.Saldo descontando lo ya aplicado en AnticiposCompras y DevolucionAnticipos
	-- =============================================
	If (@idTipoFactura = 24)
		Begin
			Declare @CXC Decimal(18, 2), @CrediPOINT Decimal(18, 2), @DevCartera Decimal(18, 2), @Devuelto Decimal(18, 2)

			Select @CXC = ISNULL(SUM(ISNULL(Valor, 0)), 0)
			From AnticiposCompras
			Where idAnticipo = @idCompra And Estado = 0

			Select @Devuelto = ISNULL(SUM(ISNULL(Valor, 0)), 0)
			From DevolucionAnticipos
			Where idAnticipo = @idCompra And Estado < 2

			Set @Saldo = ISNULL(@CXC, 0) + ISNULL(@Devuelto, 0)

			Update Anticipos Set Saldo = Abono - ISNULL(@Saldo, 0) Where idAnticipo = @idCompra

			Update Anticipos
			Set Estado = Case When ROUND(Saldo, 2) = 0 Then 3 When ROUND(Saldo, 2) <> 0 Then 1 End
			Where idAnticipo = @idCompra
		End

	-- =============================================
	-- idTipoFactura = 25 → Anticipo especial (AP / Caja Chica)
	-- =============================================
	If (@idTipoFactura = 25)
		Begin
			If Exists(Select idAutorizacionPago From AutorizacionDePagos Where idOrigen = @idCompra And Estado = 2 And idTipo = 5)
				Begin
					Select @Saldo = ISNULL(SUM(ISNULL(Valor, 0)), 0)
					From HistorialAutorizacionDePago
					Where idAsiento = @idCompra And Pago = 3 And Estado = 0

					Update Anticipos Set Saldo = Abono - @Saldo Where idAnticipo = @idCompra

					Update Anticipos
					Set Estado = Case When ROUND(Saldo, 2) = 0 Then 3 When ROUND(Saldo, 2) > 0 Then 5 End
					Where idAnticipo = @idCompra
				End
			Else
				Begin
					If Exists(Select dcc.idCajaChica From DetalleCajaChica dcc Inner Join CajaChica cc On cc.idCajaChica = dcc.idCajaChica Where idCompra = @idCompra And Estado = 3 And AP = 1)
						Begin
							Select @Saldo = ISNULL(SUM(ISNULL(Valor, 0)), 0)
							From HistorialAutorizacionDePago
							Where idAsiento = @idCompra And Pago = 3 And Estado = 0

							Update Anticipos Set Saldo = Abono - @Saldo Where idAnticipo = @idCompra

							Update Anticipos
							Set Estado = Case When ROUND(Saldo, 2) = 0 Then 3 When ROUND(Saldo, 2) > 0 Then 5 End
							Where idAnticipo = @idCompra
						End
				End
		End

	-- =============================================
	-- idTipoFactura = 27
	-- =============================================
	If (@idTipoFactura = 27)
		Begin
			Select @Saldo = ISNULL(SUM(ISNULL(VALOR, 0)), 0)
			From AnticiposCompras
			Where Estado <> 1 And idTipoFactura = 27 And idCompra = @idCompra

			Select @NotasDeCredito = ISNULL(SUM(ISNULL(VALOR, 0)), 0)
			From FacturasNotasDeCredito
			Where Estado = 0 And idCompra = @idCompra

			Set @Saldo = @Saldo + @NotasDeCredito

			Update Compra Set SaldoAnticipo = Total - @Saldo Where idCompra = @idCompra

			Return
		End

	-- =============================================
	-- idTipoFactura = 28
	-- =============================================
	If (@idTipoFactura = 28)
		Begin
			Select @Saldo = ISNULL(SUM(ISNULL(h.Valor, 0)), 0)
			From HistorialAutorizacionDePago h
				Inner Join AutorizacionDePagos a On a.idAutorizacionPago = h.idAutorizacion
			Where a.idTipo = 15 And h.Estado = 0 And a.idOrigen = @idCompra

			Update Compra Set SaldoAnticipo = Total - ISNULL(@Saldo, 0) Where idCompra = @idCompra

			If ((Select ROUND(ISNULL(SaldoAnticipo, 0), 2) From Compra Where idCompra = @idCompra) = 0)
				Update Compra Set Estado = 8 Where idCompra = @idCompra And idTipoFactura = 28
			If ((Select ROUND(ISNULL(SaldoAnticipo, 0), 2) From Compra Where idCompra = @idCompra) > 0)
				Update Compra Set Estado = 7 Where idCompra = @idCompra And idTipoFactura = 28

			Return
		End
END

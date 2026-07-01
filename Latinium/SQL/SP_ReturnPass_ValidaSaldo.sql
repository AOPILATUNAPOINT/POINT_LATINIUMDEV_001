USE [POINT]
GO
/****** Object:  StoredProcedure [dbo].[ValidaSaldoReturnPass]    Script Date: 1/7/2026 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Verifica si un cliente tiene saldo disponible de la promoción Return Pass (Anticipo Tipo=5)
-- Devuelve: idAnticipo, Saldo
-- Excluye al cliente 237196 (cliente genérico/consumidor final)

ALTER PROCEDURE [dbo].[ValidaSaldoReturnPass]
	@IdCliente Int
AS
BEGIN
	SET NOCOUNT ON;

    Select idAnticipo, Saldo
	From Anticipos
	Where ClienteProveedor = 0
		And Estado = 1
		And Tipo = 5
		And idCliente = @IdCliente
		And idCliente <> 237196
END

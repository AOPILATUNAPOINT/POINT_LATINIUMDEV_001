-- BASE DE DATOS: PROSERVICIOS / PROSERVICIOS_PRUEBAS ÚNICAMENTE
USE [PROSERVICIOS]
GO
/****** Object:  StoredProcedure [dbo].[ConsultaIndividualPromocionesLotesServicios] ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Devuelve el encabezado de una promoción para cargar frmPromociones.
-- El C# lee por índice ordinal — el orden de columnas es crítico:
--   0  idPromocion          → txtIdPromocion
--   1  idArticulo           → txtidArticulo
--   2  Tipo                 → cmbTipo
--   3  FechaDesde           → dtInicio
--   4  FechaHasta           → dtFin
--   5  Descripcion          → txtDescripcion
--   6  NumIngreso           → iNumIngreso
--   7  Estado               → lblFinalizado
--   8  Cupon                → chkCupon
--   9  Variable             → chkVariable
--  10  SumaDescuento        → chkSumaDescuentos
--  11  Adicional            → chkAdicional
--  12  Valor                → chkValor
--  13  Opcional             → chkOpcional
--  14  HoraInicia           → dtHInicio
--  15  HoraFinaliza         → dtFinal
--  16  Condicion            → chkTarjetaVIP
--  17  PuntosInicio         → txtPuntosInicial
--  18  PuntosFinal          → txtPuntosFinal
--  19  TiempoEstadia        → txtTiempo
--  20  ToqueQueda           → chkToqueQueda
--  21  Encuesta             → chkEncuesta
--  22  Preventa             → (no usado en UI actualmente)
--  23  ValidezCuponDias     → txtValidezCuponDias   [NUEVO — Segunda Visita]
--  24  PorcentajeDescuento  → txtPorcDescuentoSV    [NUEVO — Segunda Visita]
-- exec [ConsultaIndividualPromocionesLotesServicios] 346
ALTER PROCEDURE [dbo].[ConsultaIndividualPromocionesLotesServicios]
    @idPromocion Int
AS
BEGIN
    SET NOCOUNT ON;

    Select
        c.idPromocion,
        ISNULL(idArticulo, 0),
        Tipo,
        FechaDesde,
        FechaHasta,
        Descripcion,
        NumIngreso,
        Estado,
        Cupon,
        Variable,
        SumaDescuento,
        Adicional,
        Valor,
        Opcional,
        HoraInicia,
        HoraFinaliza,
        Condicion,
        PuntosInicio,
        PuntosFinal,
        ISNULL(TiempoEstadia, 0)     AS TiempoEstadia,
        ISNULL(ToqueQueda, 0)        AS ToqueQueda,
        Encuesta,
        ISNULL(Preventa, 0)          AS Preventa,
        ISNULL(ValidezCuponDias, 0)  AS ValidezCuponDias,
        ISNULL(PorcentajeDescuento, 0) AS PorcentajeDescuento
    From CompraPromocion c
    Where c.idPromocion = @idPromocion
END

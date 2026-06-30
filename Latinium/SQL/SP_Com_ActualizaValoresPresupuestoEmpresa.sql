IF OBJECT_ID('[dbo].[Com_ActualizaValoresPresupuestoEmpresa]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_ActualizaValoresPresupuestoEmpresa]
GO

CREATE PROCEDURE [dbo].[Com_ActualizaValoresPresupuestoEmpresa]
    @Periodo       Int,
    @Local         Varchar(200),
    @Mes           Varchar(50),
    @Valor         Decimal(18, 2)
AS
BEGIN
    SET NOCOUNT ON;

    Declare @Bodega Int, @idPeriodoMes Int

    Select @Bodega       = Bodega      From Bodega    Where Nombre      = @Local
    Select @idPeriodoMes = idPeriodoMes From PeriodoMes Where Descripcion = @Mes

    Update Com_Empresa
    Set    Valor = @Valor
    Where  Periodo      = @Periodo
      And  Bodega       = @Bodega
      And  idPeriodoMes = @idPeriodoMes

    /* ACTUALIZA ASIGNACION VENDEDORES */

    Declare @Ceros        Varchar(2),
            @PeriodoAV    Varchar(7),
            @DiasMes      Int,
            @PeriodoActual Varchar(8)

    Set @Ceros         = SUBSTRING('00', 1, 2 - LEN(@idPeriodoMes))
    Set @PeriodoAV     = @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '/' + CAST(@Periodo As Varchar(4))
    Set @PeriodoActual = CONVERT(varchar(4), @Periodo) + @Ceros + CAST(@idPeriodoMes As Varchar(2)) + '01'
    Set @DiasMes       = DATEPART(DAY, DATEADD(s, -1, DATEADD(mm, DATEDIFF(m, 0, CONVERT(Datetime, @PeriodoActual, 112)) + 1, 0)))

    If Exists(Select idCom_AsignacionDeVendedores From Com_AsignacionDeVendedores Where Bodega = @Bodega And Periodo = @PeriodoAV)
    Begin
        Update Com_AsignacionDeVendedores
        Set    MetaMensual = dbo.Com_RetornaMetaPersonalLocalesPorCargo(@Periodo, @idPeriodoMes, @Bodega, idCargo, 1) / @DiasMes * NumeroDias
        Where  Bodega       = @Bodega
          And  Periodo      = @PeriodoAV
          And  idCom_Estado In (1, 3)
    End

    If Exists(Select idCom_AsignacionJefes From Com_AsignacionJefes Where Bodega = @Bodega And Periodo = @PeriodoAV)
        Update Com_AsignacionJefes
        Set    MetaMensual = (@Valor / @DiasMes) * Dias
        Where  Bodega  = @Bodega
          And  Periodo = @PeriodoAV

END
GO

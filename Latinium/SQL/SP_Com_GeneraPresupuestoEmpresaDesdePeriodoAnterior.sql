IF OBJECT_ID('[dbo].[Com_GeneraPresupuestoEmpresaDesdePeriodoAnterior]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_GeneraPresupuestoEmpresaDesdePeriodoAnterior]
GO

CREATE PROCEDURE [dbo].[Com_GeneraPresupuestoEmpresaDesdePeriodoAnterior]
    @Periodo       Int,
    @PE            Int,
    @PersonalLocal Bit,
    @PL            Int,
    @AsigVend      Bit,
    @AV            Int
AS
BEGIN
    SET NOCOUNT ON;

    /* GENERA PRESUPUESTO EMPRESA */
    Insert Into Com_Empresa (Periodo, Bodega, idPeriodoMes, Valor)
    Select @Periodo, c.Bodega, idPeriodoMes,
        Case @PE
            When 1 Then Valor
            When 2 Then Case idPeriodoMes When 1  Then Valor Else 0 End
            When 3 Then Case idPeriodoMes When 12 Then Valor Else 0 End
        End
    From Com_Empresa c
    Inner Join Bodega b On b.Bodega = c.Bodega
    Where Periodo = @Periodo - 1
      And DATEPART(YEAR, ISNULL(b.FechaCierre, '20301231')) > @Periodo

    Insert Into Com_PresupuestoCredito (Periodo, Bodega, idPeriodoMes, Valor)
    Select @Periodo, c.Bodega, idPeriodoMes,
        Case @PE
            When 1 Then Valor
            When 2 Then Case idPeriodoMes When 1  Then Valor Else 0 End
            When 3 Then Case idPeriodoMes When 12 Then Valor Else 0 End
        End
    From Com_PresupuestoCredito c
    Inner Join Bodega b On b.Bodega = c.Bodega
    Where Periodo = @Periodo - 1
      And DATEPART(YEAR, ISNULL(b.FechaCierre, '20301231')) > @Periodo

    /* GENERA PERSONAL POR LOCAL */
    Declare @PeriodoInicial Varchar(7)
    Set @PeriodoInicial = '01/' + CAST(@Periodo As Varchar(4))

    If (@PersonalLocal = 1)
    Begin
        If (@PL = 1)
        Begin
            Insert Into Com_PersonalLocales (Periodo, Bodega, idCargo, Cantidad)
            Select SUBSTRING(Periodo, 1, 2) + '/' + CAST(@Periodo As Varchar(4)), p.Bodega, idCargo, Cantidad
            From Com_PersonalLocales p
            Inner Join Bodega b On b.Bodega = p.Bodega
            Where SUBSTRING(Periodo, 4, 7) = @Periodo - 1
              And DATEPART(YEAR, ISNULL(b.FechaCierre, '20301231')) > @Periodo
        End
        Else
        Begin
            If (@PL = 2) Set @PeriodoInicial = '01/' + CAST(@Periodo - 1 As Varchar(4))
            If (@PL = 3) Set @PeriodoInicial = '12/' + CAST(@Periodo - 1 As Varchar(4))

            Insert Into Com_PersonalLocales (Periodo, Bodega, idCargo, Cantidad)
            Select SUBSTRING(Periodo, 1, 2) + '/' + CAST(@Periodo As Varchar(4)), p.Bodega, idCargo, Cantidad
            From Com_PersonalLocales p
            Inner Join Bodega b On b.Bodega = p.Bodega
            Where Periodo = @PeriodoInicial
              And DATEPART(YEAR, ISNULL(b.FechaCierre, '20301231')) > @Periodo
        End
    End
    Else
    Begin
        Exec Com_GeneraPersonalLocales @PeriodoInicial
    End

    /* GENERA PRESUPUESTO MENSUAL POR LOCAL */
    Insert Into Com_PresupuestoMensualLocales (Periodo, Bodega, idPeriodoMes, idCargo, Valor)
    Select @Periodo, p.Bodega, idPeriodoMes, idCargo,
        Case @PE
            When 1 Then Valor
            When 2 Then Case idPeriodoMes When 1  Then Valor Else 0 End
            When 3 Then Case idPeriodoMes When 12 Then Valor Else 0 End
        End
    From Com_PresupuestoMensualLocales p
    Inner Join Bodega b On b.Bodega = p.Bodega
    Where Periodo = @Periodo - 1
      And DATEPART(YEAR, ISNULL(b.FechaCierre, '20301231')) > @Periodo

    If (@AsigVend = 1)
    Begin
        If (@AV = 1)
        Begin
            Insert Into Com_AsignacionDeVendedores (Periodo, Bodega, idCargo, idPersonal, idCom_Rango, Desde, Hasta, Diario, Mensual, idCom_Estado)
            Select SUBSTRING(Periodo, 1, 2) + '/' + CAST(@Periodo As Varchar(4)),
                   a.Bodega, idCargo, idPersonal, idCom_Rango,
                   DATEADD(YEAR, 1, Desde), DATEADD(YEAR, 1, Hasta),
                   Diario, Mensual, idCom_Estado
            From Com_AsignacionDeVendedores a
            Inner Join Bodega b On b.Bodega = a.Bodega
            Where SUBSTRING(Periodo, 4, 7) = @Periodo - 1
              And DATEPART(YEAR, ISNULL(b.FechaCierre, '20301231')) > @Periodo
              And idCom_Estado = 1
        End
    End

END
GO

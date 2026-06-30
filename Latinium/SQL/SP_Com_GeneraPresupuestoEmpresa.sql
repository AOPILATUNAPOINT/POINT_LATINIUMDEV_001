IF OBJECT_ID('[dbo].[Com_GeneraPresupuestoEmpresa]') IS NOT NULL
    DROP PROCEDURE [dbo].[Com_GeneraPresupuestoEmpresa]
GO

-- Exec Com_GeneraPresupuestoEmpresa 2026
CREATE PROCEDURE [dbo].[Com_GeneraPresupuestoEmpresa]
    @Periodo Int
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @Bodega Int

    /* GENERAL */
    Declare cCom_Empresa Cursor For
        Select Bodega From Bodega
        Where  Activo = 1 And Factura = 1 And Credito = 1
          And  CONVERT(Date, ISNULL(FechaCierre, '20301231')) > '20241231'
        Order By Nombre

    OPEN cCom_Empresa
    Fetch Next From cCom_Empresa Into @Bodega
    While (@@FETCH_STATUS = 0)
    Begin
        Insert Into Com_Empresa (Periodo, Bodega, idPeriodoMes, Valor)
        Select @Periodo, @Bodega, idPeriodoMes, 0 From PeriodoMes Order By idPeriodoMes
        Fetch Next From cCom_Empresa Into @Bodega
    End
    Close cCom_Empresa
    DeAllocate cCom_Empresa

    /* CREDITO */
    Declare cCom_EmpresaCredito Cursor For
        Select Bodega From Bodega
        Where  Activo = 1 And Factura = 1 And Credito = 1
          And  CONVERT(Date, ISNULL(FechaCierre, '20301231')) > '20241231'
        Order By Nombre

    OPEN cCom_EmpresaCredito
    Fetch Next From cCom_EmpresaCredito Into @Bodega
    While (@@FETCH_STATUS = 0)
    Begin
        Insert Into Com_PresupuestoCredito (Periodo, Bodega, idPeriodoMes, Valor)
        Select @Periodo, @Bodega, idPeriodoMes, 0 From PeriodoMes Order By idPeriodoMes
        Fetch Next From cCom_EmpresaCredito Into @Bodega
    End
    Close cCom_EmpresaCredito
    DeAllocate cCom_EmpresaCredito

    /* GARANTIA EXTENDIDA */
    Declare cCom_EmpresaGEx Cursor For
        Select Bodega From Bodega
        Where  Activo = 1 And Factura = 1 And Credito = 1
          And  CONVERT(Date, ISNULL(FechaCierre, '20301231')) > '20241231'
        Order By Nombre

    OPEN cCom_EmpresaGEx
    Fetch Next From cCom_EmpresaGEx Into @Bodega
    While (@@FETCH_STATUS = 0)
    Begin
        Insert Into Com_PresupuestoEmpresaGE (Periodo, Bodega, idPeriodoMes, Valor)
        Select @Periodo, @Bodega, idPeriodoMes, 0 From PeriodoMes Order By idPeriodoMes
        Fetch Next From cCom_EmpresaGEx Into @Bodega
    End
    Close cCom_EmpresaGEx
    DeAllocate cCom_EmpresaGEx

    /* META CONSULTA CEDULAS */
    Declare cCom_EmpresaConsultaCedulas Cursor For
        Select Bodega From Bodega
        Where  Activo = 1 And Factura = 1 And Credito = 1
          And  CONVERT(Date, ISNULL(FechaCierre, '20301231')) > '20241231'
        Order By Nombre

    OPEN cCom_EmpresaConsultaCedulas
    Fetch Next From cCom_EmpresaConsultaCedulas Into @Bodega
    While (@@FETCH_STATUS = 0)
    Begin
        Insert Into Com_MetaConsultaCedulas (Periodo, Bodega, idPeriodoMes, Valor)
        Select @Periodo, @Bodega, idPeriodoMes, 0 From PeriodoMes Order By idPeriodoMes
        Fetch Next From cCom_EmpresaConsultaCedulas Into @Bodega
    End
    Close cCom_EmpresaConsultaCedulas
    DeAllocate cCom_EmpresaConsultaCedulas

END
GO

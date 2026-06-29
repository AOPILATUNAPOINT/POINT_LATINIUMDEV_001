-- CREQ: Cierre de Operaciones con Descuento Especial
-- Tabla de auditoría — única tabla requerida por el proceso

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'CierreOperacionDescuentoLog')
CREATE TABLE [dbo].[CierreOperacionDescuentoLog] (
    idCierreOperacion   INT           IDENTITY(1,1) NOT NULL,
    idCompra            INT           NOT NULL,
    Cedula              VARCHAR(20)   NOT NULL,
    Cliente             VARCHAR(200)  NULL,
    Factura             VARCHAR(50)   NOT NULL,
    Segmento            VARCHAR(100)  NULL,
    SaldoOriginal       DECIMAL(18,2) NOT NULL,
    ValorDescuento      DECIMAL(18,2) NOT NULL,
    MontoAjuste         DECIMAL(18,2) NOT NULL,
    NumAutorizacion     VARCHAR(100)  NOT NULL,
    UsuarioAutoriza     VARCHAR(50)   NOT NULL,
    UsuarioEjecuta      VARCHAR(50)   NOT NULL,
    idAsiento           INT           NULL,
    FechaProceso        DATETIME      NOT NULL DEFAULT GETDATE(),
    Estado              VARCHAR(20)   NOT NULL DEFAULT 'PROCESADO',
    Observacion         VARCHAR(500)  NULL,
    CONSTRAINT PK_CierreOperacionDescuentoLog PRIMARY KEY (idCierreOperacion)
)
GO

-- ============================================================
-- SEGURIDAD — CREDIPOINT > COBRANZAS (padre=679)
-- ============================================================

-- Verificar antes de insertar
-- SELECT id_seg, nombre FROM Seguridad WHERE id_seg IN (9100, 9101)

IF NOT EXISTS (SELECT 1 FROM Seguridad WHERE id_seg = 2231)
INSERT INTO Seguridad (
    id_seg, nombre, padre, pos, tipo, menu, codigo, Visible,
    n_1,Vn_1,n_2,Vn_2,n_3,Vn_3,n_4,Vn_4,n_5,Vn_5,n_6,Vn_6,
    n_7,Vn_7,n_8,Vn_8,n_9,Vn_9,n_10,Vn_10,n_11,Vn_11,n_12,Vn_12,
    n_13,Vn_13,n_14,Vn_14,n_15,Vn_15,n_16,Vn_16,n_17,Vn_17,n_18,Vn_18,
    n_19,Vn_19,n_20,Vn_20,n_21,Vn_21,n_22,Vn_22,n_23,Vn_23,n_24,Vn_24,
    n_25,Vn_25,n_26,Vn_26,n_27,Vn_27,n_28,Vn_28,n_29,Vn_29,n_30,Vn_30,
    n_31,Vn_31,n_32,Vn_32,n_33,Vn_33,n_34,Vn_34,n_35,Vn_35,n_36,Vn_36,
    Sel_ReqPantalla)
VALUES (
    2231, 'CIERRE OPERACIONES DESCUENTO ESPECIAL', 679, 2231, 2, NULL, '0843', 1,
    1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)

IF NOT EXISTS (SELECT 1 FROM Seguridad WHERE id_seg = 2232)
INSERT INTO Seguridad (
    id_seg, nombre, padre, pos, tipo, menu, codigo, Visible,
    n_1,Vn_1,n_2,Vn_2,n_3,Vn_3,n_4,Vn_4,n_5,Vn_5,n_6,Vn_6,
    n_7,Vn_7,n_8,Vn_8,n_9,Vn_9,n_10,Vn_10,n_11,Vn_11,n_12,Vn_12,
    n_13,Vn_13,n_14,Vn_14,n_15,Vn_15,n_16,Vn_16,n_17,Vn_17,n_18,Vn_18,
    n_19,Vn_19,n_20,Vn_20,n_21,Vn_21,n_22,Vn_22,n_23,Vn_23,n_24,Vn_24,
    n_25,Vn_25,n_26,Vn_26,n_27,Vn_27,n_28,Vn_28,n_29,Vn_29,n_30,Vn_30,
    n_31,Vn_31,n_32,Vn_32,n_33,Vn_33,n_34,Vn_34,n_35,Vn_35,n_36,Vn_36,
    Sel_ReqPantalla)
VALUES (
    2232, 'EJECUTAR CIERRE', 2231, 2232, 3, NULL, '09100001', 1,
    1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
GO

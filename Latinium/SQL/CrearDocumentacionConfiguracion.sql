-- Tabla de configuracion de URLs de documentacion
-- Permite cambiar el servidor de documentacion sin recompilar

IF NOT EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_NAME = 'DocumentacionConfiguracion'
)
BEGIN
    CREATE TABLE DocumentacionConfiguracion (
        idDocumentacionConfiguracion          INT           IDENTITY(1,1) PRIMARY KEY,
        Clave       VARCHAR(100)  NOT NULL,
        Descripcion VARCHAR(500)  NULL,
        Url         VARCHAR(1000) NOT NULL,
        Activo      BIT           NOT NULL DEFAULT 1,
        FechaCreacion DATETIME    NOT NULL DEFAULT GETDATE()
    )
END
GO

IF NOT EXISTS (
    SELECT 1 FROM DocumentacionConfiguracion WHERE Clave = 'ManualFideicomiso'
)
BEGIN
    INSERT INTO DocumentacionConfiguracion (Clave, Descripcion, Url)
    VALUES (
        'ManualFideicomiso',
        'Manual de usuario - Venta de Cartera Fideicomiso',
        'http://192.168.2.27:3333/latinium/CREDIPOINT/CARTERA/FIDEICOMISO/venta-cartera-fideicomiso'
    )
END
GO

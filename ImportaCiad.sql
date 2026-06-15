--use Erne
SET nocount on
ALTER TABLE Compra DISABLE TRIGGER SaldoCxCActualiza
ALTER TABLE Compra DISABLE TRIGGER SaldoCxCBorrar
ALTER TABLE Compra DISABLE TRIGGER SaldoCxCInserta
ALTER TABLE DetCompra DISABLE TRIGGER SaldoCosto
ALTER TABLE DetCompra DISABLE TRIGGER SaldoCostoI
ALTER TABLE DetCompra DISABLE TRIGGER BorraDetCompra
--ALTER TABLE Cliente DISABLE TRIGGER AuditoriaClienteE
--ALTER TABLE Articulo DISABLE TRIGGER AuditoriaNuevo
ALTER TABLE Pago DISABLE TRIGGER PagoActualiza
ALTER TABLE Pago DISABLE TRIGGER PagoBorra
--ALTER TABLE Pago DISABLE TRIGGER PagoInserta
ALTER TABLE OrdenProduccionDetalle DISABLE TRIGGER ActualizaDetCompra
ALTER TABLE OrdenProduccion DISABLE TRIGGER CreaCompra
ALTER TABLE OrdenProduccion DISABLE TRIGGER CambiaCompraOP
--ALTER TABLE Retencion DISABLE TRIGGER RetencionInserta
ALTER TABLE Retencion DISABLE TRIGGER RetencionBorra
ALTER TABLE Retencion DISABLE TRIGGER RetencionActualiza

/*
ALTER TABLE DetCompra DISABLE TRIGGER ActualizaTotal
ALTER TABLE DetCompra DISABLE TRIGGER InsertaSaldo
ALTER TABLE DetCompra DISABLE TRIGGER ActualizaSaldo
-------------------------------------


ALTER TABLE DetCompra ENABLE TRIGGER ActualizaTotal
ALTER TABLE DetCompra ENABLE TRIGGER ActualizaSaldo
ALTER TABLE DetCompra ENABLE TRIGGER InsertaSaldo
-------------------------------------

EXEC SaldoCxCyP
ALTER TABLE Compra ENABLE TRIGGER SaldoCxCActualiza
ALTER TABLE Compra ENABLE TRIGGER SaldoCxCBorrar
ALTER TABLE Compra ENABLE TRIGGER SaldoCxCInserta
ALTER TABLE DetCompra ENABLE TRIGGER SaldoCosto
ALTER TABLE Cliente ENABLE TRIGGER AuditoriaClienteE
ALTER TABLE Articulo ENABLE TRIGGER AuditoriaNuevo
ALTER TABLE Pago ENABLE TRIGGER PagoActualiza
ALTER TABLE Pago ENABLE TRIGGER PagoBorra
ALTER TABLE Pago ENABLE TRIGGER PagoInserta
ALTER TABLE OrdenProduccionDetalle ENABLE TRIGGER ActualizaDetCompra
ALTER TABLE OrdenProduccion ENABLE TRIGGER CreaCompra
ALTER TABLE OrdenProduccion ENABLE TRIGGER CambiaCompraOP
ALTER TABLE OrdenProduccion ENABLE TRIGGER BorraCompraOP
ALTER TABLE Retencion ENABLE TRIGGER RetencionInserta
ALTER TABLE Retencion ENABLE TRIGGER RetencionBorra
ALTER TABLE Retencion ENABLE TRIGGER RetencionActualiza
*/
-- Limpia tablas antes de importar datos de CIAD
Delete DetCompra
Delete GuiaRemision
Delete DetCompraTransf
Delete Compra
Delete CompraFormaPago
Delete CompraTransf
Delete Presupuesto
Delete PresupuestoDetalle
Delete Producto
Delete ProductoDetalle
Delete OrdenProduccion
Delete OrdenProduccionDetalle
go
DBCC CHECKIDENT(OrdenProduccion, RESEED, 0)
DBCC CHECKIDENT(OrdenProduccionDetalle, RESEED, 0)
DBCC CHECKIDENT(Producto, RESEED, 0)
DBCC CHECKIDENT(ProductoDetalle, RESEED, 0)
DBCC CHECKIDENT(Presupuesto, RESEED, 0)
DBCC CHECKIDENT(PresupuestoDetalle, RESEED, 0)
DBCC CHECKIDENT(DetCompra, RESEED, 0)
DBCC CHECKIDENT(DetCompraTransf, RESEED, 0)
DBCC CHECKIDENT(Compra, RESEED, 0)
DBCC CHECKIDENT(CompraFormaPago, RESEED, 0)
DBCC CHECKIDENT(CompraTransf, RESEED, 0)
GO

Delete Auditoria
Delete OrdenProduccion
Delete PagoLote
Delete Personal
Delete RolConcepto
DBCC CHECKIDENT(OrdenProduccion, RESEED, 0)
DBCC CHECKIDENT(OrdenProduccionDetalle, RESEED, 0)
DBCC CHECKIDENT(PagoLote, RESEED, 0)
DBCC CHECKIDENT(Personal, RESEED, 0)
DBCC CHECKIDENT(RolConcepto, RESEED, 0)
GO

Delete Proyecto
Delete SubProyecto
Delete Articulo
Delete ArticuloArancel
Delete ArticuloClase
Delete ArticuloGrupo
Delete ArticuloMarca
Delete ArticuloSeccion
Delete ArticuloSubGrupo
Delete ArticuloUnidad
GO
DBCC CHECKIDENT(Proyecto, RESEED, 0)
DBCC CHECKIDENT(SubProyecto, RESEED, 0)
DBCC CHECKIDENT(Articulo, RESEED, 0)
DBCC CHECKIDENT(ArticuloArancel, RESEED, 0)
DBCC CHECKIDENT(ArticuloGrupo, RESEED, 0)
DBCC CHECKIDENT(ArticuloSubGrupo, RESEED, 0)
DBCC CHECKIDENT(ArticuloMarca, RESEED, 0)
DBCC CHECKIDENT(ArticuloUnidad, RESEED, 0)
DBCC CHECKIDENT(ArticuloClase, RESEED, 0)
DBCC CHECKIDENT(ArticuloSeccion, RESEED, 0)
GO

-- Bases relacionadas de Clientes
Delete Cliente
Delete ClienteCiudad
Delete ClienteGrupo
Delete ClienteSector
go
DBCC CHECKIDENT(Cliente, RESEED, 0)
DBCC CHECKIDENT(ClienteGrupo, RESEED, 0)
DBCC CHECKIDENT(ClienteCiudad, RESEED, 0)
DBCC CHECKIDENT(ClienteSector, RESEED, 0)
go

-- plan de Cuentas
DELETE CuentaCorriente
DELETE Cuenta
go
DBCC CHECKIDENT(CuentaCorriente, RESEED, 0)
DBCC CHECKIDENT(Cuenta, RESEED, 0)
go

-- Borra Asientos
DELETE AsientoDetalle
DELETE Asiento
DELETE Conciliacion
go
DBCC CHECKIDENT(Asiento, RESEED, 0)
DBCC CHECKIDENT(AsientoDetalle, RESEED, 0)
DBCC CHECKIDENT(Conciliacion, RESEED, 0)
go

-- Pagos
DELETE Pago
GO
DBCC CHECKIDENT(Pago, RESEED, 0)
GO

-- Personal
DELETE Personal
GO
DBCC CHECKIDENT(Personal, RESEED, 0)
GO

DELETE Retencion
DELETE RetDetalle
DELETE PorcIva
DELETE PorcIce
DELETE RetencionTipo
DELETE RetencionIva
DELETE RetencionRenta
GO
DBCC CHECKIDENT(Retencion, RESEED, 0)
DBCC CHECKIDENT(PorcIva, RESEED, 0)
DBCC CHECKIDENT(PorcIce, RESEED, 0)
DBCC CHECKIDENT(RetencionTipo, RESEED, 0)
DBCC CHECKIDENT(RetencionIva, RESEED, 0)
DBCC CHECKIDENT(RetencionRenta, RESEED, 0)
GO

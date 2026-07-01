# Funcional — Promoción "Segunda Visita Mitad de Precio"

**Solicitante:** Carlos Valle — Gerente Tantra  
**Sistema:** Latinium — Módulo Facturación Habitaciones  
**Fecha:** 2026-07-01  
**Estado general:** En análisis / Pendiente de desarrollo

---

## Resumen ejecutivo

El cliente paga la habitación al precio normal. En esa misma factura se cobra también un cupón anticipado para una **segunda estadía con 50% de descuento**. Al regresar, el cajero canjea el cupón y el sistema aplica automáticamente el descuento, bloqueando cualquier intento de fraude (cambio de categoría, días no permitidos, reutilización, etc.).

Vigencia de la promoción: **Domingo a Jueves** · Locales: **Tantra Design, Gold, Loft, Platinum** · Validez del cupón: **30 días desde emisión** · Vigencia de la config: **3 meses** (parametrizable)

---

## Infraestructura existente que se reutiliza

| Componente existente | Dónde vive | Rol en la nueva promo |
|---|---|---|
| Tabla `Anticipos` | BD | Almacena el cupón como anticipo `Tipo = X` (nuevo tipo) |
| Tabla `AnticiposCompras` | BD | Vincula el cupón a la factura de canje |
| Tabla `TiposDeAnticipos` | BD | Registra el nombre del reporte Crystal a imprimir |
| `frmAnticiposClientes` | UI | Formulario de cobro del anticipo — se reutiliza |
| `btnReturnPass` en `frmFacturacionHabitaciones` | UI | Botón de venta — se amplía o duplica para la nueva promo |
| `ValidaPromoReturnPass()` | `frmFacturacionHabitaciones.cs:6127` | Lógica de validación al cargar habitación — se extiende |
| `ImpresionAutomaticaAnticipos()` | `frmFacturacionHabitaciones.cs:7611` | Imprime el cupón usando Crystal Reports — se reutiliza |
| `ValidaSaldoReturnPass` SP | BD · `SQL/SP_ReturnPass_ValidaSaldo.sql` | Base para el nuevo SP de validación |
| `ActualizaSaldoFacturas` SP (tipo 24) | BD · `SQL/SP_ActualizaSaldoFacturas.sql` | Descuenta saldo del anticipo al canjear |
| `VentaFPagoLot` | UI | Pantalla de cobro — ya recibe `bUsoPromActiva` y `dValorPromActiva` |

---

## Lo que hay que construir

### 1. Base de datos

#### 1.1 Nueva tabla de configuración de promociones (NUEVO)
```sql
CREATE TABLE PromocionesConfig (
    idPromocion       INT IDENTITY PRIMARY KEY,
    Nombre            VARCHAR(100),
    Tipo              INT,           -- referencia a TiposDeAnticipos
    DiasPermitidos    VARCHAR(20),   -- ej: '0,1,2,3,4' (Dom=0 ... Jue=4)
    ValidezCuponDias  INT,           -- 30 días
    ValidezPromMeses  INT,           -- 3 meses
    DescuentoPorciento DECIMAL(5,2), -- 50.00
    Activo            BIT DEFAULT 1,
    FechaInicio       DATE,
    FechaFin          DATE
)
```

#### 1.2 Nueva tabla de cupones emitidos (NUEVO)
```sql
CREATE TABLE CuponesSegundaVisita (
    idCupon           INT IDENTITY PRIMARY KEY,
    idAnticipo        INT,           -- FK Anticipos
    idCompraOrigen    INT,           -- factura donde se compró
    idCliente         INT,
    idHabitacion      INT,
    idCategoria       INT,
    idBodega          INT,           -- local donde se emitió
    ValorOriginal     DECIMAL(18,2),
    ValorDescuento    DECIMAL(18,2),
    ValorCancelado    DECIMAL(18,2), -- lo que pagó por el cupón
    FechaEmision      DATETIME,
    FechaVencimiento  DATETIME,      -- FechaEmision + 30 días
    Estado            INT,           -- 1=Pendiente, 3=Utilizado, 4=Anulado, 5=Vencido
    idCompraUso       INT NULL       -- factura donde se canjeó
)
```

#### 1.3 Nuevo tipo en `TiposDeAnticipos` (NUEVO)
```sql
INSERT INTO TiposDeAnticipos (Tipo, Descripcion, Reporte)
VALUES (6, 'Segunda Visita 50%', 'RptCuponSegundaVisita.rpt')
-- El Tipo = 6 diferencia este cupón del Return Pass normal (Tipo = 5)
```

#### 1.4 Nuevos locales en configuración (NUEVO)
- Parametrizar qué `idBodega` corresponde a Tantra Design, Gold, Loft, Platinum
- Puede ser una columna en `PromocionesConfig` o una tabla de relación

---

#### 1.5 SP `ValidaConfigPromoSegundaVisita` (NUEVO)
```sql
-- Devuelve la configuración activa de la promo para una fecha y bodega dadas
-- El C# llama esto al iniciar la factura para saber si el botón debe aparecer
CREATE PROCEDURE ValidaConfigPromoSegundaVisita
    @idBodega INT, @Fecha DATETIME
AS BEGIN
    -- Valida: promo activa + bodega permitida + día de la semana permitido + dentro de FechaInicio/FechaFin
    SELECT idPromocion, Nombre, DescuentoPorciento, ValidezCuponDias
    FROM PromocionesConfig
    WHERE Activo = 1
      AND @Fecha BETWEEN FechaInicio AND FechaFin
      AND CHARINDEX(CAST(DATEPART(WEEKDAY,@Fecha)-1 AS VARCHAR), DiasPermitidos) > 0
      AND EXISTS (SELECT 1 FROM PromocionBodegas WHERE idPromocion = idPromocion AND idBodega = @idBodega)
END
```

#### 1.6 SP `ValidaSaldoCuponSegundaVisita` (NUEVO — extiende `ValidaSaldoReturnPass`)
```sql
-- Verifica que el cupón sea válido al momento del canje
CREATE PROCEDURE ValidaSaldoCuponSegundaVisita
    @idCliente INT, @idBodega INT, @idCategoria INT, @Fecha DATETIME
AS BEGIN
    SELECT c.idCupon, c.idAnticipo, a.Saldo, c.ValorDescuento, c.FechaVencimiento
    FROM CuponesSegundaVisita c
    INNER JOIN Anticipos a ON a.idAnticipo = c.idAnticipo
    WHERE c.idCliente = @idCliente
      AND c.Estado = 1                            -- Pendiente
      AND c.FechaVencimiento >= @Fecha            -- No vencido
      AND c.idBodega = @idBodega                  -- Mismo local
      AND c.idCategoria = @idCategoria            -- Misma categoría
      AND DATEPART(WEEKDAY,@Fecha)-1 NOT IN (5,6) -- No viernes ni sábado
END
```

#### 1.7 SP `GeneraCuponSegundaVisita` (NUEVO)
```sql
-- Crea el anticipo + el cupón al finalizar la venta
CREATE PROCEDURE GeneraCuponSegundaVisita
    @idCompra INT, @idCliente INT, @idHabitacion INT,
    @idCategoria INT, @idBodega INT, @ValorOriginal DECIMAL(18,2),
    @DescuentoPorciento DECIMAL(5,2), @ValidezDias INT
AS BEGIN
    -- 1. Calcular valores
    -- 2. INSERT Anticipos (Tipo=6, Estado=1, Saldo=ValorDescuento)
    -- 3. INSERT CuponesSegundaVisita
    -- 4. UPDATE Anticipos SET idOrigen = @idCompra
END
```

---

### 2. Pantalla `frmFacturacionHabitaciones.cs`

#### 2.1 Nuevo botón "SEGUNDA VISITA 50%" (NUEVO)
- Similar a `btnReturnPass` — ubicar junto a él en el toolbar
- Solo habilitado si `ValidaConfigPromoSegundaVisita` devuelve resultados (se valida en el Load/cambio de fecha)
- Al hacer clic: abre `frmAnticiposClientes` con `Tipo=6`, valor = 50% de la tarifa actual, vincula resultado a la compra

**Código referencia — actual `btnReturnPass_Click` en línea 11827:**
```csharp
// NUEVO: btnSegundaVisita_Click — misma estructura
using (frmAnticiposClientes Anticipos = new frmAnticiposClientes(
    true, sRUC, (int)this.cmbBodega.Value,
    dValorCupon,  // 50% de la tarifa actual en vez de $19 fijo
    6,            // Tipo = 6 (Segunda Visita) en vez de 5 (Return Pass)
    0, idTipoRuc, (int)this.cmbFormaPago.Value))
{
    if (DialogResult.OK == Anticipos.ShowDialog())
    {
        Funcion.EjecutaSQL(cdsSeteoF, string.Format(
            "Exec GeneraCuponSegundaVisita {0},{1},{2},{3},{4},{5},{6},{7}",
            idCompra, idCliente, idHabitacion, idCategoria,
            iBodega, dTarifaActual, 50, 30));
        ImpresionAutomaticaAnticipos((int)Anticipos.txtIdAnticipo.Value, 6, cdsSeteoF);
    }
}
```

#### 2.2 Validación al habilitar el botón (NUEVO — en Load y al cambiar fecha)
```csharp
// Verificar si la promo aplica hoy en esta bodega
int idPromoSV = Funcion.iEscalarSQL(cdsSeteoF,
    string.Format("Exec ValidaConfigPromoSegundaVisita {0}, '{1}'",
        iBodega, DateTime.Now.ToString("yyyyMMdd HH:mm")));
this.btnSegundaVisita.Enabled = (idPromoSV > 0);
```

#### 2.3 Extender `ValidaPromoReturnPass()` → nuevo método `ValidaPromoSegundaVisita()` (NUEVO)
- Se llama en los mismos puntos donde se llama `ValidaPromoReturnPass()` (líneas 2852, 2865, 8820)
- Usa `ValidaSaldoCuponSegundaVisita` en lugar de `ValidaSaldoReturnPass`
- Valida: vigencia, días, local, categoría, estado del cupón
- Si válido: aplica 50% de descuento al artículo 3422 en el grid
- Marca el cupón como `Estado=3` (Utilizado) en `CuponesSegundaVisita`

**Puntos donde se llama en el código actual:**
- `frmFacturacionHabitaciones.cs:2852` — al seleccionar habitación desde cuadrícula
- `frmFacturacionHabitaciones.cs:2865` — al confirmar selección
- `frmFacturacionHabitaciones.cs:8820` — al recargar anticipos

#### 2.4 Registrar uso al cobrar (NUEVO — en el bloque de cobro línea 6516)
```csharp
// Al confirmar el pago, si se usó cupón Segunda Visita:
if (idCuponSV > 0)
{
    sSQL = string.Format(
        "Insert Into AnticiposCompras (idCompra, idAnticipo, Valor, Estado) Values ({0},{1},{2},0) " +
        "Exec ActualizaSaldoFacturas {1}, 24, 6 " +
        "UPDATE CuponesSegundaVisita SET Estado=3, idCompraUso={0} WHERE idAnticipo={1}",
        idCompra, idAnticipoSV, dValorCuponSV);
    oCmdActualiza.CommandText = sSQL;
    oCmdActualiza.ExecuteNonQuery();
}
```

---

### 3. Reporte Crystal Reports — Cupón físico (NUEVO)

Archivo: `RptCuponSegundaVisita.rpt`

Debe imprimir:
- Nombre promoción y logo del local
- Fecha de emisión
- Fecha de vencimiento (emisión + 30 días)
- Categoría de habitación
- Local (Tantra Design / Gold / Loft / Platinum)
- Valor original de la habitación
- Valor con descuento (50%)
- Valor cancelado por el cupón
- Restricciones de uso (texto fijo)
- Días válidos: Domingo a Jueves
- Leyenda: "No canjeable por dinero · No transferible · Un solo uso"

---

### 4. Reportes de gestión (NUEVO)

| Reporte | Filtros |
|---|---|
| Cupones emitidos | Fecha, Local, Cajero, Categoría |
| Cupones utilizados | Fecha, Local, Cajero, Categoría |
| Cupones vencidos | Fecha, Local |
| Valor vendido vs redimido | Fecha, Local |
| Resumen por estado | Pendiente / Utilizado / Anulado / Vencido |

Fuente de datos: tabla `CuponesSegundaVisita` + `Anticipos` + `Compra`

---

## Flujo completo conectado al código

```
Cliente llega y elige habitación
         │
         ▼
frmFacturacionHabitaciones — Load
  → ValidaConfigPromoSegundaVisita(@idBodega, @hoy)
  → Si retorna promo activa: btnSegundaVisita.Enabled = true
         │
         ▼ (cajero hace clic en "SEGUNDA VISITA 50%")
btnSegundaVisita_Click [NUEVO]
  → Calcula dValorCupon = tarifa * 0.50
  → Abre frmAnticiposClientes(Tipo=6, monto=dValorCupon)
  → Cliente paga el cupón
  → GeneraCuponSegundaVisita SP [NUEVO]
      → INSERT Anticipos (Tipo=6, Saldo=dValorCupon)
      → INSERT CuponesSegundaVisita (categoría, local, vencimiento, Estado=1)
  → ImpresionAutomaticaAnticipos(idAnticipo, 6) → imprime RptCuponSegundaVisita.rpt
  → Factura única = habitación + cupón
         │
         ▼ (cliente regresa — nueva factura)
ValidaPromoSegundaVisita() [NUEVO — extiende el existente en línea 6127]
  → ValidaSaldoCuponSegundaVisita(@idCliente, @idBodega, @idCategoria, @hoy) [NUEVO SP]
      ✓ Estado = 1 (Pendiente)
      ✓ FechaVencimiento >= hoy
      ✓ Mismo local
      ✓ Misma categoría
      ✓ Día no es Viernes/Sábado
  → Si válido: pregunta al cajero "¿Aplicar cupón Segunda Visita?"
  → Si acepta: aplica 50% descuento al artículo 3422 en uGridDetalle
         │
         ▼ (cajero cobra — btnGuardar)
Bloque cobro línea 6516 [EXTENDER]
  → INSERT AnticiposCompras
  → ActualizaSaldoFacturas(@idAnticipo, 24, 6) → Anticipos.Saldo=0, Estado=3
  → UPDATE CuponesSegundaVisita SET Estado=3, idCompraUso=@idCompra
```

---

## Estados del cupón

| Código | Nombre | Cuándo se asigna |
|---|---|---|
| 1 | Pendiente | Al emitirlo (compra original) |
| 3 | Utilizado | Al canjear en segunda visita |
| 4 | Anulado | Anulación manual por supervisor |
| 5 | Vencido | Job nocturno o validación al consultar si FechaVencimiento < hoy |

---

## Checklist de desarrollo

### Base de datos
- [ ] Crear tabla `PromocionesConfig`
- [ ] Crear tabla `PromocionBodegas` (relación promo-local)
- [ ] Crear tabla `CuponesSegundaVisita`
- [ ] INSERT en `TiposDeAnticipos` con Tipo=6 y nombre del reporte
- [ ] Crear SP `ValidaConfigPromoSegundaVisita`
- [ ] Crear SP `ValidaSaldoCuponSegundaVisita`
- [ ] Crear SP `GeneraCuponSegundaVisita`
- [ ] Job o proceso para marcar cupones vencidos (Estado=5)

### Pantalla `frmFacturacionHabitaciones.cs`
- [ ] Declarar variables `idCuponSV`, `idAnticipoSV`, `dValorCuponSV`, `bUsoPromSV`
- [ ] Agregar botón `btnSegundaVisita` en el designer
- [ ] Validar promo activa en Load y al cambiar fecha (`ValidaConfigPromoSegundaVisita`)
- [ ] Implementar `btnSegundaVisita_Click`
- [ ] Implementar `ValidaPromoSegundaVisita()`
- [ ] Llamar `ValidaPromoSegundaVisita()` en los mismos 3 puntos donde está `ValidaPromoReturnPass()` (líneas 2852, 2865, 8820)
- [ ] Extender bloque de cobro (línea 6516) para registrar uso del cupón
- [ ] Extender `VentaFPagoLot` si necesita recibir el nuevo parámetro de cupón SV

### Reportes
- [ ] Diseñar `RptCuponSegundaVisita.rpt` en Crystal Reports
- [ ] Copiar .rpt al directorio de reportes del sistema
- [ ] Crear reporte de gestión de cupones (emitidos/utilizados/vencidos)

---

## Archivos del repositorio relacionados

| Archivo | Relevancia |
|---|---|
| `Latinium/frmFacturacionHabitaciones.cs` | Pantalla principal — aquí van todos los cambios de UI |
| `Latinium/SQL/SP_ReturnPass_ValidaSaldo.sql` | Base del nuevo SP de validación |
| `Latinium/SQL/SP_ActualizaSaldoFacturas.sql` | SP de descuento de saldo — se reutiliza con tipo 24 |

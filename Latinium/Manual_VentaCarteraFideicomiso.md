# VENTA DE CARTERA FIDEICOMISO
## Manual de Usuario — `frmFideicomisoProceso`

---

## REQUISITOS PREVIOS

| Requisito | Detalle |
|-----------|---------|
| **Acceso al módulo** | El usuario debe tener permiso en Seguridad (`id_seg = 2230`). Si no tiene acceso el sistema lo notificará al abrir la pantalla. |
| **Archivo de datos** | Un archivo Excel (`.xls`) o CSV (`.csv`) con las cédulas y números de operación a procesar. |
| **Entidad configurada** | Debe existir al menos una entidad financiera de tipo **FIDEICOMISO** activa en el sistema. |

---

## FLUJO GENERAL

```
[ CONFIGURAR ENTIDAD ] ──► [ PASO 1: CARGAR EXCEL ] ──► [ PASO 2: PREPARAR ] ──► [ PASO 3: EJECUTAR ]
```

> **IMPORTANTE:** Los pasos deben ejecutarse **en orden**. No se puede avanzar al Paso 3 si el Paso 2 tiene registros sin coincidencia.

---

## PANEL SUPERIOR — CONFIGURACIÓN INICIAL

Este panel siempre está visible en la parte superior de la pantalla.

### `[ Entidad ]` — Combo de Entidad Destino

- Seleccione la **entidad financiera** a la que se venderá la cartera.
- El combo solo muestra entidades de tipo **FIDEICOMISO** (`idCre_TipoCarte = 1`) que estén activas.
- Al seleccionar una entidad:
  - El título de la ventana cambia mostrando el nombre de la entidad.
  - El sistema calcula automáticamente el **Número de Lote** (`MAX(NumeroLote) + 1` de Compra para esa entidad).

### `[ Fecha Corte ]` — Solo lectura

- Este campo **no se edita manualmente**.
- El sistema lo calcula automáticamente en el Paso 2, tomando la **fecha más antigua** (`MIN(Compra.Fecha)`) de los registros encontrados.

### `[ Descargar Plantilla Excel ]` — Botón verde

- Genera y descarga una plantilla `.xls` con el formato correcto.
- Úsela cuando el área operativa necesite saber cómo preparar el archivo.

---

## `[ PASO 1 ]` — CARGAR EXCEL

> **Pestaña:** `1. Cargar Excel`

### Acciones disponibles

| Botón / Campo | Acción |
|---------------|--------|
| **`Buscar...`** | Abre el explorador para seleccionar el archivo Excel (`.xls`) o CSV (`.csv`). |
| **`Cargar Excel`** *(azul)* | Lee el archivo y muestra los datos en la grilla. |
| **`Buscar:` + campo de texto** | Filtra los registros de la grilla en tiempo real mientras escribe. |

### Estructura del archivo

El archivo debe tener datos **desde la fila 8**. Las primeras 7 filas son encabezados de la plantilla.

| Columna | Contenido | Observación |
|---------|-----------|-------------|
| **B** | Cédula | Máximo 10 caracteres. El sistema rellena con ceros a la izquierda. |
| **D** | Número de Operación | Máximo 9 caracteres. El sistema rellena con ceros a la izquierda. |
| **E** | Fecha de Operación | Se lee pero **no se usa** para filtrar. La fecha real viene de la tabla `Compra`. |

### Resultado esperado

- La grilla muestra todos los registros leídos.
- El contador **`Total leidos:`** indica cuántos registros se cargaron.
- El botón **`Preparar Datos`** (Paso 2) se habilita automáticamente.

> Si hay algún error en el archivo, el sistema mostrará un mensaje de error. Verifique el formato del archivo y vuelva a intentarlo.

---

## `[ PASO 2 ]` — PREPARAR Y VALIDAR

> **Pestaña:** `2. Preparar y Validar`

### Acción principal

| Botón | Descripción |
|-------|-------------|
| **`Preparar Datos`** *(verde)* | Procesa los registros del Excel contra la tabla `Compra` de la base de datos. |

### Qué hace el sistema internamente

1. **Limpia** la tabla temporal `temporalfideicomiso`.
2. **Inserta** todos los registros del Excel (cédula + número de operación).
3. **Cruza** contra `Compra` buscando registros que cumplan:
   - `idTipoFactura = 1`
   - `Borrar = 0`
   - `Estado = 10`
   - `idFormaPago = 9`
4. **Calcula la Fecha de Corte** = fecha más antigua de los registros encontrados.

### Sub-pestañas de resultado

#### `[ ENCONTRADOS ]` — Registros que SÍ cruzaron con Compra

| Columna | Descripción |
|---------|-------------|
| Cedula | Cédula del cliente |
| NumeroOperacion | Número de operación del Excel |
| NumeroCompra | Número interno en Compra |
| Entidad | Entidad financiera actual del crédito |
| EstadoSolNC | Estado de la solicitud |
| SaldoAnticipo | Saldo de anticipo |

> Use el campo **`Buscar:`** para localizar un registro específico por cédula, número de operación o entidad.

#### `[ SIN COINCIDENCIA ]` — Registros que NO cruzaron

| Columna | Descripción |
|---------|-------------|
| Cedula | Cédula del cliente |
| NumeroOperacion | Número de operación del Excel |
| Motivo | Razón por la que no se encontró |

**Motivos posibles:**

| Motivo | Causa |
|--------|-------|
| `No existe en Compra (cedula/numero no encontrado)` | La combinación cédula+número no existe en el sistema. |
| `Excluido: Borrar=1` | El registro está marcado para borrar. |
| `Excluido: TipoFactura=X` | El tipo de factura no es 1. |
| `Excluido: Estado=X` | El estado no es 10 (activo para cartera). |
| `Excluido: FormaPago=X` | La forma de pago no es 9 (crédito). |

> Use el campo **`Buscar:`** para filtrar por cédula, número de operación o motivo.

### Regla crítica

> **Si existen registros en SIN COINCIDENCIA, NO se puede continuar al Paso 3.**
>
> El sistema bloqueará la ejecución y mostrará un mensaje de error. Coordine con el área operativa para corregir el archivo y repita desde el Paso 1.

### Indicadores en pantalla

```
Total Excel: 150    Encontrados: 148 (verde)    Sin coincidencia: 2 (naranja-rojo)
```

---

## `[ PASO 3 ]` — EJECUTAR VENTA DE CARTERA

> **Pestaña:** `3. Ejecutar Venta`

> **Requisito:** Paso 2 completado y **sin registros en SIN COINCIDENCIA**.

### Acción principal

| Botón | Descripción |
|-------|-------------|
| **`Ejecutar Venta`** *(rojo oscuro)* | Aplica la venta de cartera a todos los registros encontrados en `Compra`. |

### Diálogo de confirmación

Antes de ejecutar, el sistema muestra un resumen:

```
¿Confirma la ejecución de la Venta de Cartera?

  Entidad destino : [Nombre de la entidad]
  Numero de lote  : [Número calculado automáticamente]

  Se actualizarán N registro(s) en la tabla Compra.

                  [ Sí ]   [ No ]
```

> Verifique que la entidad y el número de lote sean correctos antes de confirmar.

### Qué actualiza el sistema en `Compra`

| Campo | Valor asignado |
|-------|----------------|
| `Bloquear` | `1` |
| `CapturaLote` | `1` |
| `NumeroLote` | Número calculado automáticamente |
| `idEntidadFinanciera` | ID de la entidad seleccionada |
| `FechaTransitoLote` | Fecha y hora actual (`GETDATE()`) |
| `FechaVenta` | Fecha y hora actual (`GETDATE()`) |
| `NumeroDeOperacion` | `Ruc + '-' + Numero` |
| `Confirmado` | `1` |

### Sub-pestañas de resultado

#### `[ ACTUALIZADOS ]` — Registros procesados exitosamente

| Columna | Descripción |
|---------|-------------|
| Cedula | Cédula del cliente |
| NumeroOperacion | Número de operación |
| NumeroCompra | Número en Compra |
| NumeroLote | Lote asignado |
| Entidad | Entidad destino |
| NumeroDeOperacion | Número generado |
| FechaVenta | Fecha de la venta |

> Use el campo **`Buscar:`** para filtrar por cédula, número de operación, entidad o número de operación generado.

#### `[ SIN COINCIDENCIA ]` — Fallidos (si los hay)

Registros que por alguna razón no pudieron procesarse. Bajo condiciones normales esta pestaña debería estar vacía.

### Indicadores finales

```
Total: 150    Actualizados: 150 (verde)    Sin coincidencia: 0 (naranja-rojo)
```

---

## RESUMEN DEL PROCESO

```
┌─────────────────────────────────────────────────────────────────┐
│                    PROCESO PASO A PASO                          │
├─────┬───────────────────────────────────────────────────────────┤
│  0  │ Abrir pantalla → Seleccionar ENTIDAD DESTINO              │
├─────┼───────────────────────────────────────────────────────────┤
│  1  │ Tab "1. Cargar Excel"                                     │
│     │   → Buscar archivo → Cargar Excel                         │
│     │   → Verificar total de registros leídos                   │
├─────┼───────────────────────────────────────────────────────────┤
│  2  │ Tab "2. Preparar y Validar"                               │
│     │   → Click "Preparar Datos"                                │
│     │   → Revisar ENCONTRADOS ✔                                 │
│     │   → Verificar que SIN COINCIDENCIA = 0 ✔                 │
│     │   → Si hay sin coincidencia: CORREGIR y volver al Paso 1 │
├─────┼───────────────────────────────────────────────────────────┤
│  3  │ Tab "3. Ejecutar Venta"                                   │
│     │   → Click "Ejecutar Venta"                                │
│     │   → Confirmar en el diálogo                               │
│     │   → Revisar ACTUALIZADOS                                  │
└─────┴───────────────────────────────────────────────────────────┘
```

---

## PREGUNTAS FRECUENTES

**¿Se puede repetir el proceso con el mismo archivo?**
Sí. Cada vez que se ejecuta el Paso 1 (Cargar Excel) se limpia la tabla temporal y se puede volver a procesar. Sin embargo, si la venta ya fue ejecutada (Paso 3), los registros en `Compra` ya fueron modificados y no se debe ejecutar nuevamente sobre los mismos registros.

**¿Por qué no aparece mi entidad en el combo?**
El combo solo muestra entidades con `Activo = 1` y `idCre_TipoCarte = 1` (FIDEICOMISO). Si la entidad no aparece, verifique su configuración en la tabla `Cre_EntidadFinanciera`.

**¿Por qué la Fecha de Corte no corresponde a la del Excel?**
Por diseño del sistema, la Fecha de Corte **no proviene del Excel** sino de la fecha más antigua registrada en `Compra` para los registros encontrados. Esto garantiza que la fecha sea confiable y no dependa de errores del archivo.

**¿Qué pasa si el número de lote calculado es incorrecto?**
El número de lote se calcula como `MAX(NumeroLote) + 1` de `Compra` filtrando por `idTipoFactura = 1` y la entidad seleccionada. Si necesita un lote específico, verifique con el área técnica.

---

*Módulo: Fideicomiso | Seguridad: `id_seg = 2230` | Código acceso: `08320702`*

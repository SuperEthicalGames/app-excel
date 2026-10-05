# Arquitectura y diccionario de datos

## Estado de las fases

| Fase | Estado |
|---|---|
| 1. Análisis | Diseño y alcance documentados |
| 2. Modelo de datos | Tablas y relaciones implementadas |
| 3. Núcleo Excel | Plantilla XLSX generable desde `build/build_workbook.py` |
| 4. Automatización VBA e interfaz | 11 módulos fuente; página INICIO y asistente local de clasificación incluidos; compilación VBA pendiente |
| 5. Documentos | Importación y lectura UBL implementadas en fuente; OCR no configurado |
| 6. QA | Libro y reportes verificados en Excel con 5.000 movimientos sintéticos; compilación y ejecución VBA pendientes |
| 7. Optimización | Rangos centralizados y duplicados indexados; recálculo de 5.000 movimientos en 0,33 s en Excel 16.0 |
| 8. Entrega | XLSX, 11 módulos fuente, arquitectura, manuales, mantenimiento, QA y roadmap documentados |

## 1. Capas y relaciones

```text
CONFIGURACIÓN                     DATOS                         REPORTES
EMPRESA                           TERCEROS ──┐                   INICIO
PARÁMETROS                        DOCUMENTOS ─┴─► MOVIMIENTOS    INGRESOS / GASTOS
CATEGORÍAS                                     │                 FLUJO_CAJA
IMPUESTOS                                      ├─► PAGOS          CXC / CXP
CUENTAS                                        └─► COBROS         IMPUESTOS_REPORTE
TIPOS_DOCUMENTOS                                                 REPORTE_MENSUAL / ANUAL
ESTADOS
MEDIOS_PAGO
USUARIOS_ROLES  ASISTENTE ── REGLAS_IA  LOG (auditoría VBA)
```

`DOCUMENTOS` conserva datos extraídos del soporte; `MOVIMIENTOS` representa la operación financiera. El campo `SoporteID` relaciona ambos. `PAGOS` y `COBROS` guardan desembolsos o recaudos aplicados, incluso parciales.

## 2. Diccionario de tablas

### Configuración

| Hoja/tabla | Campos principales | Uso |
|---|---|---|
| EMPRESA | Campo, Valor | Identidad, NIT, moneda, actividad y rutas |
| PARAMETROS | Código, Parámetro, Valor, Tipo, Descripción | Prefijos, tolerancias, rutas y opciones |
| CATEGORIAS | Código, Tipo, Categoría, Subcategoría, Activa, Descripción | Clasificación editable |
| IMPUESTOS | Código, Impuesto, Porcentaje, Tipo, Base, CuentaAsociada, Activo, Observaciones | Tasas configurables; revisar normativa aplicable |
| CUENTAS | ID, Cuenta, Tipo, Entidad, NúmeroReferencia, SaldoInicial, SaldoActual, Activa, Observaciones | Cuentas y saldos |
| TIPOS_DOCUMENTOS | Código, TipoDocumento, Prefijo, RequiereCUFE, Activo, Descripción | Catálogo editable |
| ESTADOS | Código, Estado, Ámbito, Activo, Descripción | Catálogo de referencia |
| MEDIOS_PAGO | Código, MedioPago, Activo, Descripción | Catálogo editable |
| USUARIOS | ID, Usuario, Rol, Email, Activo, Observaciones | Referencia organizativa; no es control de acceso |
| REGLAS_IA | ID, Activa, PalabraClave, Tipo, Categoria, Subcategoria, MedioPago, Cuenta, TerceroOpcional, Prioridad, Coincidencia, Observaciones | Patrones editables para sugerir clasificación; prioridad 1–100; empate produce conflicto visible |

### Datos y sistema

| Hoja/tabla | Campos clave | Relaciones y reglas |
|---|---|---|
| TERCEROS | ID, Tipo, NombreRazonSocial, NIT_ID, contacto, ubicación, Estado, Observaciones | El VBA marca NIT/ID exactamente repetido después de quitar espacios y normalizar mayúsculas; no borra filas |
| DOCUMENTOS | ID, Fecha, TipoDocumento, Numero, Tercero, Identificacion, Subtotal, Impuestos, Retenciones, Total, CUFE_Referencia, Estado, RutaArchivo, NombreArchivo, TipoArchivo, FechaExtraccion, MetodoExtraccion, Confianza, Observaciones | `Total` conserva el valor final del soporte, incluido PayableAmount extraído |
| MOVIMIENTOS | ID, Fecha, Tipo, Documento, Tercero, NIT_ID, Concepto, Categoria, Subcategoria, MedioPago, Cuenta, Subtotal, Impuestos, Retenciones, Total, Estado, SoporteID, Observaciones | `Total` se calcula como Subtotal + Impuestos - Retenciones |
| PAGOS | ID, Fecha, MovimientoID, Tercero, Cuenta, MedioPago, ValorPagado, Estado, Observaciones | Salidas parciales contra movimiento |
| COBROS | ID, Fecha, MovimientoID, Tercero, Cuenta, MedioPago, ValorCobrado, Estado, Observaciones | Entradas parciales contra movimiento |
| LOG | Fecha, Hora, Usuario, Acción, Módulo, RegistroAfectado, ValorAnterior, ValorNuevo, Resultado | Auditoría escrita por VBA |

`INICIO` es la página de navegación y puesta en marcha. `ASISTENTE` recibe un tercero y un concepto, y propone campos mediante coincidencias de texto. `REGLAS_IA` conserva hasta 100 reglas editables. Si dos reglas de máxima prioridad coinciden, el asistente pide ajustar la prioridad en vez de escoger silenciosamente.

El total de DOCUMENTOS se guarda como dato porque el valor facturado puede incluir ajustes que no equivalen a una suma simple. El total de MOVIMIENTOS es calculado por fórmula. La conciliación del XML informa diferencias y requiere revisión humana.

## 3. Compatibilidad y fórmulas

- El libro base usa fórmulas disponibles en Excel 2019 y tablas para captura, filtrado y expansión.
- El asistente local utiliza `IF`, `MAX`, `COUNTIF`, `INDEX`, `MATCH`, `SEARCH` e `IFERROR`, disponibles en Excel 2019. No usa una API ni transmite datos. Es un clasificador por reglas, no un modelo generativo, no aprende automáticamente, no determina impuestos y no crea movimientos.
- Las fórmulas de reporte usan rangos absolutos generados desde `SHEET_LAYOUT`, no referencias estructuradas.
- El libro solicita cálculo automático y recálculo completo al abrirse para refrescar resultados en Excel.
- Las tablas comienzan con un área breve de captura; fórmulas, desplegables y reportes analizan hasta la capacidad central definida en `ANALYTIC_ROW_CAPACITY`.
- La capacidad actual es 10.000 filas analíticas para DOCUMENTOS, MOVIMIENTOS, PAGOS y COBROS; 2.000 para TERCEROS y 20.000 para LOG. Para ampliarla se ajusta el mapa y se reconstruye el libro.
- Los reportes mensuales usan el año editable en B3. El reporte anual presenta 12 años desde el año inicial editable en B3.
- El cálculo y los filtros se verificaron en Excel 16.0 con 5.000 movimientos. No se ejecutó una instalación independiente de Excel 2019; las fórmulas usadas pertenecen al conjunto común de Excel 2019.
- El archivo de entrega es XLSX y no contiene los módulos VBA. Para habilitar macros, se guarda una copia como XLSM y se importan los archivos .bas.

## 4. Automatización y documentos

El VBA usa módulos separados: `modMain`, `modFiles`, `modXmlImport`, `modDocumentos`, `modMovimientos`, `modTerceros`, `modDuplicados`, `modValidation`, `modReportes`, `modPagosCobros` y `modUtils`.

El importador reconoce UBL `Invoice`, `CreditNote` y `DebitNote` mediante MSXML 6 y XPath por nombre local. Compara moneda con EMPRESA.MonedaPrincipal, guarda campos encontrados y deja la fila en `Revision Manual`. No valida firma, autenticidad, CUFE ante DIAN ni cumplimiento tributario. PDF e imágenes quedan pendientes de OCR externo o digitación manual.

La validación masiva de documentos y terceros construye índices de duplicidad con `Scripting.Dictionary` y late binding para reducir el trabajo de O(n²) a O(n). En documentos compara CUFE, NIT+Número y NIT+Fecha+Total; en terceros compara NIT/ID exacto. Los nombres similares y la combinación nombre+número+total no se puntúan. No elimina registros.

## 5. Límites y extensiones

- `PARAMETROS.RUTA_DOCUMENTOS` controla dónde se copian los soportes. Las rutas relativas parten de la carpeta del libro; una ruta absoluta también es válida. `EMPRESA.RutaRaizArchivos` queda como dato editable reservado y no la consume el VBA actual.
- Las hojas no se protegen por defecto para permitir la expansión de las tablas. Los roles de USUARIOS son informativos, no autentican ni restringen acciones. El LOG solo registra las acciones que ejecuta el VBA.

- Los rangos analíticos terminan en la capacidad indicada; si una empresa supera esos límites, hay que aumentar el mapa y reconstruir el XLSX.
- Los reportes de ingresos y gastos presentan las categorías iniciales de la plantilla; nuevas categorías requieren ampliar el diseño del reporte.
- Las tasas e impuestos no son asesoría tributaria. Se deben configurar y verificar contra las reglas oficiales vigentes.
- OCR, APIs, DIAN, bancos, inventario, nómina y TRA no están conectados.
- La TRA debe añadirse como módulo independiente y no introducir datos de alojamiento en el núcleo financiero.

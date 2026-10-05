# QA de fases 6 y 7

## Entorno y alcance

Libro XLSX generado para la plantilla 2026. Las verificaciones de fórmulas y carga se ejecutaron en Excel Desktop mediante COM, versión informada `16.0`. La carga de 5.000 movimientos fue sintética y se cerró sin guardar; no se dejaron filas de prueba en el archivo de entrega.

## Casos ejecutados

| Caso | Entrada | Resultado esperado | Resultado obtenido | Estado | Observaciones |
|---|---|---|---|---|---|
| Abrir el XLSX | Libro principal | Excel lo abre sin reparación | Abierto en Excel 16.0 | Ejecutado | No se guardó desde Excel |
| Recálculo de plantilla vacía | Libro sin movimientos | Sin errores de fórmula | 0 errores | Ejecutado | `CalculateFullRebuild` |
| Expansión de tabla | 5.000 movimientos insertados en memoria | Tabla MOVIMIENTOS se expande | Rango A4:R5004 | Ejecutado | Cerrado sin guardar |
| Indicadores pagados | 2.500 ventas/ingresos y 1.000 gastos/compras; subtotal, impuesto y retención definidos | Ingresos 297.500; gastos 47.600; flujo 249.900 | Coinciden | Ejecutado | Importes incluyen impuestos según Total |
| Cuentas pendientes | 500 ingresos y 1.000 compras pendientes | CxC 71.400; CxP 89.250 | Coinciden | Ejecutado | Total calculado desde subtotal+impuestos-retenciones |
| Impuestos pagados | Impuestos de filas pagadas | 55.100 | 55.100 | Ejecutado | Suma de la columna Impuestos |
| Reporte mensual | Año 2026, septiembre | Ingresos 297.500; gastos 47.600; resultado 249.900 | Coinciden | Ejecutado | B3 define el año |
| Reporte anual | Ventana que incluye 2026 | Ingresos 297.500; gastos 47.600; resultado 249.900 | Coinciden | Ejecutado | 2026 quedó en la fila correspondiente al año |
| Filtro Tipo | Tipo Ingreso | Ingresos pagados 148.750 | Coincide | Ejecutado | Considera solo filas pagadas con Tipo exacto Ingreso |
| Filtro Categoría | Categoría Ventas | Ingresos 297.500 | Coincide | Ejecutado | |
| Filtro Estado | Estado Pendiente | Ingresos 71.400 | Coincide | Ejecutado | |
| Filtro Año | Año 2025 sin datos | 0 | 0 | Ejecutado | |
| Filtro Mes | Agosto sin datos | 0 | 0 | Ejecutado | |
| Filtro Tercero | Nombre sin coincidencia | 0 | 0 | Ejecutado | |
| Filtro Cuenta | Cuenta Caja | Ingresos 297.500 | Coincide | Ejecutado | |
| Validación de Año | B5 | Validación de número entero | Tipo de validación 1 en Excel | Ejecutado | |
| Validación de Tipo | B7 | Lista de tipos | Tipo de validación 3 en Excel | Ejecutado | |
| Errores de fórmula con 5.000 filas | Carga sintética y reportes recalculados | 0 errores | 0 errores | Ejecutado | |
| Tiempo de recálculo | `CalculateFullRebuild` con 5.000 filas | Medición registrada | 0,33 s | Ejecutado | Tiempo observado en esta máquina, no una garantía |
| Compilación de módulos VBA | 11 archivos `.bas` | Proyecto VBA compila | No ejecutado | Pendiente | Excel no expuso VBProject/VBE a COM |
| Pruebas funcionales de macros | Importación, validación, LOG, renombrado, CXC/CXP | Resultado esperado por macro | No ejecutado desde esta sesión | Pendiente | Seguir `02_Manual_Instalacion_VBA.md` |
| Excel 2019 independiente | Abrir, editar y recalcular | Sin errores | No ejecutado | Pendiente | Excel instalado informó versión 16.0 |
| OCR de PDF e imágenes | PDF, JPG, PNG | Extraer y validar campos | No implementado | Fuera de alcance actual | No hay proveedor OCR configurado |
| Abrir interfaz | Libro final con portada | Excel abre el libro y muestra INICIO primero | Abierto en Excel 16.0; 28 hojas; INICIO activo | Ejecutado | Hipervínculos internos nativos |
| Navegación principal | Tarjetas de INICIO y enlaces del asistente | Cada vínculo abre su módulo | Verificado por destinos y referencias internas | Ejecutado | Configuración, movimientos, asistente, documentos, dashboard, reportes y reglas |
| Desplegables de reglas | REGLAS_IA | Activa, Tipo, Categoría, MedioPago, Cuenta y Prioridad tienen validación | Listas y prioridad 1–100 presentes | Ejecutado | Tabla editable A4:L104 |
| Asistente sin entrada | B5:B6 vacíos | Pedir un concepto | Coincide | Ejecutado | Sin falso resultado |
| Sugerencia de clasificación | Concepto "Pago de arriendo local" | Gasto / Administracion / Arriendo, prioridad 20 | Coincide | Ejecutado | No crea movimientos |
| Regla personalizada | Concepto "Pago de parqueadero" y regla temporal prioridad 35 | Sugiere Gasto / Administracion / Arriendo, palabra y prioridad | Coincide | Ejecutado | La regla y los datos de prueba se retiraron antes del guardado final |
| Conflicto de reglas | "Venta de productos y venta de servicios" | Mostrar conflicto por prioridades iguales y ocultar clasificación | Coincide | Ejecutado | Revisión humana requerida |
| Sin coincidencia | Concepto sin palabra configurada | Mostrar "Sin coincidencia" | Coincide | Ejecutado | Registro manual permitido |
| Integridad del libro con asistente | Recálculo completo en Excel 16.0 | Sin errores de fórmula | 0 errores | Ejecutado | Entradas de prueba vacías al guardar |

## Resultado

Los cálculos, filtros, reportes, validaciones del dashboard, expansión de la tabla y asistente local pasaron en Excel 16.0. El recálculo completo de las fórmulas del libro con 5.000 registros tomó 0,33 s y no produjo errores. La carga se hizo en memoria y no se guardó. Las pruebas del asistente también se limpiaron antes de guardar.

La aprobación QA no cubre los once módulos VBA: el editor no quedó accesible a la automatización y el XLSX no contiene proyectos VBA. Para cerrar esa parte, guarde una copia como XLSM, importe los módulos, compile y ejecute los casos descritos en `02_Manual_Instalacion_VBA.md`.


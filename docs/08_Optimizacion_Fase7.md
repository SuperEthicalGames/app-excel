# Optimización de fase 7

## Cambios aplicados

- Centralización de los límites de análisis en `ANALYTIC_ROW_CAPACITY`.
- Tablas cortas al inicio y rangos de análisis limitados, para que el XLSX no cree miles de filas de captura vacías.
- Fórmulas de reportes, filtros y validaciones extendidas a una capacidad de 10.000 filas en MOVIMIENTOS, DOCUMENTOS, PAGOS y COBROS; 2.000 en TERCEROS y 20.000 en LOG.
- Fórmulas de Total de MOVIMIENTOS copiables al expandir la tabla. El Total de DOCUMENTOS queda como valor importado o digitado.
- Detección masiva de duplicados indexada mediante `Scripting.Dictionary` con late binding. El recorrido pasa de comparaciones repetidas O(n²) a construcción y consulta de índices O(n), con memoria proporcional al número de filas.
- `SiguienteID` recorre los IDs en memoria, en lugar de consultar celdas una por una.
- `ActualizarTodo` calcula una vez, actualiza CXC/CXP y restaura `ScreenUpdating` aunque se produzca un error.

## Medición en Excel

| Medida | Resultado |
|---|---|
| Versión reportada por Excel | 16.0 |
| Filas sintéticas cargadas | 5.000 |
| Rango de tabla tras la carga | A4:R5004 |
| Recálculo completo observado | 0,33 segundos |
| Errores de fórmula | 0 |
| Guardado de datos sintéticos | No; el libro se cerró descartando cambios |

La medición cubre recálculo de fórmulas, dashboard y reportes en esta máquina. No mide OCR, macros, apertura en red, tiempo de entrada de datos ni uso concurrente. El tiempo no es una promesa de rendimiento para otros equipos.

## Escalado y mantenimiento

Las hojas de reporte vuelven a evaluar rangos acotados hasta el límite analítico en cada recálculo. Si el volumen excede habitualmente los límites o existen varios usuarios simultáneos, conviene mover la persistencia a una base de datos y usar Excel como interfaz o capa de análisis. Power Query, Power BI y las integraciones externas quedan como extensiones posteriores.


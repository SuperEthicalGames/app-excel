# Mantenimiento y roadmap

## Archivos fuente

| Componente | Ubicación | Función |
|---|---|---|
| Generador del libro | `build/build_workbook.py` | Reconstruye hojas, tablas, fórmulas, validaciones, formatos y gráficos |
| Interfaz y asistente | `INICIO`, `ASISTENTE`, `REGLAS_IA` en el libro; funciones equivalentes en el generador | Navegación guiada y clasificación local configurable |
| Fuentes VBA | `vba/*.bas` | Automatizaciones para importar soportes, validar, registrar y actualizar reportes |
| Verificación de interfaz y asistente | `build/verify_assistant_excel.ps1` | Abre el XLSX en Excel, comprueba navegación y reglas, y limpia los datos de prueba |
| Fixtures XML | `tests/fixtures/` | Ejemplos sintéticos válidos y malformados para revisión técnica |
| XLSX base | `dist/Sistema_Financiero_Empresa_2026.xlsx` | Libro sin macros para captura, configuración y reportes |
| Arquitectura y datos | `docs/01_Arquitectura_y_Diccionario_Datos.md` | Capas, relaciones, campos y límites |

El generador requiere Python y `openpyxl`. Instale la versión fijada con `pip install -r requirements.txt`. Ejecútelo desde la raíz del proyecto y especifique una ruta de salida distinta si quiere conservar una versión anterior. Abra el resultado en Excel y ejecute el recálculo completo antes de distribuirlo.

## Cambios al esquema

Al agregar, quitar o renombrar campos, revise cada referencia correspondiente en:

1. La tabla y el diccionario en `build/build_workbook.py`.
2. El mapa `ANALYTIC_ROW_CAPACITY`, rangos con nombre, fórmulas, validaciones y formatos condicionales.
3. Los accesos `ListColumns("...")` de los módulos VBA.
4. Las reglas de validación, documentación del flujo y casos QA.
5. La copia XLSM de trabajo: retire el módulo importado anterior e importe la versión actualizada; después use Depuración > Compilar VBAProject.

Las fórmulas analíticas usan rangos limitados para conservar compatibilidad con Excel 2019 y evitar cálculos de columna completa. La capacidad se amplía en un solo lugar, `ANALYTIC_ROW_CAPACITY`, y se reconstruye el libro. Las tablas comienzan con pocas filas y se expanden al capturar datos.

## Mantenimiento de VBA

- Mantenga `Option Explicit`, manejo de errores y referencias tardías donde sea viable.
- `modDuplicados` usa `Scripting.Dictionary` con late binding y conserva los primeros IDs detectados para señalar ambas filas.
- La detección actual compara CUFE, NIT+Número y NIT+Fecha+Total en documentos; en terceros compara NIT/ID exacto normalizado. No calcula similitud aproximada de nombres ni la combinación nombre+número+total.
- La validación no elimina documentos ni terceros. Los estados y las notas de observación quedan disponibles para revisión humana.
- La lectura UBL requiere MSXML 6.0, instalado en Windows. El parser deshabilita DTD y resolución de entidades externas.
- `PARAMETROS.RUTA_DOCUMENTOS` controla la carpeta de importación. Las rutas relativas parten de la carpeta del libro; las absolutas pueden apuntar a una ubicación compartida.
- El clasificador consulta hasta 100 filas de `REGLAS_IA`. Si se amplía ese límite, actualice el mapa `ANALYTIC_ROW_CAPACITY` y los rangos en las fórmulas de `ASISTENTE`, luego regenere y verifique el libro. Las sugerencias son deterministas, locales y requieren revisión humana.
- Haga un respaldo del XLSM y de los soportes antes de cambiar módulos o mover la carpeta documental.

## Límites conocidos

- El archivo entregado es XLSX; no contiene VBA. La compilación de los módulos y las pruebas funcionales de macros quedan pendientes de ejecutarse en Excel Desktop desde el editor VBA.
- La sesión pudo abrir y recalcular XLSX con Excel 16.0, pero no expuso `VBProject/VBE`. No se activó globalmente la opción de confianza de Office.
- No se hizo una prueba separada en Excel 2019. Las fórmulas del núcleo son funciones clásicas disponibles en esa versión.
- No hay OCR para PDF, JPG, PNG o fotos. Los PDF e imágenes se copian y se completan manualmente.
- No hay validación XSD, autenticación DIAN, consulta CUFE, integración bancaria, API ni conciliación automática.
- Los reportes de ingresos/gastos mantienen el catálogo inicial de categorías. Las categorías nuevas pueden registrarse, pero ampliar su presentación en reportes requiere modificar el generador.
- Las tasas tributarias deben configurarse y contrastarse con la normativa oficial vigente. El libro no determina obligaciones tributarias.

## Roadmap funcional

| Etapa | Alcance propuesto | Condición de entrada |
|---|---|---|
| 1. Sistema financiero base | Configuración, terceros, movimientos, documentos, pagos, cobros y reportes | Núcleo actual validado y esquema aprobado por la empresa |
| 2. Digitalización documental | Flujo de importación, clasificación, revisión y vínculo con movimientos | Compilación VBA y políticas de almacenamiento definidas |
| 3. OCR | Extracción de PDF e imágenes con bandeja de revisión y trazabilidad de proveedor | Elegir proveedor, presupuesto, retención y controles de datos |
| 4. IA avanzada | Evaluar normalización y sugerencias con modelos locales o servicios externos | Presupuesto, privacidad, métricas de error y revisión humana definidos; el asistente por reglas local ya está disponible |
| 5. Integraciones | DIAN, bancos, facturación, Power BI o base externa mediante adaptadores | API, autorizaciones, seguridad, costos y reglas vigentes verificados |
| 6. TRA | Módulo independiente de registro de alojamiento | Especificación oficial vigente y proceso hotelero confirmados |

La TRA se mantiene fuera del modelo financiero central. Las siguientes integraciones se implementarán como módulos que leen o escriben mediante interfaces documentadas, sin agregar campos de alojamiento a `MOVIMIENTOS` ni a `DOCUMENTOS`.


# QA de la Fase 5: documentos

## Alcance y criterio

Este registro separa los controles estructurales y las pruebas de entorno de las pruebas funcionales del VBA. Abrir el libro con Excel o cargar XML con MSXML no demuestra que los módulos compilen ni que el procedimiento VBA extraiga los campos correctamente.

## Casos y resultados

| Caso | Entrada | Resultado esperado | Resultado obtenido | Estado | Observaciones |
|---|---|---|---|---|---|
| Estructura de módulos VBA | Los 11 archivos .bas | Bloques Sub, Function, If, For y With balanceados | El análisis estático encontró aperturas y cierres balanceados | Ejecutado | No equivale a compilar |
| Estructura de DOCUMENTOS | Libro XLSX de entrega | Están las columnas usadas por el parser | Se inspeccionó la tabla; están presentes las 19 columnas requeridas | Ejecutado | Inspección con openpyxl |
| Abrir libro en Excel | Sistema_Financiero_Empresa_2026.xlsx | Excel abre el libro sin error | Excel COM 16.0 lo abrió en modo oculto | Ejecutado | No se guardó sobre el original |
| Carga de MSXML 6 y XPath | tests/fixtures/ubl_invoice_valid.xml | Carga XML y encuentra el nodo Invoice | MSXML 6 cargó el archivo y XPath encontró un Invoice | Ejecutado | Prueba de entorno, no ejecuta modXmlImport |
| Fixture UBL válida | tests/fixtures/ubl_invoice_valid.xml | XML bien formado y campos/valores de ejemplo coherentes | ElementTree leyó factura, fecha, moneda, NIT/DV e importes; la aritmética de ejemplo concilia | Ejecutado parcialmente | No prueba el código VBA |
| Fixture XML mal formada | tests/fixtures/ubl_malformed.xml | El parser rechaza el XML | ElementTree detectó el error de parseo | Ejecutado parcialmente | No prueba el estado Error de VBA |
| Acceso al proyecto VBA | Excel COM y libro de prueba | Poder importar módulos para compilarlos y ejecutarlos | Excel abrió el libro, pero el objeto VBProject/VBE no quedó accesible desde esta sesión | No disponible | No se modificó la configuración global de confianza de Office |
| Compilación de los 11 módulos | Archivos .bas en Excel Desktop | Compilación sin errores y macros visibles | No ejecutado | Pendiente | Requiere acceso al proyecto VBA en Excel |
| Extracción de factura UBL | Fixture válida desde modXmlImport | Guardar tipo, fecha, tercero, NIT, subtotal, impuestos, retenciones, total y CUFE; estado Revision Manual | No ejecutado desde VBA | Pendiente | Incluye regresión para el mapeo de PayableAmount a Total |
| XML mal formado desde VBA | Fixture mal formada | Estado Error con detalle del parser | No ejecutado desde VBA | Pendiente | |
| Nota crédito, nota débito y prefijos alternativos | XML UBL de cada tipo | Tipo correcto y nodos hallados por local-name() | No ejecutado desde VBA | Pendiente | |
| Moneda ausente o distinta | XML con moneda no coincidente con EMPRESA | No escribir importes y dejar aviso | No ejecutado desde VBA | Pendiente | |
| Campos incompletos y tercero ambiguo | XML sin fecha, UUID, tercero o NIT coincidente | Conservar campos disponibles y añadir avisos | No ejecutado desde VBA | Pendiente | |
| Total fuera de tolerancia | PayableAmount que no concilia | Guardar Total y añadir DIFERENCIA_TOTAL | No ejecutado desde VBA | Pendiente | Conciliación informativa; exige revisión humana |
| CUFE duplicado | Dos filas con la misma referencia | Marcar posible duplicado sin borrar datos | No ejecutado desde VBA | Pendiente | Ejecutar Menu_ValidarTodo |
| Importar PDF o imagen | PDF, JPG o PNG | Copiar soporte y dejar fila pendiente, sin afirmar OCR | No ejecutado desde VBA | Pendiente | No hay proveedor OCR configurado |
| Relectura y confirmación humana | Fila con datos previos y marca XML_REVISION | Pedir confirmación antes de sobrescribir y conservar revisión hasta confirmar | No ejecutado desde VBA | Pendiente | |

## Resultado de QA

Los controles estáticos, la inspección de la tabla y las pruebas independientes de XML pasaron. También se comprobó que Excel y MSXML 6 están disponibles. La integración VBA todavía no tiene aprobación funcional porque no se pudo acceder al proyecto VBA para importar, compilar y ejecutar los módulos desde esta sesión.

Las fixtures son sintéticas y no representan una factura emitida ni validada por la DIAN. Los módulos .bas permanecen separados del libro .xlsx; todavía no se ha generado un libro macrohabilitado.

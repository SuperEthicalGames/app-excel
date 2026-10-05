# Fases 4 y 5: instalación de módulos VBA

## Guardar el archivo habilitado para macros

1. Abra dist/Sistema_Financiero_Empresa_2026.xlsx.
2. Use Guardar como y elija Libro de Excel habilitado para macros (.xlsm).
3. Presione Alt+F11 para abrir el editor de VBA.
4. En Archivo, elija Importar archivo.
5. Importe los once módulos de la carpeta vba, uno por uno:
   - modUtils.bas
   - modValidation.bas
   - modDuplicados.bas
   - modMovimientos.bas
   - modDocumentos.bas
   - modTerceros.bas
   - modPagosCobros.bas
   - modFiles.bas
   - modReportes.bas
   - modXmlImport.bas
   - modMain.bas
6. En el editor VBA, use Depuración > Compilar VBAProject. Si aparece un error, corrija el módulo y vuelva a compilar.
7. Cierre el editor con Alt+Q y guarde con Ctrl+S.
8. Abra Alt+F8 y ejecute Menu_ImportarDocumento con tests/fixtures/ubl_invoice_valid.xml. Compruebe Numero, Fecha, Tercero, Subtotal, Impuestos y Total, y que Estado quede en Revision Manual.
9. Repita con tests/fixtures/ubl_malformed.xml; debe quedar Estado=Error y mostrar el detalle del parser.

## Configurar la carpeta de soportes

En `PARAMETROS`, el valor `RUTA_DOCUMENTOS` define la carpeta usada por la importación. El valor inicial `Documentos` se interpreta relativo a la carpeta del libro. También puede configurarse una ruta absoluta local o de red. Si el libro aún no se ha guardado, las rutas relativas parten de la carpeta de trabajo actual de Excel.

El libro contiene fórmulas, validaciones y reportes con capacidad analítica inicial de 10.000 filas para MOVIMIENTOS, DOCUMENTOS, PAGOS y COBROS; 2.000 para TERCEROS y 20.000 para LOG. Para cambiar esos límites, edite `ANALYTIC_ROW_CAPACITY` en `build/build_workbook.py` y vuelva a generar el XLSX antes de capturar datos.

La tabla se expande al agregar filas. No escriba sobre la columna `Total` de MOVIMIENTOS porque se calcula desde Subtotal, Impuestos y Retenciones. El campo `Total` de DOCUMENTOS sí es un dato capturado o extraído del soporte.

La fixture UBL válida es sintética y solo sirve para comprobar el recorrido técnico. No es una factura real ni válida para soportar una operación. La importación de las macros no se ha compilado ni ejecutado desde esta sesión porque Excel no expuso el objeto VBProject/VBE a la automatización.

No es necesario activar Trust access to the VBA project object model para importar módulos manualmente.

## Acciones para documentos

- Menu_ImportarDocumento: copia el soporte a la ruta configurada y crea un registro en DOCUMENTOS. Para XML UBL intenta extraer campos estructurados. PDF, imágenes y CSV quedan pendientes para revisión manual.
- Menu_ExtraerXMLDocumentoActivo: relee el XML de la fila seleccionada. Si ya hay número, CUFE o subtotal, pide confirmación antes de sustituir los campos presentes en el XML.
- Menu_ConfirmarRevisionXMLDocumentoActivo: después de comparar los campos con el soporte y resolver avisos, confirma la revisión humana. Luego ejecute Menu_ValidarTodo.
- Menu_RenombrarArchivoDocumentoActivo: renombra el soporte seleccionado por fecha, tercero, tipo y número. No sobrescribe un archivo destino existente.
- Menu_ValidarTodo: valida movimientos, documentos y terceros. Revise la extracción antes de ejecutar la validación.

La extracción XML no valida firma, autenticidad, CUFE ante DIAN, estado de aceptación ni reglas tributarias. El estado inicial es Revision Manual y Menu_ValidarTodo lo conserva así mientras exista la marca XML_REVISION. La persona confirma la revisión con Menu_ConfirmarRevisionXMLDocumentoActivo después de contrastar el soporte.

## Módulos

| Módulo | Responsabilidad |
|---|---|
| modUtils | Tablas, IDs, parámetros, notas por marcador, LOG y archivos |
| modValidation | Campos obligatorios y valores numéricos o fecha |
| modDuplicados | Posibles duplicados de documentos y terceros |
| modMovimientos | Alta y validación de movimientos |
| modDocumentos | Alta y validación de documentos |
| modTerceros | Alta y validación de terceros |
| modPagosCobros | Altas de pagos y cobros |
| modFiles | Importación, relectura XML y renombrado |
| modReportes | Actualización de CXC/CXP y recálculo |
| modXmlImport | Lectura UBL con MSXML 6.0 |
| modMain | Macros de entrada para Alt+F8 y botones |

Al importar una nueva versión, elimine primero el módulo anterior con el mismo nombre y luego importe el archivo actualizado.


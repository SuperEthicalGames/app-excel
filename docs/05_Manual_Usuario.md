# Manual de usuario

## Abrir y preparar una empresa

1. Guarde una copia del libro con el nombre de su empresa. El XLSX sirve para captura y reportes sin macros.
2. Si necesita importación de archivos, generación de IDs, validación masiva, actualización de CXC/CXP o lectura XML, guarde la copia como XLSM e importe los once módulos descritos en `02_Manual_Instalacion_VBA.md`.
3. En `EMPRESA`, actualice nombre, NIT, actividad, moneda y datos de contacto.
4. En `PARAMETROS`, revise prefijos, modo compatible y `RUTA_DOCUMENTOS`. La ruta inicial `Documentos` se crea junto al archivo del libro. Puede usar una ruta absoluta local o compartida.
5. Revise `CATEGORIAS`, `IMPUESTOS`, `CUENTAS`, `TIPOS_DOCUMENTOS`, `ESTADOS` y `MEDIOS_PAGO`. Las tasas de ejemplo no son universales: valide cada impuesto y parámetro contra la normativa aplicable y la política de la empresa.
6. Cargue una sola vez los clientes, proveedores y demás terceros en `TERCEROS`. El identificador fiscal debe conservarse como texto cuando incluya guion o dígito de verificación.

Al abrir el libro, use `INICIO` para ir a las tareas frecuentes. Cada título de tarea enlaza a la hoja correspondiente. Para una primera configuración, complete la empresa, revise catálogos y cuentas, agregue terceros y después registre operaciones.

## Usar el asistente local

1. Abra `ASISTENTE` desde `INICIO`.
2. Escriba el nombre del tercero en B5 si desea aplicar una regla específica para ese tercero. Este campo es opcional.
3. Escriba el concepto en B6. El asistente busca palabras o frases en el concepto y propone tipo, categoría, subcategoría, medio de pago y cuenta cuando la regla los contiene.
4. Lea el estado y revise cada sugerencia. Copie o use los valores al registrar el movimiento en `MOVIMIENTOS`; no se crea ni modifica ningún registro automáticamente.
5. Si aparece `Sin coincidencia`, registre manualmente o agregue una regla. Si aparece `Conflicto: ajuste prioridad`, abra `REGLAS_IA` y deje una prioridad mayor para la regla que corresponda.

En `REGLAS_IA`, active o desactive reglas, cambie sus palabras clave y complete la clasificación. Puede agregar reglas en las filas vacías hasta la fila 104. La prioridad aceptada es de 1 a 100. Mantenga una sola regla con la mayor prioridad para cada caso. Las listas desplegables toman los catálogos configurados en el libro.

El asistente funciona localmente con fórmulas de Excel, sin conexión, API o costo de servicio adicional. Es automatización por reglas editables: no es un modelo generativo ni aprende solo. Las reglas de ejemplo pueden no representar las políticas de su empresa; revíselas antes de usarlas. No calcula impuestos ni valida tratamiento tributario.

## Registrar movimientos

Escriba cada operación en una fila de `MOVIMIENTOS` o, en el XLSM, ejecute `Menu_NuevoMovimiento` para crearla con ID y fecha. Use los catálogos de Tipo, Categoría, Medio de pago, Cuenta y Estado.

`Total` se calcula como `Subtotal + Impuestos - Retenciones`; no se digita manualmente. Las tasas no se calculan automáticamente. El estado `Pagado` incluye el valor en los indicadores de caja del dashboard; `Pendiente` alimenta cuentas por cobrar o pagar según el tipo del movimiento. `Anulado` se excluye de los indicadores pagados y pendientes.

Si existe un soporte, escriba su ID en `SoporteID`. La operación y el documento son registros separados: la extracción no crea automáticamente un movimiento financiero.

## Importar y revisar documentos

En el XLSM ejecute `Menu_ImportarDocumento`, seleccione un archivo y complete los campos que falten. XML UBL compatible puede proporcionar tipo, fecha, número, tercero, identificación, subtotal, impuestos, retenciones, total y UUID/CUFE de referencia. La extracción deja el documento en `Revision Manual` para cotejo humano.

PDF, imágenes y CSV se copian a la carpeta configurada, pero esta versión no ejecuta OCR ni extracción de texto. Complete los campos desde el soporte. La importación inicial usa un nombre temporal con fecha e ID; después de completar Fecha, Tercero, TipoDocumento y Numero, ejecute `Menu_RenombrarArchivoDocumentoActivo` para usar la convención final.

Seleccione una celda de la fila del documento antes de ejecutar `Menu_ExtraerXMLDocumentoActivo` o `Menu_ConfirmarRevisionXMLDocumentoActivo`. La relectura pide confirmación si reemplazará datos existentes. Confirme la revisión solo después de cotejar los valores y resolver los avisos. Luego ejecute `Menu_ValidarTodo`.

La herramienta no valida firmas, autenticidad, estado DIAN, validez fiscal del CUFE, XML contra XSD ni reglas tributarias. Un XML legible no equivale a un documento aceptado fiscalmente.

## Dashboard y reportes

En `DASHBOARD`, deje un filtro vacío para incluir todos sus valores. Los filtros son Año, Mes, Tipo, Categoría, Tercero, Cuenta y Estado. Los conteos de documentos aplican Año, Mes y Tercero; los importes financieros aplican todos los filtros que correspondan a MOVIMIENTOS.

En `REPORTE_MENSUAL`, escriba el año en B3. En `REPORTE_ANUAL`, escriba en B3 el primer año de la ventana de 12 años. Los reportes incluyen movimientos con estado `Pagado` y calculan los importes desde Total.

En el XLSM, `Menu_ActualizarReportes` recalcula las fórmulas y repuebla `CXC` y `CXP` con movimientos pendientes. `LOG` registra acciones ejecutadas por las macros; las ediciones manuales de celdas no se auditan.

## Copias y cuidado de datos

- Mantenga una copia de respaldo del XLSM y de la carpeta de soportes antes de actualizaciones estructurales.
- No elimine filas duplicadas desde los avisos: confirme el documento fuente y decida manualmente cuál conservar.
- La protección de hojas está desactivada para permitir la edición libre y la expansión de tablas. Los roles de `USUARIOS_ROLES` son informativos y no constituyen controles de acceso.
- El diseño tiene rangos analíticos iniciales de 10.000 filas para movimientos, documentos, pagos y cobros; 2.000 para terceros. Consulte el manual técnico antes de superar esos límites.


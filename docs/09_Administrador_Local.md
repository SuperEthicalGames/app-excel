# Administrador financiero local

## Objetivo y estado

Aplicación de escritorio con interfaz web local que administra el libro XLSX existente. Busca que una persona pueda consultar y mantener los registros desde formularios simples o describiendo la tarea en lenguaje natural.

## Arquitectura

```text
Navegador local
  ├── Panel, tablas, formularios y vista previa de cambios
  └── HTTP restringido a 127.0.0.1:8765
          ├── Python: validación, respaldo, escritura y auditoría
          ├── Ollama: inferencia del modelo en el equipo
          ├── SQLite local: conversaciones y memorias aprobadas, separadas por libro
          └── openpyxl: lectura y escritura del XLSX
```

El XLSX conserva el papel de fuente de datos financiera editable. SQLite almacena únicamente el historial conversacional y las preferencias que el usuario confirma; no sustituye los registros financieros del libro. Los catálogos que sugieren los formularios se leen del libro seleccionado.

## Recorrido del asistente

```text
Solicitud en español
    ↓
Modelo local genera JSON limitado por esquema
    ↓
Validación de campos, tipos, categorías, cuentas y duplicados
    ↓
Vista previa con supuestos y correcciones transparentes
    ↓
Confirmación expresa del usuario
    ↓
Copia de respaldo → escritura XLSX → fila en LOG
```

La propuesta caduca a los diez minutos y se invalida si cambia el libro entre la vista previa y la confirmación. La app valida nuevamente cada acción al guardar. Un ajuste automático de categoría solo se ofrece cuando la subcategoría y el tipo apuntan a una única fila configurada; la corrección aparece en la vista previa.

El asistente entiende expresiones cotidianas y referencias a turnos anteriores. Cuando un dato obligatorio no se puede inferir, pregunta por ese dato concreto y ofrece las opciones configuradas cuando están disponibles. Las conversaciones se guardan en `app/data/assistant_memory.sqlite3`; el selector del chat permite retomarlas, iniciar otra y borrar una conversación. La memoria de cada conversación se limita al libro activo.

El usuario puede decir, por ejemplo, «recuerda que prefiero usar la cuenta Bancolombia». La app propone la preferencia y solo la guarda al confirmar. Las memorias aprobadas se muestran en el panel del asistente, se pueden olvidar individualmente y se envían al modelo local como contexto del libro. No se guardan contraseñas, tokens ni claves de acceso. La base incluye referencias de origen y una tabla preparada para ampliar esta memoria con reglas explícitas y auditables.

Los alias de campos se resuelven según el módulo: por ejemplo, «identificación» corresponde a `NIT_ID` en movimientos y terceros, y a `Identificacion` en documentos. Los campos opcionales se pueden vaciar desde el formulario con **Vaciar**; un campo en blanco no modifica el valor existente si no se marcó esa acción.

## Módulos actuales

| Módulo | Acciones |
|---|---|
| Resumen | Indicadores del mes, documentos y movimientos recientes |
| Movimientos | Consultar, crear, editar y anular mediante el estado |
| Documentos | Consultar, crear y editar metadatos; detectar coincidencia por CUFE, NIT+número, NIT+fecha+total o tercero+número+total. Conserva la fila y la marca **Duplicado** |
| Terceros | Consultar, crear y editar; bloquea NIT/ID repetido aunque cambie puntuación y marca nombres coincidentes para revisión |
| Cuentas | Consultar, crear y editar cuentas |
| Categorías | Consultar, crear y editar el catálogo |
| Reglas del asistente | Consultar, crear, editar o desactivar reglas |
| Registro de actividad | Consultar cambios escritos desde la app |

Las altas y modificaciones desde formularios también generan respaldo y LOG. Para modificar un registro, se selecciona la fila y se pulsa **Editar**. Los formularios cargan opciones de tipos y catálogos existentes.

## Presupuesto y privacidad

- No hay llamadas a una API pagada ni suscripción de IA.
- El modelo se ejecuta mediante Ollama en el mismo equipo; la primera descarga del modelo requiere internet y varios GB de almacenamiento.
- El servidor escucha solamente en `127.0.0.1`; el libro, los chats y las memorias no se transmiten a un servicio externo.
- El archivo SQLite se mantiene aparte del Excel, bajo `app/data/assistant_memory.sqlite3`, y separa las conversaciones por ruta del libro. Borrar una conversación elimina sus mensajes; las memorias aprobadas se administran por separado.
- Se necesita Python con `openpyxl`; se instala con `python -m pip install -r requirements.txt`.
- El rendimiento de inferencia depende de la memoria y el procesador/GPU disponibles.

## Guardado, compatibilidad y límites

- Solo se aceptan libros `.xlsx` con hojas, tablas y columnas requeridas por la app.
- Las solicitudes que escriben datos deben llegar desde el origen local de la interfaz; se rechazan solicitudes de otros sitios y cuerpos mayores a 2 MB.
- El cambio de libro y las escrituras se sincronizan para evitar que una operación termine guardada en otro libro.
- Las altas manuales llevan un identificador temporal para que repetir una solicitud tras un fallo de red no duplique el registro.
- Los formularios de **Cancelar** y **Cerrar** omiten la validación del navegador y nunca ejecutan el guardado.
- Cierra el libro en Excel antes de guardar desde la app para evitar cambios simultáneos.
- Antes de cada escritura se crea `backups/<libro>_before_<fecha>.xlsx` junto al libro.
- La escritura se valida como archivo XLSX y activa el recálculo de fórmulas al abrirlo en Excel.
- La app no borra registros físicamente. Se usan estados como **Anulado** o **No** cuando aplica.
- No incluye aún importación de archivos, OCR, conexión bancaria/DIAN, nómina, impuestos configurables, administración de EMPRESA/PARÁMETROS ni TRA.
- Excel 2019 no es requisito para ejecutar la interfaz; el formato de datos sigue siendo XLSX. La app no depende de fórmulas exclusivas de Microsoft 365.

## Pruebas realizadas

| Caso | Entrada | Resultado esperado | Resultado obtenido | Estado |
|---|---|---|---|---|
| Inicio local | `GET /api/status` | Libro y modelo local disponibles | XLSX principal y `gemma4:e4b-it-qat` detectados | Aprobado |
| Interfaz | Abrir la página local | Panel visible, navegación y tarjetas | Panel y módulos visibles en navegador | Aprobado |
| Propuesta IA | “Registra un gasto de 120000 por arriendo, pendiente” | JSON válido y vista previa | Propuso el gasto; ajustó Operacion→Administracion por coincidencia única y mostró la razón | Aprobado |
| Confirmación | Cerrar la vista previa de la prueba | No escribir el movimiento de prueba | No se confirmó ni se modificó el libro real | Aprobado |
| Alta y edición | Movimiento en copia temporal | Crear y después actualizar un campo | Ambas operaciones guardadas en copia temporal | Aprobado |
| Cálculo | Movimiento de prueba con subtotal 120000 | Total calculado y fórmula presente | Total 120000 y fórmula presente | Aprobado |
| Respaldo y auditoría | Altas y edición temporales | Copia previa y entradas LOG | Se creó respaldo por escritura y filas LOG | Aprobado |
| NIT duplicado | Crear dos terceros con el mismo NIT | Rechazar la segunda alta | Rechazo con mensaje y sin alterar el archivo original | Aprobado |
| Alias de campos | Identificación, número y descripción en distintos módulos | Mapear a la columna correcta según el módulo | Pruebas unitarias para movimientos, documentos, terceros y categorías | Aprobado |
| Campos vacíos | Vaciar Email al editar un tercero | Guardar vacío solo con instrucción explícita y dejar respaldo/LOG | Copia temporal guardada, respaldada y auditada | Aprobado |
| Duplicado de tercero | NIT con y sin puntuación; nombre con tildes y signos | Bloquear NIT repetido; marcar nombre posible | NIT bloqueado; nombre guardado con observación de revisión | Aprobado |
| Duplicado documental | Mismo NIT y número | Conservar y marcar, sin eliminar | Propuesta marcada como **Duplicado** con observación y ajuste | Aprobado |
| Origen HTTP | Solicitud local y solicitud externa | Aceptar mismo origen; rechazar origen ajeno | Política comprobada con pruebas unitarias | Aprobado |
| Cancelar formulario | Abrir un alta vacía y pulsar Cancelar | Cerrar sin validar ni enviar guardado | El diálogo cerró; no hubo POST a `/api/records/save` | Aprobado |
| Reintento de alta | Enviar dos veces el mismo identificador de solicitud | Crear solo una fila | Prueba temporal devolvió el mismo resultado y mantuvo un solo registro | Aprobado |
| Estado de tercero | Abrir formulario de alta | Mostrar solo Activo/Inactivo | Selector muestra las dos opciones correctas | Aprobado |
| Sintaxis | `py_compile` y `node --check` | Python y JavaScript válidos | Ambos comandos terminaron correctamente | Aprobado |
| Referencias VBA | Columnas y tablas citadas en los módulos | Resolver contra las tablas del libro | No se encontraron referencias inexistentes | Aprobado |
| Compilación VBA | Compilar módulos en Excel | Sin errores de compilación | Excel no estaba abierto durante la auditoría; sin compilación en vivo | Pendiente |
| Excel 2019 | Abrir desde Excel 2019 | Verificar edición y recálculo | No ejecutado en Excel 2019 | Pendiente |
| Importación/OCR | PDF, XML e imagen desde la interfaz | Extraer y revisar documento | No implementado en esta app | Pendiente |

Las pruebas de escritura se ejecutaron sobre copias temporales. El libro real solo se leyó; la solicitud de prueba en el navegador se canceló antes de confirmar.

## Inicio y mantenimiento

```powershell
python -m pip install -r requirements.txt
ollama pull gemma4:e4b-it-qat
```

Después, cierra Excel y ejecuta `Abrir_Administrador.bat`. Para detener el servidor, cierra la ventana de consola si lo abriste manualmente. Revisa los respaldos y `LOG` según las políticas de retención de la empresa.

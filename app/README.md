# Administrador financiero local

Aplicación web local para consultar y administrar el libro financiero de Excel. La interfaz se ejecuta en el navegador, pero el servidor solo escucha en `127.0.0.1`; no se publica en internet.

## Abrir

1. Instala Python 3.10 o posterior.
2. Instala la dependencia que ya usa el proyecto: `python -m pip install -r requirements.txt`.
3. Instala Ollama y descarga una vez el modelo local con `ollama pull gemma4:e4b-it-qat`.
4. Cierra el libro en Excel.
5. Abre `Abrir_Administrador.bat`. La app abre el navegador en `http://127.0.0.1:8765`.

La app carga por defecto `dist/Sistema_Financiero_Empresa_2026.xlsx`. Usa **Cambiar libro** para seleccionar otra copia con la misma estructura.

## Funciones disponibles

- Panel de indicadores y movimientos recientes.
- Crear, consultar y editar movimientos, documentos, terceros, cuentas, categorías y reglas del asistente.
- Formularios con catálogos del libro para reducir errores de digitación.
- Asistente conversacional local para preparar altas y cambios; siempre presenta una vista previa y requiere confirmación.
- Detección de NIT repetido en terceros; coincidencias de nombre quedan marcadas para revisión. Los documentos posiblemente duplicados se conservan con estado **Duplicado** para revisión.
- En edición, **Vaciar** borra de forma explícita un campo opcional; dejarlo en blanco sin marcarlo conserva el dato guardado.
- Respaldo automático y registro de auditoría antes de guardar.

La aplicación cuesta **$0 en servicios de IA**: usa un modelo local y no requiere API ni suscripción. La descarga inicial del modelo necesita conexión a internet y espacio en disco. El hardware y el consumo eléctrico dependen del equipo.

## Asistente

El asistente usa Ollama en este equipo. La primera respuesta puede tardar mientras el modelo se carga. La IA no guarda cambios por sí sola: presenta una propuesta y espera a que el usuario la confirme. Puede crear o editar movimientos, documentos, terceros, cuentas, categorías y reglas. No admite la eliminación física de registros. Los movimientos se pueden anular y las reglas, categorías y cuentas se pueden desactivar mediante una edición.

Si Ollama está detenido, la app mantiene disponibles la consulta de datos y los formularios de creación y edición.

El modelo local es el motor de IA de esta versión; no se entrenó un modelo desde cero. Si el modelo preferido no está instalado pero existe otro modelo local, la app lo usa como respaldo.

## Guardado y respaldo

Cada guardado genera primero una copia fechada dentro de `backups` junto al libro seleccionado. La app verifica hojas, tablas y columnas requeridas, conserva los campos calculados, solicita recálculo al abrir en Excel y agrega cada cambio a la hoja `LOG`. Cierra Excel antes de guardar desde la app para evitar ediciones simultáneas.

Los cambios se hacen en el XLSX seleccionado. Protege el archivo y la carpeta de respaldos con las políticas de la empresa.

La app no elimina registros físicamente. Las funciones de importación de PDF/XML/imágenes y OCR siguen en el flujo Excel/VBA del proyecto y no se ejecutan desde esta interfaz todavía. Consulta el alcance y el QA en `docs/09_Administrador_Local.md`.

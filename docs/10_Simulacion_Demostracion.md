# Simulación guiada del administrador

## Acceso

Con el servidor local activo, abre `http://127.0.0.1:8765/simulacion.html` o usa **Ver simulación** en la barra superior del administrador.

## Alcance

El recorrido muestra siete etapas: configuración de una empresa ficticia, recepción ilustrativa de XML e imagen, extracción, validación, registro de movimientos, resumen financiero y descarga del libro Excel.

La extracción y el OCR de la página son demostrativos. No se cargan soportes reales ni se llama a un modelo o servicio externo. El documento con confianza baja queda en revisión manual para ilustrar el control.

## Archivo de ejemplo

Ejecuta `python build/create_demo_workbook.py` para regenerar `app/static/demo/Sistema_Financiero_Demo_2026.xlsx` y su manifiesto. El script parte de `dist/Sistema_Financiero_Empresa_2026.xlsx` y guarda una copia separada; no modifica ese libro. El archivo incluye todas las hojas originales, fórmulas y catálogos, además de datos sintéticos en terceros, documentos, movimientos, pagos, cobros y auditoría. El reporte de CxC/CxP usa fórmulas compatibles con Excel 2019 y se recalcula al abrir el libro en Excel.

Los importes se usan solo para mostrar el flujo. Impuestos y retenciones quedan en cero en la muestra; no son una recomendación ni una liquidación tributaria. Las identificaciones contienen la marca `FICTICIO` y no deben usarse como datos empresariales.

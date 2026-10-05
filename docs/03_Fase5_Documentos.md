# Fase 5: importación y lectura de documentos

## Flujo

1. `Menu_ImportarDocumento` copia el soporte a la ruta configurada en `PARAMETROS`.
2. Para XML, `modXmlImport` usa MSXML 6.0 y XPath para leer campos UBL de `Invoice`, `CreditNote` o `DebitNote`.
3. PDF, JPG, PNG y CSV se conservan y quedan pendientes para OCR o lectura externa.
4. Los XML reconocidos quedan en `Revision Manual` para comparar los datos con el soporte.
5. `Menu_ValidarTodo` no marca un XML como Validado mientras conserve la marca `XML_REVISION`.
6. Después de revisar los campos y resolver avisos, `Menu_ConfirmarRevisionXMLDocumentoActivo` registra la confirmación humana. Luego se ejecuta `Menu_ValidarTodo`.
7. Crear el movimiento financiero y relacionarlo con el soporte sigue siendo una decisión manual.

## Campos UBL mapeados

| Campo DOCUMENTOS | Ruta relativa UBL | Regla |
|---|---|---|
| TipoDocumento | Nodo Invoice, CreditNote o DebitNote | Se traduce al catálogo actual |
| Numero | ID del documento | Texto exacto |
| Fecha | IssueDate | Fecha ISO AAAA-MM-DD |
| CUFE_Referencia | UUID | Se guarda como referencia; no se verifica |
| Tercero | RegistrationName; alternativas Name del Party | Se escoge la parte distinta de EMPRESA.NIT si la identificación coincide |
| Identificacion | PartyTaxScheme/CompanyID, PartyLegalEntity/CompanyID o PartyIdentification/ID | Si schemeName=31 y schemeID contiene un dígito, se agrega como dígito de verificación |
| Subtotal | LegalMonetaryTotal/TaxExclusiveAmount; alternativa LineExtensionAmount | Valor numérico |
| Impuestos | Suma de TaxTotal/TaxAmount del documento | No infiere tasas |
| Retenciones | Suma de WithholdingTaxTotal/TaxAmount del documento | No infiere tasas |
| Total | LegalMonetaryTotal/PayableAmount | Se copia cuando puede convertirse a número |
| Comparación del total | LegalMonetaryTotal/PayableAmount | Compara subtotal + impuestos - retenciones y deja diferencias en Observaciones |
| Confianza | Estado de lectura estructural | El valor 1 indica que se reconoció una estructura UBL y se ejecutó la extracción; no expresa certeza, exactitud semántica ni validez tributaria |

Los prefijos XML pueden variar porque se busca por nombre local. Los importes solo se escriben cuando `DocumentCurrencyCode` coincide con `EMPRESA.MonedaPrincipal`. Si falta el código o las monedas difieren, se dejan sin cambios y se registra un aviso. La tolerancia se configura con `TOLERANCIA_XML_TOTAL`; si el parámetro no existe, se usa 1 unidad monetaria.

La conciliación es una alerta básica. No modela por separado descuentos, cargos, anticipos, redondeos ni otras reglas particulares del documento; una diferencia exige revisión humana y no determina por sí sola un error tributario.

## Límites

La extracción lee el XML, pero no comprueba XSD, firma digital, autenticidad, estado DIAN, validez del CUFE, duplicidad fiscal ni cumplimiento tributario. La marca `XML_REVISION` impide cambiar el estado a Validado hasta que una persona la confirme expresamente. La DIAN mantiene anexos técnicos cuya vigencia debe confirmarse para cualquier uso normativo.

Se procesa XML UBL con nodos `Invoice`, `CreditNote` o `DebitNote`. No se procesan ZIP, XML incrustados como texto o base64 dentro de `AttachedDocument`, PDF, imágenes ni fotos. El módulo no llama a APIs ni transmite documentos.

MSXML 6.0 se crea con late binding. La carga prohíbe DTD y desactiva la resolución de entidades externas.

## OCR

No se configuró un proveedor OCR. PDFs e imágenes quedan pendientes para revisión manual. Una integración futura requiere definir proveedor, permisos, tratamiento de datos, costos y retención. El módulo XML queda aislado para que un adaptador OCR futuro alimente las mismas columnas de DOCUMENTOS.

## Fuentes técnicas

- DIAN, [documentación técnica de factura electrónica](https://www.dian.gov.co/impuestos/factura-electronica/documentacion/paginas/documentacion-tecnica.aspx).
- DIAN, [Resolución 165 de 2023 compilada, con referencias a la Resolución 227 de 2025](https://normograma.dian.gov.co/dian/compilacion/docs/resolucion_dian_0165_2023.htm).
- Microsoft, [XPath y SelectNodes en MSXML](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms754523%28v%3Dvs.85%29).
- Microsoft, [ProhibitDTD en MSXML](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms762632%28v%3Dvs.85%29).

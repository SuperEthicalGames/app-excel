Attribute VB_Name = "modXmlImport"
Option Explicit

' Lee campos de Invoice, CreditNote o DebitNote UBL.
' Late binding MSXML 6.0: no requiere activar referencias VBA.
' No valida firma, autenticidad, CUFE ante DIAN ni reglas tributarias.

Public Function ExtraerUBL(ByVal rutaXml As String, ByVal tabla As ListObject, _
                           ByVal fila As ListRow) As String
    On Error GoTo ManejarError

    Dim xml As Object, documento As Object, parseError As Object
    Dim nombreRaiz As String, numero As String, fechaTexto As String
    Dim fechaDocumento As Date, cufe As String, subtotalTexto As String
    Dim monedaXML As String, monedaEmpresa As String, monedaOk As Boolean
    Dim totalTexto As String, impuestoTotal As Double, retencionTotal As Double
    Dim totalDeclarado As Double, totalCalculado As Double, tolerancia As Double
    Dim subtotal As Double, subtotalOk As Boolean, impuestoOk As Boolean
    Dim retencionOk As Boolean, payableOk As Boolean, payableParseOk As Boolean
    Dim proveedorNombre As String, proveedorID As String
    Dim clienteNombre As String, clienteID As String
    Dim terceroNombre As String, terceroID As String, nitEmpresa As String
    Dim warnings As String, detalle As String
    Dim camposXml As Variant, valoresPrevios() As Variant, i As Long
    Dim datosPreviosCapturados As Boolean

    camposXml = Array("Fecha", "TipoDocumento", "Numero", "Tercero", "Identificacion", _
                      "Subtotal", "Impuestos", "Retenciones", "Total", "CUFE_Referencia", "Confianza")
    ReDim valoresPrevios(LBound(camposXml) To UBound(camposXml))
    For i = LBound(camposXml) To UBound(camposXml)
        valoresPrevios(i) = fila.Range.Cells(1, tabla.ListColumns(CStr(camposXml(i))).Index).Value
    Next i
    datosPreviosCapturados = True
    For i = LBound(camposXml) To UBound(camposXml)
        fila.Range.Cells(1, tabla.ListColumns(CStr(camposXml(i))).Index).Value = Empty
    Next i

    EscribirCampo tabla, fila, "MetodoExtraccion", "XML"
    EscribirCampo tabla, fila, "FechaExtraccion", Now

    Set xml = CreateObject("Msxml2.DOMDocument.6.0")
    xml.async = False
    xml.validateOnParse = False
    xml.resolveExternals = False
    xml.setProperty "ProhibitDTD", True

    If Not CBool(xml.Load(rutaXml)) Then
        Set parseError = xml.parseError
        detalle = "XML mal formado: " & CStr(parseError.reason) & _
                  " (linea " & CStr(parseError.line) & ")"
        RestaurarCamposXML tabla, fila, camposXml, valoresPrevios
        RegistrarEstadoXML tabla, fila, "Error", detalle
        ExtraerUBL = detalle
        Exit Function
    End If

    Set documento = xml.selectSingleNode( _
        "//*[local-name()='Invoice' or local-name()='CreditNote' or local-name()='DebitNote']")
    If documento Is Nothing Then
        detalle = "XML bien formado, pero no se encontro Invoice, CreditNote o DebitNote UBL."
        RestaurarCamposXML tabla, fila, camposXml, valoresPrevios
        RegistrarEstadoXML tabla, fila, "Revision Manual", detalle, _
            "Confirme manualmente el tipo y los campos del documento."
        ExtraerUBL = detalle
        Exit Function
    End If

    nombreRaiz = CStr(documento.baseName)
    Select Case UCase$(nombreRaiz)
        Case "INVOICE"
            EscribirCampo tabla, fila, "TipoDocumento", "Factura electronica"
        Case "CREDITNOTE"
            EscribirCampo tabla, fila, "TipoDocumento", "Nota credito"
        Case "DEBITNOTE"
            EscribirCampo tabla, fila, "TipoDocumento", "Nota debito"
    End Select

    numero = XmlTexto(documento, "./*[local-name()='ID']")
    fechaTexto = XmlTexto(documento, "./*[local-name()='IssueDate']")
    cufe = XmlTexto(documento, "./*[local-name()='UUID']")

    If numero <> "" Then
        EscribirCampo tabla, fila, "Numero", numero
    Else
        warnings = AgregarAviso(warnings, "Falta el numero de documento.")
    End If
    If ParsearFechaISO(fechaTexto, fechaDocumento) Then
        EscribirCampo tabla, fila, "Fecha", fechaDocumento
    Else
        warnings = AgregarAviso(warnings, "Falta IssueDate o no tiene formato ISO valido.")
    End If
    If cufe <> "" Then
        EscribirCampo tabla, fila, "CUFE_Referencia", cufe
    Else
        warnings = AgregarAviso(warnings, "No se encontro UUID/CUFE; verificar si aplica.")
    End If

    LeerParte documento, "AccountingSupplierParty", proveedorNombre, proveedorID
    LeerParte documento, "AccountingCustomerParty", clienteNombre, clienteID

    nitEmpresa = modUtils.ObtenerEmpresaValor("NIT", "")
    If NormalizarID(nitEmpresa) <> "" And _
       NormalizarID(nitEmpresa) = NormalizarID(proveedorID) Then
        terceroNombre = clienteNombre
        terceroID = clienteID
    ElseIf NormalizarID(nitEmpresa) <> "" And _
           NormalizarID(nitEmpresa) = NormalizarID(clienteID) Then
        terceroNombre = proveedorNombre
        terceroID = proveedorID
    Else
        terceroNombre = proveedorNombre
        terceroID = proveedorID
        warnings = AgregarAviso(warnings, _
            "No se pudo identificar la empresa entre las partes; se propone el emisor como tercero. Confirmar.")
    End If

    If terceroNombre <> "" Then
        EscribirCampo tabla, fila, "Tercero", terceroNombre
    Else
        warnings = AgregarAviso(warnings, "No se encontro el nombre del tercero.")
    End If
    If terceroID <> "" Then
        EscribirCampo tabla, fila, "Identificacion", terceroID
    Else
        warnings = AgregarAviso(warnings, "No se encontro la identificacion del tercero.")
    End If

    monedaXML = UCase$(XmlTexto(documento, "./*[local-name()='DocumentCurrencyCode']"))
    monedaEmpresa = UCase$(Trim$(modUtils.ObtenerEmpresaValor("MonedaPrincipal", "COP")))
    monedaOk = (monedaXML <> "" And monedaEmpresa <> "" And monedaXML = monedaEmpresa)
    If Not monedaOk Then
        warnings = AgregarAviso(warnings, _
            "Moneda XML=" & IIf(monedaXML = "", "sin informar", monedaXML) & _
            "; moneda empresa=" & IIf(monedaEmpresa = "", "sin configurar", monedaEmpresa) & _
            ". No se escribieron los valores monetarios; confirmar moneda.")
    End If

    If monedaOk Then
        subtotalTexto = XmlTexto(documento, _
            "./*[local-name()='LegalMonetaryTotal']/*[local-name()='TaxExclusiveAmount']")
        If subtotalTexto = "" Then
            subtotalTexto = XmlTexto(documento, _
                "./*[local-name()='LegalMonetaryTotal']/*[local-name()='LineExtensionAmount']")
        End If
        subtotalOk = ParsearMontoUBL(subtotalTexto, subtotal)
        If subtotalOk Then
            EscribirCampo tabla, fila, "Subtotal", subtotal
        Else
            warnings = AgregarAviso(warnings, "No se extrajo un subtotal numerico.")
        End If

        impuestoOk = SumarMontosUBL(documento, _
            "./*[local-name()='TaxTotal']/*[local-name()='TaxAmount']", impuestoTotal)
        If impuestoOk Then
            EscribirCampo tabla, fila, "Impuestos", impuestoTotal
        Else
            warnings = AgregarAviso(warnings, "Hay un TaxAmount que no se pudo leer.")
        End If

        retencionOk = SumarMontosUBL(documento, _
            "./*[local-name()='WithholdingTaxTotal']/*[local-name()='TaxAmount']", retencionTotal)
        If retencionOk Then
            EscribirCampo tabla, fila, "Retenciones", retencionTotal
        Else
            warnings = AgregarAviso(warnings, "Hay una retencion que no se pudo leer.")
        End If

        totalTexto = XmlTexto(documento, _
            "./*[local-name()='LegalMonetaryTotal']/*[local-name()='PayableAmount']")
        payableOk = (totalTexto <> "")
        payableParseOk = ParsearMontoUBL(totalTexto, totalDeclarado)
        If payableParseOk Then EscribirCampo tabla, fila, "Total", totalDeclarado

        If payableOk And Not payableParseOk Then
            warnings = AgregarAviso(warnings, "PayableAmount no es numerico.")
        ElseIf payableParseOk And subtotalOk And impuestoOk And retencionOk Then
            totalCalculado = subtotal + impuestoTotal - retencionTotal
            tolerancia = Val(modUtils.ObtenerParametro("TOLERANCIA_XML_TOTAL", "1"))
            If tolerancia < 0 Then tolerancia = 1
            If Abs(totalCalculado - totalDeclarado) > tolerancia Then
                warnings = AgregarAviso(warnings, _
                    "DIFERENCIA_TOTAL: XML=" & Format$(totalDeclarado, "0.00") & _
                    "; calculado=" & Format$(totalCalculado, "0.00") & _
                    "; tolerancia=" & Format$(tolerancia, "0.00") & ".")
            End If
        ElseIf Not payableOk Then
            warnings = AgregarAviso(warnings, "No se encontro PayableAmount para comparar el total.")
        End If
    End If

    EscribirCampo tabla, fila, "Confianza", 1
    detalle = "Extraccion UBL terminada. La lectura no verifica firma, autenticidad ni estado DIAN."
    If warnings <> "" Then detalle = detalle & " " & warnings
    RegistrarEstadoXML tabla, fila, "Revision Manual", detalle, _
        "Confirme que reviso el soporte y los campos extraidos." & _
        IIf(warnings <> "", " " & warnings, "")

    ExtraerUBL = "Campos UBL leidos. Estado=Revision Manual." & _
                 IIf(warnings <> "", " " & warnings, "")
    Exit Function

ManejarError:
    detalle = "Error de extraccion XML (" & CStr(Err.Number) & "): " & Err.Description
    On Error Resume Next
    If datosPreviosCapturados Then RestaurarCamposXML tabla, fila, camposXml, valoresPrevios
    RegistrarEstadoXML tabla, fila, "Error", detalle
    ExtraerUBL = detalle
End Function

Private Sub RestaurarCamposXML(ByVal tabla As ListObject, ByVal fila As ListRow, _
                               ByVal camposXml As Variant, ByRef valoresPrevios() As Variant)
    Dim i As Long
    For i = LBound(camposXml) To UBound(camposXml)
        fila.Range.Cells(1, tabla.ListColumns(CStr(camposXml(i))).Index).Value = valoresPrevios(i)
    Next i
End Sub

Private Sub LeerParte(ByVal documento As Object, ByVal nombreGrupo As String, _
                      ByRef nombre As String, ByRef identificacion As String)
    Dim parte As Object, nodoID As Object, atributo As Object
    Dim esquema As String, digito As String

    nombre = ""
    identificacion = ""
    Set parte = documento.selectSingleNode( _
        "./*[local-name()='" & nombreGrupo & "']/*[local-name()='Party']")
    If parte Is Nothing Then Exit Sub

    nombre = XmlTexto(parte, "./*[local-name()='PartyLegalEntity']/*[local-name()='RegistrationName']")
    If nombre = "" Then nombre = XmlTexto(parte, "./*[local-name()='PartyName']/*[local-name()='Name']")
    If nombre = "" Then
        nombre = XmlTexto(parte, "./*[local-name()='PartyTaxScheme']/*[local-name()='RegistrationName']")
    End If

    Set nodoID = parte.selectSingleNode( _
        "./*[local-name()='PartyTaxScheme']/*[local-name()='CompanyID']")
    If nodoID Is Nothing Then
        Set nodoID = parte.selectSingleNode( _
            "./*[local-name()='PartyLegalEntity']/*[local-name()='CompanyID']")
    End If
    If nodoID Is Nothing Then
        Set nodoID = parte.selectSingleNode( _
            "./*[local-name()='PartyIdentification']/*[local-name()='ID']")
    End If
    If nodoID Is Nothing Then Exit Sub

    identificacion = Trim$(CStr(nodoID.Text))
    Set atributo = nodoID.Attributes.getNamedItem("schemeName")
    If Not atributo Is Nothing Then esquema = Trim$(CStr(atributo.Text))
    Set atributo = nodoID.Attributes.getNamedItem("schemeID")
    If Not atributo Is Nothing Then digito = Trim$(CStr(atributo.Text))

    If esquema = "31" And Len(digito) = 1 And digito Like "#" Then
        If NormalizarID(identificacion) <> "" Then identificacion = identificacion & "-" & digito
    End If
End Sub

Private Function XmlTexto(ByVal nodoPadre As Object, ByVal xpath As String) As String
    Dim nodo As Object
    Set nodo = nodoPadre.selectSingleNode(xpath)
    If Not nodo Is Nothing Then XmlTexto = Trim$(CStr(nodo.Text))
End Function

Private Function SumarMontosUBL(ByVal documento As Object, ByVal xpath As String, _
                                ByRef suma As Double) As Boolean
    Dim nodos As Object, nodo As Object, valor As Double
    suma = 0
    Set nodos = documento.selectNodes(xpath)
    If nodos Is Nothing Then
        SumarMontosUBL = True
        Exit Function
    End If
    For Each nodo In nodos
        If Not ParsearMontoUBL(CStr(nodo.Text), valor) Then Exit Function
        suma = suma + valor
    Next nodo
    SumarMontosUBL = True
End Function

Private Function ParsearMontoUBL(ByVal texto As String, ByRef resultado As Double) As Boolean
    Dim normalizado As String, separador As String
    normalizado = Trim$(texto)
    If normalizado = "" Or InStr(1, normalizado, ",", vbBinaryCompare) > 0 Then Exit Function
    separador = CStr(Application.International(xlDecimalSeparator))
    If separador <> "." Then normalizado = Replace(normalizado, ".", separador)
    If Not IsNumeric(normalizado) Then Exit Function
    resultado = CDbl(normalizado)
    ParsearMontoUBL = True
End Function

Private Function ParsearFechaISO(ByVal texto As String, ByRef resultado As Date) As Boolean
    Dim anio As Long, mes As Long, dia As Long, candidata As Date
    texto = Trim$(texto)
    If Len(texto) <> 10 Then Exit Function
    If Mid$(texto, 5, 1) <> "-" Or Mid$(texto, 8, 1) <> "-" Then Exit Function
    If Not IsNumeric(Left$(texto, 4)) Or Not IsNumeric(Mid$(texto, 6, 2)) _
       Or Not IsNumeric(Right$(texto, 2)) Then Exit Function
    anio = CLng(Left$(texto, 4))
    mes = CLng(Mid$(texto, 6, 2))
    dia = CLng(Right$(texto, 2))
    On Error GoTo FechaInvalida
    candidata = DateSerial(anio, mes, dia)
    If Year(candidata) <> anio Or Month(candidata) <> mes Or Day(candidata) <> dia Then Exit Function
    resultado = candidata
    ParsearFechaISO = True
    Exit Function
FechaInvalida:
    ParsearFechaISO = False
End Function

Private Function NormalizarID(ByVal texto As String) As String
    Dim i As Long, caracter As String, salida As String
    For i = 1 To Len(texto)
        caracter = Mid$(texto, i, 1)
        If caracter Like "#" Then salida = salida & caracter
    Next i
    NormalizarID = salida
End Function

Private Function AgregarAviso(ByVal lista As String, ByVal aviso As String) As String
    If lista = "" Then
        AgregarAviso = aviso
    Else
        AgregarAviso = lista & " " & aviso
    End If
End Function

Private Sub EscribirCampo(ByVal tabla As ListObject, ByVal fila As ListRow, _
                          ByVal nombreColumna As String, ByVal valor As Variant)
    fila.Range.Cells(1, tabla.ListColumns(nombreColumna).Index).Value = valor
End Sub

Private Sub RegistrarEstadoXML(ByVal tabla As ListObject, ByVal fila As ListRow, _
                               ByVal estado As String, ByVal mensaje As String, _
                               Optional ByVal revision As String = "")
    EscribirCampo tabla, fila, "Estado", estado
    EscribirCampo tabla, fila, "MetodoExtraccion", "XML"
    EscribirCampo tabla, fila, "FechaExtraccion", Now
    If estado = "Error" Then EscribirCampo tabla, fila, "Confianza", Empty
    modUtils.EstablecerNotaMarcada tabla, fila.Index, "XML", mensaje
    modUtils.EstablecerNotaMarcada tabla, fila.Index, "XML_REVISION", revision
End Sub

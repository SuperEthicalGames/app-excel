Attribute VB_Name = "modFiles"
Option Explicit

' Importa un soporte y crea una fila en DOCUMENTOS.
' XML UBL pasa por modXmlImport; PDF, imagen y CSV quedan para OCR externo
' o digitacion manual.

Public Sub ImportarDocumento()
    On Error GoTo ManejarError
    Dim rutaOrigen As Variant
    Dim tabla As ListObject, fila As ListRow, id As String
    Dim nombreOriginal As String, extension As String
    Dim carpetaDestino As String, nombreDestino As String, rutaDestino As String
    Dim resultadoExtraccion As String, metodo As String, nota As String, estadoExtraccion As String
    Dim archivoCopiado As Boolean, filaCreada As Boolean, importacionCompleta As Boolean
    Dim numeroError As Long, descripcionError As String

    rutaOrigen = Application.GetOpenFilename( _
        "Documentos (*.pdf;*.jpg;*.jpeg;*.png;*.xml;*.csv),*.pdf;*.jpg;*.jpeg;*.png;*.xml;*.csv", _
        , "Seleccionar documento a digitalizar")
    If VarType(rutaOrigen) = vbBoolean Then Exit Sub

    nombreOriginal = Dir(CStr(rutaOrigen))
    If InStrRev(nombreOriginal, ".") = 0 Then Err.Raise vbObjectError + 723, "ImportarDocumento", _
        "El archivo no tiene una extension reconocida."
    extension = LCase$(Mid$(nombreOriginal, InStrRev(nombreOriginal, ".") + 1))
    Select Case extension
        Case "pdf", "jpg", "jpeg", "png", "xml", "csv"
        Case Else
            Err.Raise vbObjectError + 724, "ImportarDocumento", "Formato de archivo no soportado: " & extension
    End Select
    carpetaDestino = ResolverRutaDocumentos( _
        modUtils.ObtenerParametro("RUTA_DOCUMENTOS", "Documentos"))
    modUtils.AsegurarCarpeta carpetaDestino

    Set tabla = modUtils.ObtenerTabla("DOCUMENTOS", "DOCUMENTOS")
    id = modUtils.SiguienteID(tabla, "ID", modUtils.ObtenerParametro("PREFIJO_DOCUMENTO", "DOC-"), 5)
    nombreDestino = Format(Date, "yyyy-mm-dd") & "_SINCLASIFICAR_OTRO_" & id & "." & extension
    rutaDestino = carpetaDestino & nombreDestino
    If Dir(rutaDestino) <> "" Then Err.Raise vbObjectError + 701, "ImportarDocumento", _
        "Ya existe el archivo destino: " & rutaDestino
    FileCopy CStr(rutaOrigen), rutaDestino
    archivoCopiado = True

    Set fila = tabla.ListRows.Add(AlwaysInsert:=True)
    filaCreada = True
    metodo = "Manual"
    If UCase$(extension) = "XML" Then metodo = "XML"
    With fila.Range
        .Cells(1, tabla.ListColumns("ID").Index).Value = id
        .Cells(1, tabla.ListColumns("Fecha").Index).Value = Date
        .Cells(1, tabla.ListColumns("Estado").Index).Value = "Pendiente"
        .Cells(1, tabla.ListColumns("RutaArchivo").Index).Value = rutaDestino
        .Cells(1, tabla.ListColumns("NombreArchivo").Index).Value = nombreDestino
        .Cells(1, tabla.ListColumns("TipoArchivo").Index).Value = UCase$(extension)
        .Cells(1, tabla.ListColumns("FechaExtraccion").Index).Value = Empty
        .Cells(1, tabla.ListColumns("MetodoExtraccion").Index).Value = metodo
        .Cells(1, tabla.ListColumns("Confianza").Index).Value = Empty
        .Cells(1, tabla.ListColumns("Observaciones").Index).Value = ""
    End With

    If UCase$(extension) = "XML" Then
        resultadoExtraccion = modXmlImport.ExtraerUBL(rutaDestino, tabla, fila)
        estadoExtraccion = CStr(tabla.ListColumns("Estado").DataBodyRange.Cells(fila.Index, 1).Value)
        modUtils.EscribirLog "Extraer XML", "DOCUMENTOS", id, "", resultadoExtraccion, _
            IIf(estadoExtraccion = "Error", "ERROR", "REVISION")
    Else
        nota = "Importado sin OCR ni extraccion PDF configurados. Complete los campos manualmente y revise el soporte."
        modUtils.EstablecerNotaMarcada tabla, fila.Index, "EXTRACCION", nota
        modUtils.EscribirLog "Importar", "DOCUMENTOS", id, "", nombreDestino, "PENDIENTE"
        resultadoExtraccion = nota
    End If

    importacionCompleta = True
    MsgBox "Documento copiado a:" & vbCrLf & rutaDestino & vbCrLf & vbCrLf & _
           "Registro: " & id & " | Estado: " & _
           CStr(tabla.ListColumns("Estado").DataBodyRange.Cells(fila.Index, 1).Value) & vbCrLf & _
           resultadoExtraccion, vbInformation, "Importacion"
    Application.Goto fila.Range.Cells(1, tabla.ListColumns("TipoDocumento").Index), True
    Exit Sub

ManejarError:
    numeroError = Err.Number
    descripcionError = Err.Description
    On Error Resume Next
    If Not importacionCompleta Then
        If filaCreada Then fila.Delete
        If archivoCopiado Then
            If Len(rutaDestino) > 0 Then
                If Dir(rutaDestino) <> "" Then Kill rutaDestino
            End If
        End If
    End If
    On Error GoTo 0
    modUtils.EscribirLog "Importar", "DOCUMENTOS", id, "", CStr(rutaOrigen), _
        "ERROR " & CStr(numeroError) & ": " & descripcionError
    MsgBox "No se pudo importar el documento: " & descripcionError, vbCritical, "Error"
End Sub

Public Sub ExtraerXMLDocumentoActivo()
    On Error GoTo ManejarError
    Dim tabla As ListObject, fila As ListRow
    Dim ruta As String, extension As String, resultado As String
    Dim numero As String, cufe As String, subtotal As String

    If Not ObtenerDocumentoActivo(tabla, fila) Then Exit Sub
    ruta = CStr(fila.Range.Cells(1, tabla.ListColumns("RutaArchivo").Index).Value)
    If ruta = "" Or Dir(ruta) = "" Then
        MsgBox "La fila no tiene un archivo accesible en RutaArchivo.", vbExclamation
        Exit Sub
    End If
    extension = Mid$(ruta, InStrRev(ruta, ".") + 1)
    If UCase$(extension) <> "XML" Then
        MsgBox "La extraccion estructurada procesa archivos XML UBL.", vbInformation
        Exit Sub
    End If

    numero = CStr(fila.Range.Cells(1, tabla.ListColumns("Numero").Index).Value)
    cufe = CStr(fila.Range.Cells(1, tabla.ListColumns("CUFE_Referencia").Index).Value)
    subtotal = CStr(fila.Range.Cells(1, tabla.ListColumns("Subtotal").Index).Value)
    If numero <> "" Or cufe <> "" Or subtotal <> "" Then
        If MsgBox("La relectura puede sustituir campos presentes en el XML. Continuar?", _
                  vbQuestion + vbYesNo, "Reprocesar XML") <> vbYes Then Exit Sub
    End If

    resultado = modXmlImport.ExtraerUBL(ruta, tabla, fila)
    modUtils.EscribirLog "Reprocesar XML", "DOCUMENTOS", _
        CStr(fila.Range.Cells(1, tabla.ListColumns("ID").Index).Value), "", resultado, _
        IIf(CStr(fila.Range.Cells(1, tabla.ListColumns("Estado").Index).Value) = "Error", "ERROR", "REVISION")
    MsgBox resultado, vbInformation, "Extraccion XML"
    Exit Sub

ManejarError:
    MsgBox "No se pudo extraer el XML: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub ConfirmarRevisionXMLDocumentoActivo()
    On Error GoTo ManejarError
    Dim tabla As ListObject, fila As ListRow, observaciones As String, id As String

    If Not ObtenerDocumentoActivo(tabla, fila) Then Exit Sub
    If UCase$(CStr(fila.Range.Cells(1, tabla.ListColumns("MetodoExtraccion").Index).Value)) <> "XML" Then
        MsgBox "La confirmacion aplica a documentos extraidos desde XML.", vbInformation
        Exit Sub
    End If

    observaciones = CStr(fila.Range.Cells(1, tabla.ListColumns("Observaciones").Index).Value)
    If InStr(1, observaciones, "[XML_REVISION]", vbTextCompare) = 0 Then
        MsgBox "Esta fila no tiene una revision XML pendiente.", vbInformation
        Exit Sub
    End If

    If MsgBox("Confirma que comparo los campos con el XML original y resolvio los avisos?", _
              vbQuestion + vbYesNo, "Confirmar revision humana") <> vbYes Then Exit Sub

    id = CStr(fila.Range.Cells(1, tabla.ListColumns("ID").Index).Value)
    modUtils.EstablecerNotaMarcada tabla, fila.Index, "XML_REVISION", ""
    modUtils.EscribirLog "Confirmar revision XML", "DOCUMENTOS", id, _
        "Revision Manual", "Revision humana confirmada", "OK"
    MsgBox "Revision confirmada. Ejecute Menu_ValidarTodo para aplicar las reglas y detectar duplicados.", _
           vbInformation, "Revision XML"
    Exit Sub

ManejarError:
    MsgBox "No se pudo confirmar la revision XML: " & Err.Description, vbCritical, "Error"
End Sub

Private Function ObtenerDocumentoActivo(ByRef tabla As ListObject, _
                                        ByRef fila As ListRow) As Boolean
    Dim filaActiva As Range, indice As Long
    Set tabla = modUtils.ObtenerTabla("DOCUMENTOS", "DOCUMENTOS")
    If tabla.DataBodyRange Is Nothing Then
        MsgBox "No hay documentos registrados.", vbInformation
        Exit Function
    End If
    If ActiveSheet.Name <> tabla.Parent.Name Then
        MsgBox "Seleccione una fila de la tabla DOCUMENTOS.", vbExclamation
        Exit Function
    End If
    Set filaActiva = Intersect(ActiveCell.EntireRow, tabla.DataBodyRange)
    If filaActiva Is Nothing Then
        MsgBox "Seleccione una fila de la tabla DOCUMENTOS.", vbExclamation
        Exit Function
    End If
    indice = filaActiva.Row - tabla.HeaderRowRange.Row
    Set fila = tabla.ListRows(indice)
    ObtenerDocumentoActivo = True
End Function

Private Function ResolverRutaDocumentos(ByVal rutaConfigurada As String) As String
    Dim base As String, ruta As String, esUnidadAbsoluta As Boolean
    base = Replace(Trim$(rutaConfigurada), "/", "\")
    If base = "" Then base = "Documentos"

    esUnidadAbsoluta = (Len(base) >= 3 And Mid$(base, 2, 1) = ":" And Mid$(base, 3, 1) = "\")
    If Len(base) >= 2 And Mid$(base, 2, 1) = ":" And Not esUnidadAbsoluta Then _
        Err.Raise vbObjectError + 722, "ResolverRutaDocumentos", _
                  "Use una ruta absoluta como C:\Documentos, no una ruta relativa a una unidad."

    If esUnidadAbsoluta Or _
       Left$(base, 2) = "\\" Or Left$(base, 1) = "\" Then
        ruta = base
    Else
        If ThisWorkbook.Path <> "" Then
            ruta = ThisWorkbook.Path & "\" & base
        Else
            ruta = CurDir$ & "\" & base
        End If
    End If

    If Right$(ruta, 1) <> "\" Then ruta = ruta & "\"
    ResolverRutaDocumentos = ruta
End Function

Public Sub RenombrarArchivoDesdeFila()
    On Error GoTo ManejarError
    Dim tabla As ListObject, filaActiva As Range, indice As Long
    Dim rutaActual As String, carpeta As String, extension As String
    Dim nuevoNombre As String, nuevaRuta As String
    Dim fecha As Variant, tercero As String, tipoDoc As String, numero As String
    Dim nombreAnterior As String, idDocumento As String, archivoRenombrado As Boolean
    Dim numeroError As Long, descripcionError As String

    Set tabla = modUtils.ObtenerTabla("DOCUMENTOS", "DOCUMENTOS")
    If tabla.DataBodyRange Is Nothing Then
        MsgBox "No hay documentos registrados.", vbExclamation
        Exit Sub
    End If
    If ActiveSheet.Name <> tabla.Parent.Name Then
        MsgBox "Seleccione una fila de la tabla DOCUMENTOS.", vbExclamation
        Exit Sub
    End If
    Set filaActiva = Intersect(ActiveCell.EntireRow, tabla.DataBodyRange)
    If filaActiva Is Nothing Then
        MsgBox "Ubique el cursor en una fila de la tabla DOCUMENTOS antes de ejecutar esta accion.", vbExclamation
        Exit Sub
    End If
    indice = filaActiva.Row - tabla.HeaderRowRange.Row
    rutaActual = CStr(tabla.ListColumns("RutaArchivo").DataBodyRange.Cells(indice, 1).Value)
    nombreAnterior = CStr(tabla.ListColumns("NombreArchivo").DataBodyRange.Cells(indice, 1).Value)
    idDocumento = CStr(tabla.ListColumns("ID").DataBodyRange.Cells(indice, 1).Value)
    If rutaActual = "" Or Dir(rutaActual) = "" Then
        MsgBox "Esta fila no tiene un archivo valido en RutaArchivo.", vbExclamation
        Exit Sub
    End If

    fecha = tabla.ListColumns("Fecha").DataBodyRange.Cells(indice, 1).Value
    tercero = CStr(tabla.ListColumns("Tercero").DataBodyRange.Cells(indice, 1).Value)
    tipoDoc = CStr(tabla.ListColumns("TipoDocumento").DataBodyRange.Cells(indice, 1).Value)
    numero = CStr(tabla.ListColumns("Numero").DataBodyRange.Cells(indice, 1).Value)
    If Not IsDate(fecha) Or tercero = "" Or tipoDoc = "" Or numero = "" Then
        MsgBox "Complete Fecha, Tercero, TipoDocumento y Numero antes de renombrar.", vbExclamation
        Exit Sub
    End If

    extension = Mid$(rutaActual, InStrRev(rutaActual, ".") + 1)
    If extension = "" Then
        MsgBox "El archivo no tiene una extensión que pueda conservarse al renombrar.", vbExclamation
        Exit Sub
    End If
    carpeta = Left$(rutaActual, InStrRev(rutaActual, "\"))
    nuevoNombre = Format(fecha, "yyyy-mm-dd") & "_" & modUtils.LimpiarNombreArchivo(tercero) & "_" & _
                  modUtils.LimpiarNombreArchivo(tipoDoc) & "_" & modUtils.LimpiarNombreArchivo(numero) & _
                  "." & extension
    nuevaRuta = carpeta & nuevoNombre
    If nuevaRuta = rutaActual Then
        MsgBox "El archivo ya tiene el nombre solicitado.", vbInformation
        Exit Sub
    End If
    If Dir(nuevaRuta) <> "" Then
        MsgBox "Ya existe un archivo con ese nombre destino:" & vbCrLf & nuevaRuta, vbExclamation
        Exit Sub
    End If

    Name rutaActual As nuevaRuta
    archivoRenombrado = True
    tabla.ListColumns("RutaArchivo").DataBodyRange.Cells(indice, 1).Value = nuevaRuta
    tabla.ListColumns("NombreArchivo").DataBodyRange.Cells(indice, 1).Value = nuevoNombre
    modUtils.EscribirLog "Renombrar", "DOCUMENTOS", idDocumento, rutaActual, nuevaRuta, "OK"
    MsgBox "Archivo renombrado a:" & vbCrLf & nuevoNombre, vbInformation
    Exit Sub

ManejarError:
    numeroError = Err.Number
    descripcionError = Err.Description
    If archivoRenombrado Then
        On Error Resume Next
        If Dir(rutaActual) = "" Then
            If Dir(nuevaRuta) <> "" Then Name nuevaRuta As rutaActual
        End If
        tabla.ListColumns("RutaArchivo").DataBodyRange.Cells(indice, 1).Value = rutaActual
        tabla.ListColumns("NombreArchivo").DataBodyRange.Cells(indice, 1).Value = nombreAnterior
        On Error GoTo 0
    End If
    MsgBox "No se pudo renombrar el archivo: " & descripcionError & " (" & numeroError & ")", vbCritical, "Error"
End Sub

Attribute VB_Name = "modUtils"
Option Explicit

' Funciones de apoyo transversales: acceso a tablas, generacion de IDs,
' lectura de PARAMETROS/EMPRESA, registro en LOG y utilidades de archivos.

Public Function ObtenerTabla(ByVal nombreHoja As String, ByVal nombreTabla As String) As ListObject
    Set ObtenerTabla = ThisWorkbook.Worksheets(nombreHoja).ListObjects(nombreTabla)
End Function

Public Function SiguienteID(ByVal tabla As ListObject, ByVal columnaID As String, _
                             ByVal prefijo As String, ByVal digitos As Integer) As String
    Dim maxNum As Long, i As Long, n As Long, limite As Long
    Dim texto As String, numeroTexto As String, valores As Variant
    maxNum = 0
    If Not tabla.DataBodyRange Is Nothing Then
        valores = tabla.ListColumns(columnaID).DataBodyRange.Value2
        If IsArray(valores) Then
            limite = UBound(valores, 1)
        Else
            limite = 1
        End If
        For i = 1 To limite
            If IsArray(valores) Then
                texto = CStr(valores(i, 1))
            Else
                texto = CStr(valores)
            End If
            If Left$(texto, Len(prefijo)) = prefijo Then
                numeroTexto = Mid$(texto, Len(prefijo) + 1)
                If IsNumeric(numeroTexto) Then
                    n = CLng(numeroTexto)
                    If n > maxNum Then maxNum = n
                End If
            End If
        Next i
    End If
    SiguienteID = prefijo & Format(maxNum + 1, String(digitos, "0"))
End Function

Public Function LimpiarNombreArchivo(ByVal texto As String) As String
    Dim invalidos As String, i As Long, codigo As Long
    Dim resultado As String, base As String, punto As Long
    invalidos = "\/:*?""<>|"
    For i = 1 To Len(texto)
        codigo = AscW(Mid$(texto, i, 1))
        If codigo >= 32 And InStr(1, invalidos, Mid$(texto, i, 1), vbBinaryCompare) = 0 Then
            resultado = resultado & Mid$(texto, i, 1)
        End If
    Next i
    resultado = Replace(Trim$(resultado), " ", "_")
    Do While Len(resultado) > 0 And (Right$(resultado, 1) = "." Or Right$(resultado, 1) = " ")
        resultado = Left$(resultado, Len(resultado) - 1)
    Loop
    If resultado = "" Then resultado = "SINDATO"
    If Len(resultado) > 60 Then resultado = Left$(resultado, 60)
    Do While Len(resultado) > 0 And (Right$(resultado, 1) = "." Or Right$(resultado, 1) = " ")
        resultado = Left$(resultado, Len(resultado) - 1)
    Loop
    If resultado = "" Then resultado = "SINDATO"
    punto = InStr(1, resultado, ".", vbBinaryCompare)
    If punto > 0 Then
        base = UCase$(Left$(resultado, punto - 1))
    Else
        base = UCase$(resultado)
    End If
    Select Case base
        Case "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", _
             "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            resultado = "_" & resultado
    End Select
    LimpiarNombreArchivo = resultado
End Function

Public Function NormalizarIdentificacion(ByVal valor As String) As String
    Dim i As Long, caracter As String, codigo As Long, salida As String
    valor = UCase$(Trim$(valor))
    For i = 1 To Len(valor)
        caracter = Mid$(valor, i, 1)
        codigo = AscW(caracter)
        If (codigo >= 48 And codigo <= 57) Or (codigo >= 65 And codigo <= 90) Then
            salida = salida & caracter
        End If
    Next i
    NormalizarIdentificacion = salida
End Function

Public Function NormalizarNombreClave(ByVal valor As String) As String
    Dim i As Long, caracter As String, codigo As Long, salida As String
    valor = UCase$(Trim$(valor))
    valor = Replace(valor, "Á", "A")
    valor = Replace(valor, "É", "E")
    valor = Replace(valor, "Í", "I")
    valor = Replace(valor, "Ó", "O")
    valor = Replace(valor, "Ú", "U")
    valor = Replace(valor, "Ü", "U")
    valor = Replace(valor, "Ñ", "N")
    For i = 1 To Len(valor)
        caracter = Mid$(valor, i, 1)
        codigo = AscW(caracter)
        If (codigo >= 48 And codigo <= 57) Or (codigo >= 65 And codigo <= 90) Then
            salida = salida & caracter
        End If
    Next i
    NormalizarNombreClave = salida
End Function

Public Sub AsegurarCarpeta(ByVal ruta As String)
    Dim sistemaArchivos As Object, padre As String
    ruta = Replace(Trim$(ruta), "/", "\")
    Do While Len(ruta) > 3 And Right$(ruta, 1) = "\"
        ruta = Left$(ruta, Len(ruta) - 1)
    Loop
    If ruta = "" Then Err.Raise vbObjectError + 720, "AsegurarCarpeta", "La ruta de documentos está vacía."
    Set sistemaArchivos = CreateObject("Scripting.FileSystemObject")
    If sistemaArchivos.FolderExists(ruta) Then Exit Sub
    padre = sistemaArchivos.GetParentFolderName(ruta)
    If padre = "" Or StrComp(padre, ruta, vbTextCompare) = 0 Then _
        Err.Raise vbObjectError + 721, "AsegurarCarpeta", "No se puede crear la carpeta: " & ruta
    If Not sistemaArchivos.FolderExists(padre) Then AsegurarCarpeta padre
    sistemaArchivos.CreateFolder ruta
End Sub

Public Function ObtenerParametro(ByVal codigo As String, Optional ByVal porDefecto As String = "") As String
    Dim tabla As ListObject, i As Long
    Set tabla = ObtenerTabla("PARAMETROS", "PARAMETROS")
    ObtenerParametro = porDefecto
    If tabla.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To tabla.ListRows.Count
        If CStr(tabla.ListColumns("Parametro").DataBodyRange.Cells(i, 1).Value) = codigo Then
            ObtenerParametro = CStr(tabla.ListColumns("Valor").DataBodyRange.Cells(i, 1).Value)
            Exit Function
        End If
    Next i
End Function

Public Function ObtenerEmpresaValor(ByVal campo As String, Optional ByVal porDefecto As String = "") As String
    Dim tabla As ListObject, i As Long
    Set tabla = ObtenerTabla("EMPRESA", "EMPRESA")
    ObtenerEmpresaValor = porDefecto
    If tabla.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To tabla.ListRows.Count
        If CStr(tabla.ListColumns("Campo").DataBodyRange.Cells(i, 1).Value) = campo Then
            ObtenerEmpresaValor = CStr(tabla.ListColumns("Valor").DataBodyRange.Cells(i, 1).Value)
            Exit Function
        End If
    Next i
End Function

' Reemplaza la nota del modulo indicado y conserva las demas notas.
' Formato: [ETIQUETA] texto, separadas por " | ".
Public Sub EstablecerNotaMarcada(ByVal tabla As ListObject, ByVal indiceFila As Long, _
                                 ByVal etiqueta As String, ByVal mensaje As String)
    Dim celda As Range, segmentos As Variant, segmento As Variant
    Dim marca As String, actual As String, salida As String, texto As String
    If Len(Trim$(etiqueta)) = 0 Then Exit Sub

    Set celda = tabla.ListColumns("Observaciones").DataBodyRange.Cells(indiceFila, 1)
    marca = "[" & UCase$(Trim$(etiqueta)) & "]"
    actual = CStr(celda.Value)
    segmentos = Split(actual, " | ")
    For Each segmento In segmentos
        texto = Trim$(CStr(segmento))
        If texto <> "" And UCase$(Left$(texto, Len(marca))) <> marca Then
            If salida <> "" Then salida = salida & " | "
            salida = salida & texto
        End If
    Next segmento

    texto = Replace(Replace(Trim$(mensaje), vbCr, " "), vbLf, " ")
    texto = Replace(texto, " | ", " / ")
    If texto <> "" Then
        If salida <> "" Then salida = salida & " | "
        salida = salida & marca & " " & texto
    End If
    celda.Value = salida
End Sub

Public Sub EscribirLog(ByVal accion As String, ByVal modulo As String, ByVal registroAfectado As String, _
                        Optional ByVal valorAnterior As String = "", Optional ByVal valorNuevo As String = "", _
                        Optional ByVal resultado As String = "OK")
    Static avisoMostrado As Boolean
    Dim tabla As ListObject, fila As ListRow
    Dim numeroError As Long, descripcionError As String
    On Error Resume Next
    Err.Clear
    Set tabla = ObtenerTabla("LOG", "LOG")
    Set fila = tabla.ListRows.Add(AlwaysInsert:=True)
    With fila.Range
        .Cells(1, tabla.ListColumns("Fecha").Index).Value = Date
        .Cells(1, tabla.ListColumns("Hora").Index).Value = Format(Now, "hh:mm:ss")
        .Cells(1, tabla.ListColumns("Usuario").Index).Value = Application.UserName
        .Cells(1, tabla.ListColumns("Accion").Index).Value = accion
        .Cells(1, tabla.ListColumns("Modulo").Index).Value = modulo
        .Cells(1, tabla.ListColumns("RegistroAfectado").Index).Value = registroAfectado
        .Cells(1, tabla.ListColumns("ValorAnterior").Index).Value = valorAnterior
        .Cells(1, tabla.ListColumns("ValorNuevo").Index).Value = valorNuevo
        .Cells(1, tabla.ListColumns("Resultado").Index).Value = resultado
    End With
    numeroError = Err.Number
    descripcionError = Err.Description
    Err.Clear
    On Error GoTo 0
    If numeroError <> 0 Then
        Debug.Print "LOG no escrito (" & numeroError & "): " & descripcionError
        If Not avisoMostrado Then
            avisoMostrado = True
            MsgBox "No se pudo escribir en LOG. La acción pudo completarse sin auditoría: " & _
                   descripcionError, vbExclamation, "Aviso de auditoría"
        End If
    End If
End Sub

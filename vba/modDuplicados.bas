Attribute VB_Name = "modDuplicados"
Option Explicit

' Deteccion de posibles duplicados. Nunca elimina ni bloquea nada: solo
' devuelve el ID de la fila con la que coincide, para que otro modulo decida
' que Estado/Observacion poner. Ver seccion 16 del requerimiento original.

Public Function BuscarDuplicadoDocumento(ByVal tabla As ListObject, ByVal indiceFila As Long) As String
    Dim i As Long
    Dim nit As String, numero As String, cufe As String, nombre As String
    Dim fecha As Variant, total As Variant
    Dim nitI As String, numeroI As String, cufeI As String, nombreI As String
    Dim fechaI As Variant, totalI As Variant, idI As String
    Dim fila As Range

    If tabla.DataBodyRange Is Nothing Then Exit Function
    If indiceFila < 1 Or indiceFila > tabla.ListRows.Count Then Exit Function
    Set fila = tabla.ListRows(indiceFila).Range
    nit = modUtils.NormalizarIdentificacion(CStr(fila.Cells(1, tabla.ListColumns("Identificacion").Index).Value))
    numero = Trim$(UCase$(CStr(fila.Cells(1, tabla.ListColumns("Numero").Index).Value)))
    cufe = Trim$(UCase$(CStr(fila.Cells(1, tabla.ListColumns("CUFE_Referencia").Index).Value)))
    nombre = modUtils.NormalizarNombreClave(CStr(fila.Cells(1, tabla.ListColumns("Tercero").Index).Value))
    fecha = fila.Cells(1, tabla.ListColumns("Fecha").Index).Value
    total = fila.Cells(1, tabla.ListColumns("Total").Index).Value

    BuscarDuplicadoDocumento = ""
    For i = 1 To tabla.ListRows.Count
        If i <> indiceFila Then
            idI = CStr(tabla.ListColumns("ID").DataBodyRange.Cells(i, 1).Value)
            If Trim$(idI) <> "" Then
                cufeI = Trim$(UCase$(CStr(tabla.ListColumns("CUFE_Referencia").DataBodyRange.Cells(i, 1).Value)))
                numeroI = Trim$(UCase$(CStr(tabla.ListColumns("Numero").DataBodyRange.Cells(i, 1).Value)))
                nombreI = modUtils.NormalizarNombreClave(CStr(tabla.ListColumns("Tercero").DataBodyRange.Cells(i, 1).Value))
                nitI = modUtils.NormalizarIdentificacion(CStr(tabla.ListColumns("Identificacion").DataBodyRange.Cells(i, 1).Value))
                fechaI = tabla.ListColumns("Fecha").DataBodyRange.Cells(i, 1).Value
                totalI = tabla.ListColumns("Total").DataBodyRange.Cells(i, 1).Value

                If cufe <> "" And cufe = cufeI Then
                    BuscarDuplicadoDocumento = idI
                    Exit Function
                ElseIf nit <> "" And numero <> "" And nit = nitI And numero = numeroI Then
                    BuscarDuplicadoDocumento = idI
                    Exit Function
                End If
                If nit <> "" And nit = nitI Then
                    If IsDate(fecha) And IsDate(fechaI) Then
                        If IsNumeric(total) And IsNumeric(totalI) Then
                            If CDate(fecha) = CDate(fechaI) And Abs(CDbl(total) - CDbl(totalI)) <= 0.01 Then
                                BuscarDuplicadoDocumento = idI
                                Exit Function
                            End If
                        End If
                    End If
                End If
                If Len(nombre) >= 4 And numero <> "" And numero = numeroI And nombre = nombreI Then
                    If IsNumeric(total) And IsNumeric(totalI) Then
                        If Abs(CDbl(total) - CDbl(totalI)) <= 0.01 Then
                            BuscarDuplicadoDocumento = idI
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Function

Public Function BuscarDuplicadoTercero(ByVal tabla As ListObject, ByVal indiceFila As Long) As String
    Dim i As Long, clave As String, claveI As String, idI As String
    Dim nombre As String, nombreI As String

    If tabla.DataBodyRange Is Nothing Then Exit Function
    If indiceFila < 1 Or indiceFila > tabla.ListRows.Count Then Exit Function
    clave = modUtils.NormalizarIdentificacion(CStr(tabla.ListColumns("NIT_ID").DataBodyRange.Cells(indiceFila, 1).Value))
    nombre = modUtils.NormalizarNombreClave(CStr(tabla.ListColumns("NombreRazonSocial").DataBodyRange.Cells(indiceFila, 1).Value))
    BuscarDuplicadoTercero = ""
    If clave = "" And Len(nombre) < 4 Then Exit Function

    For i = 1 To tabla.ListRows.Count
        If i <> indiceFila Then
            idI = CStr(tabla.ListColumns("ID").DataBodyRange.Cells(i, 1).Value)
            If Trim$(idI) <> "" Then
                claveI = modUtils.NormalizarIdentificacion(CStr(tabla.ListColumns("NIT_ID").DataBodyRange.Cells(i, 1).Value))
                nombreI = modUtils.NormalizarNombreClave(CStr(tabla.ListColumns("NombreRazonSocial").DataBodyRange.Cells(i, 1).Value))
                If clave <> "" And claveI = clave Then
                    BuscarDuplicadoTercero = idI
                    Exit Function
                ElseIf Len(nombre) >= 4 And nombreI = nombre Then
                    BuscarDuplicadoTercero = idI
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' Índices para validaciones por lotes: O(n) en lugar de comparar cada fila
' contra todas las demás (O(n²)). Los índices guardan las dos primeras filas
' para que ambas queden señaladas como posibles duplicados.
Public Sub CrearIndicesDocumentos(ByRef indiceCUFE As Object, _
                                  ByRef indiceNitNumero As Object, _
                                  ByRef indiceNitFechaTotal As Object, _
                                  ByRef indiceNombreNumeroTotal As Object)
    Set indiceCUFE = CreateObject("Scripting.Dictionary")
    Set indiceNitNumero = CreateObject("Scripting.Dictionary")
    Set indiceNitFechaTotal = CreateObject("Scripting.Dictionary")
    Set indiceNombreNumeroTotal = CreateObject("Scripting.Dictionary")
    indiceCUFE.CompareMode = vbTextCompare
    indiceNitNumero.CompareMode = vbTextCompare
    indiceNitFechaTotal.CompareMode = vbTextCompare
    indiceNombreNumeroTotal.CompareMode = vbTextCompare
End Sub

Public Sub AgregarDocumentoAIndices(ByVal tabla As ListObject, ByVal indiceFila As Long, _
                                    ByVal indiceCUFE As Object, _
                                    ByVal indiceNitNumero As Object, _
                                    ByVal indiceNitFechaTotal As Object, _
                                    ByVal indiceNombreNumeroTotal As Object)
    Dim fila As Range, id As String, nit As String, numero As String, nombre As String
    Dim cufe As String, fecha As Variant, total As Variant
    Set fila = tabla.ListRows(indiceFila).Range
    id = Trim$(CStr(fila.Cells(1, tabla.ListColumns("ID").Index).Value))
    If id = "" Then Exit Sub

    nit = modUtils.NormalizarIdentificacion(CStr(fila.Cells(1, tabla.ListColumns("Identificacion").Index).Value))
    numero = Trim$(UCase$(CStr(fila.Cells(1, tabla.ListColumns("Numero").Index).Value)))
    nombre = modUtils.NormalizarNombreClave(CStr(fila.Cells(1, tabla.ListColumns("Tercero").Index).Value))
    cufe = Trim$(UCase$(CStr(fila.Cells(1, tabla.ListColumns("CUFE_Referencia").Index).Value)))
    fecha = fila.Cells(1, tabla.ListColumns("Fecha").Index).Value
    total = fila.Cells(1, tabla.ListColumns("Total").Index).Value

    AgregarClave indiceCUFE, cufe, id
    If nit <> "" And numero <> "" Then AgregarClave indiceNitNumero, nit & "|" & numero, id
    If nit <> "" And IsDate(fecha) And IsNumeric(total) Then
        AgregarClave indiceNitFechaTotal, nit & "|" & Format$(CDate(fecha), "yyyymmdd") & _
                     "|" & Format$(CDbl(total), "0.############"), id
    End If
    If Len(nombre) >= 4 And numero <> "" And IsNumeric(total) Then
        AgregarClave indiceNombreNumeroTotal, nombre & "|" & numero & _
                     "|" & Format$(CDbl(total), "0.############"), id
    End If
End Sub

Public Function BuscarDuplicadoDocumentoEnIndices(ByVal tabla As ListObject, _
                                                   ByVal indiceFila As Long, _
                                                   ByVal indiceCUFE As Object, _
                                                   ByVal indiceNitNumero As Object, _
                                                   ByVal indiceNitFechaTotal As Object, _
                                                   ByVal indiceNombreNumeroTotal As Object) As String
    Dim fila As Range, id As String, nit As String, numero As String, nombre As String
    Dim cufe As String, fecha As Variant, total As Variant, clave As String
    If tabla.DataBodyRange Is Nothing Then Exit Function
    If indiceFila < 1 Or indiceFila > tabla.ListRows.Count Then Exit Function
    Set fila = tabla.ListRows(indiceFila).Range
    id = Trim$(CStr(fila.Cells(1, tabla.ListColumns("ID").Index).Value))
    nit = modUtils.NormalizarIdentificacion(CStr(fila.Cells(1, tabla.ListColumns("Identificacion").Index).Value))
    numero = Trim$(UCase$(CStr(fila.Cells(1, tabla.ListColumns("Numero").Index).Value)))
    nombre = modUtils.NormalizarNombreClave(CStr(fila.Cells(1, tabla.ListColumns("Tercero").Index).Value))
    cufe = Trim$(UCase$(CStr(fila.Cells(1, tabla.ListColumns("CUFE_Referencia").Index).Value)))
    fecha = fila.Cells(1, tabla.ListColumns("Fecha").Index).Value
    total = fila.Cells(1, tabla.ListColumns("Total").Index).Value

    If cufe <> "" Then
        clave = OtraID(indiceCUFE, cufe, id)
        If clave <> "" Then
            BuscarDuplicadoDocumentoEnIndices = clave
            Exit Function
        End If
    End If
    If nit <> "" And numero <> "" Then
        clave = OtraID(indiceNitNumero, nit & "|" & numero, id)
        If clave <> "" Then
            BuscarDuplicadoDocumentoEnIndices = clave
            Exit Function
        End If
    End If
    If nit <> "" And IsDate(fecha) And IsNumeric(total) Then
        clave = OtraID(indiceNitFechaTotal, nit & "|" & Format$(CDate(fecha), "yyyymmdd") & _
                       "|" & Format$(CDbl(total), "0.############"), id)
        If clave <> "" Then
            BuscarDuplicadoDocumentoEnIndices = clave
            Exit Function
        End If
    End If
    If Len(nombre) >= 4 And numero <> "" And IsNumeric(total) Then
        clave = OtraID(indiceNombreNumeroTotal, nombre & "|" & numero & _
                       "|" & Format$(CDbl(total), "0.############"), id)
        BuscarDuplicadoDocumentoEnIndices = clave
    End If
End Function

Public Sub CrearIndiceTerceros(ByRef indiceNIT As Object, ByRef indiceNombre As Object)
    Set indiceNIT = CreateObject("Scripting.Dictionary")
    Set indiceNombre = CreateObject("Scripting.Dictionary")
    indiceNIT.CompareMode = vbTextCompare
    indiceNombre.CompareMode = vbTextCompare
End Sub

Public Sub AgregarTerceroAIndice(ByVal tabla As ListObject, ByVal indiceFila As Long, _
                                 ByVal indiceNIT As Object, ByVal indiceNombre As Object)
    Dim fila As Range, id As String, nit As String, nombre As String
    Set fila = tabla.ListRows(indiceFila).Range
    id = Trim$(CStr(fila.Cells(1, tabla.ListColumns("ID").Index).Value))
    nit = modUtils.NormalizarIdentificacion(CStr(fila.Cells(1, tabla.ListColumns("NIT_ID").Index).Value))
    nombre = modUtils.NormalizarNombreClave(CStr(fila.Cells(1, tabla.ListColumns("NombreRazonSocial").Index).Value))
    If id <> "" And nit <> "" Then AgregarClave indiceNIT, nit, id
    If id <> "" And Len(nombre) >= 4 Then AgregarClave indiceNombre, nombre, id
End Sub

Public Function BuscarDuplicadoTerceroEnIndice(ByVal tabla As ListObject, _
                                                ByVal indiceFila As Long, _
                                                ByVal indiceNIT As Object, _
                                                ByVal indiceNombre As Object) As String
    Dim fila As Range, id As String, nit As String, nombre As String
    If tabla.DataBodyRange Is Nothing Then Exit Function
    If indiceFila < 1 Or indiceFila > tabla.ListRows.Count Then Exit Function
    Set fila = tabla.ListRows(indiceFila).Range
    id = Trim$(CStr(fila.Cells(1, tabla.ListColumns("ID").Index).Value))
    nit = modUtils.NormalizarIdentificacion(CStr(fila.Cells(1, tabla.ListColumns("NIT_ID").Index).Value))
    If nit <> "" Then BuscarDuplicadoTerceroEnIndice = OtraID(indiceNIT, nit, id)
    If BuscarDuplicadoTerceroEnIndice = "" Then
        nombre = modUtils.NormalizarNombreClave(CStr(fila.Cells(1, tabla.ListColumns("NombreRazonSocial").Index).Value))
        If Len(nombre) >= 4 Then BuscarDuplicadoTerceroEnIndice = OtraID(indiceNombre, nombre, id)
    End If
End Function

Private Sub AgregarClave(ByVal indice As Object, ByVal clave As String, ByVal id As String)
    Dim ids As Variant
    If clave = "" Or id = "" Then Exit Sub
    If Not indice.Exists(clave) Then
        indice.Add clave, Array(id, "")
    Else
        ids = indice(clave)
        If CStr(ids(1)) = "" And CStr(ids(0)) <> id Then
            ids(1) = id
            indice(clave) = ids
        End If
    End If
End Sub

Private Function OtraID(ByVal indice As Object, ByVal clave As String, _
                        ByVal idActual As String) As String
    Dim ids As Variant
    If clave = "" Or Not indice.Exists(clave) Then Exit Function
    ids = indice(clave)
    If CStr(ids(0)) <> idActual Then
        OtraID = CStr(ids(0))
    ElseIf CStr(ids(1)) <> "" Then
        OtraID = CStr(ids(1))
    End If
End Function

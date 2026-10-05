Attribute VB_Name = "modDocumentos"
Option Explicit

' Acciones sobre DOCUMENTOS: alta, validacion y deteccion de duplicados.

Public Sub NuevoDocumento()
    On Error GoTo ManejarError
    Dim tabla As ListObject, fila As ListRow, id As String

    Set tabla = modUtils.ObtenerTabla("DOCUMENTOS", "DOCUMENTOS")
    id = modUtils.SiguienteID(tabla, "ID", modUtils.ObtenerParametro("PREFIJO_DOCUMENTO", "DOC-"), 5)
    Set fila = tabla.ListRows.Add(AlwaysInsert:=True)
    With fila.Range
        .Cells(1, tabla.ListColumns("ID").Index).Value = id
        .Cells(1, tabla.ListColumns("Fecha").Index).Value = Date
        .Cells(1, tabla.ListColumns("Estado").Index).Value = "Pendiente"
        .Cells(1, tabla.ListColumns("MetodoExtraccion").Index).Value = "Manual"
    End With
    modUtils.EscribirLog "Crear", "DOCUMENTOS", id, "", "Fila creada", "OK"
    Application.Goto fila.Range.Cells(1, tabla.ListColumns("TipoDocumento").Index), True
    Exit Sub

ManejarError:
    MsgBox "No se pudo crear el documento: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub ValidarDocumentos()
    On Error GoTo ManejarError
    Dim tabla As ListObject, i As Long, idFila As String, msg As String, dupId As String
    Dim estadoAnterior As String, estadoNuevo As String
    Dim contadorError As Long, contadorDup As Long, contadorOk As Long, contadorRevision As Long
    Dim indiceCUFE As Object, indiceNitNumero As Object, indiceNitFechaTotal As Object
    Dim indiceNombreNumeroTotal As Object

    Set tabla = modUtils.ObtenerTabla("DOCUMENTOS", "DOCUMENTOS")
    If tabla.DataBodyRange Is Nothing Then
        MsgBox "No hay documentos para validar.", vbInformation
        Exit Sub
    End If

    modDuplicados.CrearIndicesDocumentos indiceCUFE, indiceNitNumero, indiceNitFechaTotal, _
                                         indiceNombreNumeroTotal
    For i = 1 To tabla.ListRows.Count
        modDuplicados.AgregarDocumentoAIndices tabla, i, _
            indiceCUFE, indiceNitNumero, indiceNitFechaTotal, indiceNombreNumeroTotal
    Next i

    For i = 1 To tabla.ListRows.Count
        idFila = CStr(tabla.ListColumns("ID").DataBodyRange.Cells(i, 1).Value)
        If Trim$(idFila) <> "" Then
            estadoAnterior = CStr(tabla.ListColumns("Estado").DataBodyRange.Cells(i, 1).Value)
            msg = modValidation.ValidarDocumentoFila(tabla, i)
            If msg <> "" Then
                estadoNuevo = "Error"
                modUtils.EstablecerNotaMarcada tabla, i, "VALIDACION", msg
                modUtils.EstablecerNotaMarcada tabla, i, "DUPLICADO", ""
                contadorError = contadorError + 1
            Else
                modUtils.EstablecerNotaMarcada tabla, i, "VALIDACION", ""
                dupId = modDuplicados.BuscarDuplicadoDocumentoEnIndices( _
                    tabla, i, indiceCUFE, indiceNitNumero, indiceNitFechaTotal, _
                    indiceNombreNumeroTotal)
                If dupId <> "" Then
                    estadoNuevo = "Duplicado"
                    modUtils.EstablecerNotaMarcada tabla, i, "DUPLICADO", _
                        "Posible duplicado de " & dupId & "; revisar antes de continuar."
                    contadorDup = contadorDup + 1
                ElseIf InStr(1, CStr(tabla.ListColumns("Observaciones").DataBodyRange.Cells(i, 1).Value), _
                             "[XML_REVISION]", vbTextCompare) > 0 Then
                    estadoNuevo = "Revision Manual"
                    modUtils.EstablecerNotaMarcada tabla, i, "DUPLICADO", ""
                    contadorRevision = contadorRevision + 1
                Else
                    estadoNuevo = "Validado"
                    modUtils.EstablecerNotaMarcada tabla, i, "DUPLICADO", ""
                    contadorOk = contadorOk + 1
                End If
            End If
            If estadoNuevo <> estadoAnterior Then
                tabla.ListColumns("Estado").DataBodyRange.Cells(i, 1).Value = estadoNuevo
                modUtils.EscribirLog "Validar", "DOCUMENTOS", idFila, estadoAnterior, estadoNuevo, "OK"
            End If
        End If
    Next i

    MsgBox "Validacion de DOCUMENTOS completada." & vbCrLf & _
           "Validados: " & contadorOk & vbCrLf & _
           "Posibles duplicados: " & contadorDup & vbCrLf & _
           "Pendientes de revision humana: " & contadorRevision & vbCrLf & _
           "Con error: " & contadorError, vbInformation
    Exit Sub

ManejarError:
    MsgBox "Error validando documentos: " & Err.Description, vbCritical, "Error"
End Sub

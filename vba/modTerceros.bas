Attribute VB_Name = "modTerceros"
Option Explicit

' Acciones sobre la tabla TERCEROS: alta con ID automatico y validacion de
' NIT/ID duplicado (ver seccion 9 del requerimiento: "evita duplicar terceros").

Public Sub NuevoTercero()
    On Error GoTo ManejarError
    Dim tabla As ListObject, fila As ListRow, id As String

    Set tabla = modUtils.ObtenerTabla("TERCEROS", "TERCEROS")
    id = modUtils.SiguienteID(tabla, "ID", "TER-", 5)

    Set fila = tabla.ListRows.Add(AlwaysInsert:=True)
    With fila.Range
        .Cells(1, tabla.ListColumns("ID").Index).Value = id
        .Cells(1, tabla.ListColumns("Estado").Index).Value = "Activo"
    End With

    modUtils.EscribirLog "Crear", "TERCEROS", id, "", "Fila creada", "OK"
    Application.Goto fila.Range.Cells(1, tabla.ListColumns("Tipo").Index), True
    Exit Sub

ManejarError:
    MsgBox "No se pudo crear el tercero: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub ValidarTerceros()
    On Error GoTo ManejarError
    Dim tabla As ListObject, i As Long, idFila As String, dupId As String, msg As String
    Dim contadorDup As Long, contadorError As Long, indiceNIT As Object, indiceNombre As Object

    Set tabla = modUtils.ObtenerTabla("TERCEROS", "TERCEROS")
    If tabla.DataBodyRange Is Nothing Then
        MsgBox "No hay terceros para validar.", vbInformation
        Exit Sub
    End If

    modDuplicados.CrearIndiceTerceros indiceNIT, indiceNombre
    For i = 1 To tabla.ListRows.Count
        modDuplicados.AgregarTerceroAIndice tabla, i, indiceNIT, indiceNombre
    Next i

    For i = 1 To tabla.ListRows.Count
        idFila = CStr(tabla.ListColumns("ID").DataBodyRange.Cells(i, 1).Value)
        If Trim$(idFila) <> "" Then
            msg = modValidation.ValidarTerceroFila(tabla, i)
            If msg <> "" Then
                modUtils.EstablecerNotaMarcada tabla, i, "VALIDACION", msg
                contadorError = contadorError + 1
            Else
                modUtils.EstablecerNotaMarcada tabla, i, "VALIDACION", ""
            End If
            dupId = modDuplicados.BuscarDuplicadoTerceroEnIndice(tabla, i, indiceNIT, indiceNombre)
            If dupId <> "" Then
                modUtils.EstablecerNotaMarcada tabla, i, "DUPLICADO", _
                    "Posible coincidencia de NIT/ID o nombre con " & dupId & "; revisar antes de continuar."
                contadorDup = contadorDup + 1
            Else
                modUtils.EstablecerNotaMarcada tabla, i, "DUPLICADO", ""
            End If
        End If
    Next i

    MsgBox "Validacion de TERCEROS completada." & vbCrLf & _
           "Con datos obligatorios incompletos: " & contadorError & vbCrLf & _
           "Posibles duplicados: " & contadorDup, vbInformation
    Exit Sub

ManejarError:
    MsgBox "Error validando terceros: " & Err.Description, vbCritical, "Error"
End Sub

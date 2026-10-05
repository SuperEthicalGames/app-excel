Attribute VB_Name = "modMovimientos"
Option Explicit

' Acciones sobre la tabla MOVIMIENTOS: alta con ID automatico y validacion masiva.
' La logica de "que es un movimiento valido" vive en modValidation, no aqui.

Public Sub NuevoMovimiento()
    On Error GoTo ManejarError
    Dim tabla As ListObject, fila As ListRow, id As String

    Set tabla = modUtils.ObtenerTabla("MOVIMIENTOS", "MOVIMIENTOS")
    id = modUtils.SiguienteID(tabla, "ID", modUtils.ObtenerParametro("PREFIJO_MOVIMIENTO", "MOV-"), 5)

    Set fila = tabla.ListRows.Add(AlwaysInsert:=True)
    With fila.Range
        .Cells(1, tabla.ListColumns("ID").Index).Value = id
        .Cells(1, tabla.ListColumns("Fecha").Index).Value = Date
        .Cells(1, tabla.ListColumns("Estado").Index).Value = "Pendiente"
        .Cells(1, tabla.ListColumns("Subtotal").Index).Value = 0
        .Cells(1, tabla.ListColumns("Impuestos").Index).Value = 0
        .Cells(1, tabla.ListColumns("Retenciones").Index).Value = 0
    End With

    modUtils.EscribirLog "Crear", "MOVIMIENTOS", id, "", "Fila creada", "OK"
    Application.Goto fila.Range.Cells(1, tabla.ListColumns("Tipo").Index), True
    Exit Sub

ManejarError:
    MsgBox "No se pudo crear el movimiento: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub ValidarMovimientos()
    On Error GoTo ManejarError
    Dim tabla As ListObject, i As Long, idFila As String, msg As String
    Dim contadorError As Long, contadorOk As Long

    Set tabla = modUtils.ObtenerTabla("MOVIMIENTOS", "MOVIMIENTOS")
    If tabla.DataBodyRange Is Nothing Then
        MsgBox "No hay movimientos para validar.", vbInformation
        Exit Sub
    End If

    For i = 1 To tabla.ListRows.Count
        idFila = CStr(tabla.ListColumns("ID").DataBodyRange.Cells(i, 1).Value)
        If Trim$(idFila) <> "" Then
            msg = modValidation.ValidarMovimientoFila(tabla, i)
            If msg <> "" Then
                modUtils.EstablecerNotaMarcada tabla, i, "VALIDACION", msg
                modUtils.EscribirLog "Validar", "MOVIMIENTOS", idFila, "", msg, "ERROR"
                contadorError = contadorError + 1
            Else
                modUtils.EstablecerNotaMarcada tabla, i, "VALIDACION", ""
                contadorOk = contadorOk + 1
            End If
        End If
    Next i

    MsgBox "Validacion de MOVIMIENTOS completada." & vbCrLf & _
           "Correctos: " & contadorOk & vbCrLf & "Con error: " & contadorError, vbInformation
    Exit Sub

ManejarError:
    MsgBox "Error validando movimientos: " & Err.Description, vbCritical, "Error"
End Sub

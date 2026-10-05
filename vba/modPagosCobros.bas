Attribute VB_Name = "modPagosCobros"
Option Explicit

' Alta de filas en PAGOS y COBROS (los sub-libros de caja contra MOVIMIENTOS).
' Simetricos a proposito: mismo patron que modMovimientos/modDocumentos.

Public Sub NuevoPago()
    On Error GoTo ManejarError
    Dim tabla As ListObject, fila As ListRow, id As String

    Set tabla = modUtils.ObtenerTabla("PAGOS", "PAGOS")
    id = modUtils.SiguienteID(tabla, "ID", "PAG-", 5)

    Set fila = tabla.ListRows.Add(AlwaysInsert:=True)
    With fila.Range
        .Cells(1, tabla.ListColumns("ID").Index).Value = id
        .Cells(1, tabla.ListColumns("Fecha").Index).Value = Date
        .Cells(1, tabla.ListColumns("Estado").Index).Value = "Pendiente"
    End With

    modUtils.EscribirLog "Crear", "PAGOS", id, "", "Fila creada", "OK"
    Application.Goto fila.Range.Cells(1, tabla.ListColumns("MovimientoID").Index), True
    Exit Sub

ManejarError:
    MsgBox "No se pudo crear el pago: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub NuevoCobro()
    On Error GoTo ManejarError
    Dim tabla As ListObject, fila As ListRow, id As String

    Set tabla = modUtils.ObtenerTabla("COBROS", "COBROS")
    id = modUtils.SiguienteID(tabla, "ID", "COB-", 5)

    Set fila = tabla.ListRows.Add(AlwaysInsert:=True)
    With fila.Range
        .Cells(1, tabla.ListColumns("ID").Index).Value = id
        .Cells(1, tabla.ListColumns("Fecha").Index).Value = Date
        .Cells(1, tabla.ListColumns("Estado").Index).Value = "Pendiente"
    End With

    modUtils.EscribirLog "Crear", "COBROS", id, "", "Fila creada", "OK"
    Application.Goto fila.Range.Cells(1, tabla.ListColumns("MovimientoID").Index), True
    Exit Sub

ManejarError:
    MsgBox "No se pudo crear el cobro: " & Err.Description, vbCritical, "Error"
End Sub

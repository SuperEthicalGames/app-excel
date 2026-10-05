Attribute VB_Name = "modReportes"
Option Explicit

' Poblado de CXC/CXP (Fase 1 las dejo como guia manual; aqui se automatizan)
' y refresco general de reportes.

Public Sub ActualizarCXC()
    ActualizarPendientes "CXC", "Ingreso", "Venta"
End Sub

Public Sub ActualizarCXP()
    ActualizarPendientes "CXP", "Gasto", "Compra"
End Sub

Private Sub ActualizarPendientes(ByVal nombreHoja As String, ByVal tipoA As String, ByVal tipoB As String)
    Dim movTabla As ListObject, ws As Worksheet
    Dim i As Long, filaDestino As Long, ultimaFila As Long
    Dim tipoMov As String, estadoMov As String

    Set movTabla = modUtils.ObtenerTabla("MOVIMIENTOS", "MOVIMIENTOS")
    Set ws = ThisWorkbook.Worksheets(nombreHoja)

    ' La fila 4 es encabezado y la 5 es la nota; los datos empiezan en la 6.
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultimaFila >= 6 Then ws.Range("A6:G" & ultimaFila).ClearContents

    filaDestino = 6
    If Not movTabla.DataBodyRange Is Nothing Then
        For i = 1 To movTabla.ListRows.Count
            tipoMov = CStr(movTabla.ListColumns("Tipo").DataBodyRange.Cells(i, 1).Value)
            estadoMov = CStr(movTabla.ListColumns("Estado").DataBodyRange.Cells(i, 1).Value)
            If (tipoMov = tipoA Or tipoMov = tipoB) And estadoMov = "Pendiente" Then
                ws.Cells(filaDestino, 1).Value = movTabla.ListColumns("ID").DataBodyRange.Cells(i, 1).Value
                ws.Cells(filaDestino, 2).Value = movTabla.ListColumns("Fecha").DataBodyRange.Cells(i, 1).Value
                ws.Cells(filaDestino, 2).NumberFormat = "dd/mm/yyyy"
                ws.Cells(filaDestino, 3).Value = movTabla.ListColumns("Tercero").DataBodyRange.Cells(i, 1).Value
                ws.Cells(filaDestino, 4).Value = movTabla.ListColumns("Documento").DataBodyRange.Cells(i, 1).Value
                ws.Cells(filaDestino, 5).Value = movTabla.ListColumns("Total").DataBodyRange.Cells(i, 1).Value
                ws.Cells(filaDestino, 5).NumberFormat = """$"" #,##0"
                If IsDate(movTabla.ListColumns("Fecha").DataBodyRange.Cells(i, 1).Value) Then
                    ws.Cells(filaDestino, 6).Value = Date - CDate(movTabla.ListColumns("Fecha").DataBodyRange.Cells(i, 1).Value)
                End If
                ws.Cells(filaDestino, 7).Value = estadoMov
                filaDestino = filaDestino + 1
            End If
        Next i
    End If
End Sub

Public Sub ActualizarTodo()
    Dim actualizacionPantalla As Boolean
    On Error GoTo ManejarError
    actualizacionPantalla = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.Calculate
    ActualizarCXC
    ActualizarCXP
    Application.ScreenUpdating = actualizacionPantalla
    MsgBox "Reportes actualizados (Dashboard, CXC, CXP).", vbInformation
    Exit Sub

ManejarError:
    Application.ScreenUpdating = actualizacionPantalla
    MsgBox "Error actualizando reportes: " & Err.Description, vbCritical, "Error"
End Sub

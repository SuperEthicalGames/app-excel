Attribute VB_Name = "modValidation"
Option Explicit

' Validaciones puras de una fila: devuelven mensajes sin escribir ni cambiar estado.

Public Function ValidarMovimientoFila(ByVal tabla As ListObject, ByVal indiceFila As Long) As String
    Dim msg As String, fila As Range, valor As Variant
    Set fila = tabla.ListRows(indiceFila).Range
    valor = fila.Cells(1, tabla.ListColumns("Fecha").Index).Value
    If IsError(valor) Then
        msg = msg & "Fecha contiene un error de Excel. "
    ElseIf Not IsDate(valor) Then
        msg = msg & "Fecha invalida o vacia. "
    End If
    valor = fila.Cells(1, tabla.ListColumns("Tipo").Index).Value
    If IsError(valor) Then
        msg = msg & "Tipo contiene un error de Excel. "
    ElseIf TextoVacio(valor) Then
        msg = msg & "Tipo vacio. "
    End If
    valor = fila.Cells(1, tabla.ListColumns("Concepto").Index).Value
    If IsError(valor) Then
        msg = msg & "Concepto contiene un error de Excel. "
    ElseIf TextoVacio(valor) Then
        msg = msg & "Concepto vacio. "
    End If
    If Not IsNumeric(fila.Cells(1, tabla.ListColumns("Subtotal").Index).Value) Then msg = msg & "Subtotal no numerico. "
    If Not EsNumeroOpcional(fila.Cells(1, tabla.ListColumns("Impuestos").Index).Value) Then msg = msg & "Impuestos no numericos. "
    If Not EsNumeroOpcional(fila.Cells(1, tabla.ListColumns("Retenciones").Index).Value) Then msg = msg & "Retenciones no numericas. "
    valor = fila.Cells(1, tabla.ListColumns("Cuenta").Index).Value
    If IsError(valor) Then
        msg = msg & "Cuenta contiene un error de Excel. "
    ElseIf TextoVacio(valor) Then
        msg = msg & "Cuenta vacia. "
    End If
    ValidarMovimientoFila = Trim$(msg)
End Function

Public Function ValidarDocumentoFila(ByVal tabla As ListObject, ByVal indiceFila As Long) As String
    Dim msg As String, fila As Range, valor As Variant
    Set fila = tabla.ListRows(indiceFila).Range
    valor = fila.Cells(1, tabla.ListColumns("Fecha").Index).Value
    If IsError(valor) Then
        msg = msg & "Fecha contiene un error de Excel. "
    ElseIf Not IsDate(valor) Then
        msg = msg & "Fecha invalida o vacia. "
    End If
    valor = fila.Cells(1, tabla.ListColumns("TipoDocumento").Index).Value
    If IsError(valor) Then
        msg = msg & "TipoDocumento contiene un error de Excel. "
    ElseIf TextoVacio(valor) Then
        msg = msg & "TipoDocumento vacio. "
    End If
    valor = fila.Cells(1, tabla.ListColumns("Numero").Index).Value
    If IsError(valor) Then
        msg = msg & "Numero contiene un error de Excel. "
    ElseIf TextoVacio(valor) Then
        msg = msg & "Numero de documento vacio. "
    End If
    valor = fila.Cells(1, tabla.ListColumns("Tercero").Index).Value
    If IsError(valor) Then
        msg = msg & "Tercero contiene un error de Excel. "
    ElseIf TextoVacio(valor) Then
        msg = msg & "Tercero vacio. "
    End If
    If Not IsNumeric(fila.Cells(1, tabla.ListColumns("Subtotal").Index).Value) Then msg = msg & "Subtotal no numerico. "
    If Not EsNumeroOpcional(fila.Cells(1, tabla.ListColumns("Impuestos").Index).Value) Then msg = msg & "Impuestos no numericos. "
    If Not EsNumeroOpcional(fila.Cells(1, tabla.ListColumns("Retenciones").Index).Value) Then msg = msg & "Retenciones no numericas. "
    If Not IsNumeric(fila.Cells(1, tabla.ListColumns("Total").Index).Value) Then msg = msg & "Total no numerico. "
    ValidarDocumentoFila = Trim$(msg)
End Function

Public Function ValidarTerceroFila(ByVal tabla As ListObject, ByVal indiceFila As Long) As String
    Dim fila As Range, msg As String, valor As Variant
    Set fila = tabla.ListRows(indiceFila).Range
    valor = fila.Cells(1, tabla.ListColumns("Tipo").Index).Value
    If IsError(valor) Then
        msg = msg & "Tipo de tercero contiene un error de Excel. "
    ElseIf TextoVacio(valor) Then
        msg = msg & "Tipo de tercero vacio. "
    End If
    valor = fila.Cells(1, tabla.ListColumns("NombreRazonSocial").Index).Value
    If IsError(valor) Then
        msg = msg & "Nombre contiene un error de Excel. "
    ElseIf TextoVacio(valor) Then
        msg = msg & "Nombre o razon social vacio. "
    End If
    ValidarTerceroFila = Trim$(msg)
End Function

Private Function EsNumeroOpcional(ByVal valor As Variant) As Boolean
    If IsError(valor) Or IsNull(valor) Then Exit Function
    If IsEmpty(valor) Then
        EsNumeroOpcional = True
    ElseIf VarType(valor) = vbString And Len(Trim$(CStr(valor))) = 0 Then
        EsNumeroOpcional = True
    Else
        EsNumeroOpcional = IsNumeric(valor)
    End If
End Function

Private Function TextoVacio(ByVal valor As Variant) As Boolean
    If IsError(valor) Or IsNull(valor) Then Exit Function
    If IsEmpty(valor) Then
        TextoVacio = True
    Else
        TextoVacio = (Len(Trim$(CStr(valor))) = 0)
    End If
End Function

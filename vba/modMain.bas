Attribute VB_Name = "modMain"
Option Explicit

' Puntos de entrada para el usuario, asignables a botones en MENU.

Public Sub Menu_ImportarDocumento()
    modFiles.ImportarDocumento
End Sub

Public Sub Menu_ExtraerXMLDocumentoActivo()
    modFiles.ExtraerXMLDocumentoActivo
End Sub

Public Sub Menu_ConfirmarRevisionXMLDocumentoActivo()
    modFiles.ConfirmarRevisionXMLDocumentoActivo
End Sub

Public Sub Menu_RenombrarArchivoDocumentoActivo()
    modFiles.RenombrarArchivoDesdeFila
End Sub

Public Sub Menu_NuevoMovimiento()
    modMovimientos.NuevoMovimiento
End Sub

Public Sub Menu_NuevoDocumento()
    modDocumentos.NuevoDocumento
End Sub

Public Sub Menu_NuevoTercero()
    modTerceros.NuevoTercero
End Sub

Public Sub Menu_NuevoPago()
    modPagosCobros.NuevoPago
End Sub

Public Sub Menu_NuevoCobro()
    modPagosCobros.NuevoCobro
End Sub

Public Sub Menu_ValidarTodo()
    modMovimientos.ValidarMovimientos
    modDocumentos.ValidarDocumentos
    modTerceros.ValidarTerceros
End Sub

Public Sub Menu_ActualizarReportes()
    modReportes.ActualizarTodo
End Sub

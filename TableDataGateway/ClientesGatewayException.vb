Imports System

''' <summary>
''' Excepción personalizada para el Gateway de la tabla Clientes
''' </summary>
Public Class ClientesGatewayException
    Inherits Exception

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(mensaje As String)
        MyBase.New(mensaje)
    End Sub

    Public Sub New(mensaje As String, inner As Exception)
        MyBase.New(mensaje, inner)
    End Sub


End Class

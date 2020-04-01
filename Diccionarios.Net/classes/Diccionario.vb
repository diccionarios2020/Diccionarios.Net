Imports Microsoft.VisualBasic.CompilerServices
Imports System

<OptionText()>
 Public Class Diccionario
    Private vID_Libro As Integer

    Private vTitulo As String

    Private vIdioma As String

    Private vRuta As String

    Private vReferencia As String

    Public Property id_libro As Integer
        Get
            Return Me.vID_Libro
        End Get
        Set(ByVal value As Integer)
            Me.vID_Libro = value
        End Set
    End Property

    Public Property idioma As String
        Get
            Return Me.vIdioma
        End Get
        Set(ByVal value As String)
            Me.vIdioma = value
        End Set
    End Property

    Public Property referencia As String
        Get
            Return Me.vReferencia
        End Get
        Set(ByVal value As String)
            Me.vReferencia = value
        End Set
    End Property

    Public Property ruta As String
        Get
            Return Me.vRuta
        End Get
        Set(ByVal value As String)
            Me.vRuta = value
        End Set
    End Property

    Public Property titulo As String
        Get
            Return Me.vTitulo
        End Get
        Set(ByVal value As String)
            Me.vTitulo = value
        End Set
    End Property

    Public Sub New()
        MyBase.New()
    End Sub
End Class
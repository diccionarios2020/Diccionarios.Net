Imports Microsoft.VisualBasic.CompilerServices
Imports System

<OptionText()>
 Public Class Paginas
    Private vDirRepositorio As String

    Private vNombreDirDiccionario As String

    Private vNombreArchivoImagen As String

    Public Property dirRepositorio As String
        Get
            Return Me.vDirRepositorio
        End Get
        Set(ByVal value As String)
            Me.vDirRepositorio = value
        End Set
    End Property

    Public Property nombreArchivoImagen As String
        Get
            Return Me.vNombreArchivoImagen
        End Get
        Set(ByVal value As String)
            Me.vNombreArchivoImagen = value
        End Set
    End Property

    Public Property nombreDirDiccionario As String
        Get
            Return Me.vNombreDirDiccionario
        End Get
        Set(ByVal value As String)
            Me.vNombreDirDiccionario = value
        End Set
    End Property

    Public ReadOnly Property urlImagen As String
        Get
            Dim str As String = String.Concat(Me.dirRepositorio, Me.nombreDirDiccionario, "\", Me.nombreArchivoImagen)
            Return str
        End Get
    End Property

    Public Sub New()
        MyBase.New()
    End Sub
End Class
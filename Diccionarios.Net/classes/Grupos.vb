Imports Microsoft.VisualBasic.CompilerServices
Imports System

<OptionText()>
 Public Class Grupos
    Private vID_Grupo As Integer

    Private vNombreGrupo As String

    Private vDescripcion As String

    Public Property descripcion As String
        Get
            Return Me.vDescripcion
        End Get
        Set(ByVal value As String)
            Me.vDescripcion = value
        End Set
    End Property

    Public Property id_Grupo As Integer
        Get
            Return Me.vID_Grupo
        End Get
        Set(ByVal value As Integer)
            Me.vID_Grupo = value
        End Set
    End Property

    Public Property nombreGrupo As String
        Get
            Return Me.vNombreGrupo
        End Get
        Set(ByVal value As String)
            Me.vNombreGrupo = value
        End Set
    End Property

    Public Sub New()
        MyBase.New()
    End Sub
End Class
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.Common
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Drawing
Imports System.Resources
Imports System.Runtime.CompilerServices
Imports System.Windows.Forms

<OptionText()>
Public Class frmGrupos
    Inherits Form

    Public hayCambios As Boolean
    Public conexion As String
    Public idGrupoActual As Integer
    Public idDiccionarioActual As Integer

    Public Sub New()
        MyBase.New()
        Me.hayCambios = False
        Me.InitializeComponent()
    End Sub


    Private Sub cargarCbGrupos(ByRef cbGrupos As ComboBox)
        Dim oleDbConnection As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(Me.conexion)
        Dim dataSet As System.Data.DataSet = New System.Data.DataSet()
        Dim oleDbDataAdapter As System.Data.OleDb.OleDbDataAdapter = New System.Data.OleDb.OleDbDataAdapter()
        Try
            Try
                oleDbConnection.Open()
                oleDbDataAdapter.SelectCommand = New OleDbCommand("SELECT * FROM Grupos", oleDbConnection)
                oleDbDataAdapter.Fill(dataSet)
                cbGrupos.DataSource = dataSet.Tables(0)
                cbGrupos.DisplayMember = "nombreGrupo"
                cbGrupos.ValueMember = "ID_Grupo"
            Catch exception1 As System.Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As System.Exception = exception1
                Dim number As Integer = Information.Err().Number
                MessageBox.Show(String.Concat("error 1", number.ToString(), "" & vbCrLf & "", exception.Message))
                ProjectData.ClearProjectError()
            End Try
        Finally
            If ((oleDbConnection.State And ConnectionState.Open) <> ConnectionState.Closed) Then
                oleDbConnection.Close()
            End If
        End Try
    End Sub

    Private Sub cargarDiccionariosEnElGrupo(ByVal unaIdDiccionario As Integer, ByRef dgv As DataGridView)
        Dim oleDbConnection As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(Me.conexion)
        Dim dataSet As System.Data.DataSet = New System.Data.DataSet()
        Dim oleDbDataAdapter As System.Data.OleDb.OleDbDataAdapter = New System.Data.OleDb.OleDbDataAdapter()
        Try
            Try
                oleDbConnection.Open()
                Dim str As String = String.Concat("SELECT Libros.referencia, Libros.ID_Libro FROM LibrosEnGrupo AS LeC INNER JOIN Libros ON LeC.ID_Libro = Libros.ID_Libro WHERE LeC.ID_Grupo = ", Conversions.ToString(unaIdDiccionario))
                oleDbDataAdapter.SelectCommand = New OleDbCommand(str, oleDbConnection)
                oleDbDataAdapter.Fill(dataSet)
                dgv.DataSource = dataSet.Tables(0)
            Catch exception1 As System.Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As System.Exception = exception1
                Dim number As Integer = Information.Err().Number
                MessageBox.Show(String.Concat("error 2", number.ToString(), "" & vbCrLf & "", exception.Message))
                ProjectData.ClearProjectError()
            End Try
        Finally
            If ((oleDbConnection.State And ConnectionState.Open) <> ConnectionState.Closed) Then
                oleDbConnection.Close()
            End If
        End Try
    End Sub


    Private Sub darFormatoCbGrupos()
        Me.cbGrupos.Font = New System.Drawing.Font("Times New Roman", 10.0!, FontStyle.Regular)
    End Sub

    Private Sub darFormatoDgvGrupos()
        Dim font As DataGridView = Me.dgvDiccionarios
        font.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        font.Columns(0).HeaderText = "Referencia bibliográfica"
        font.ColumnHeadersVisible = False
        font.Columns(0).DefaultCellStyle.Font = New System.Drawing.Font("Times New Roman", 9.0!, FontStyle.Regular)
        font.Columns(1).Visible = False
        font = Nothing
    End Sub

    Private Sub dgvDiccionarios_DoubleClick(ByVal sender As Object, ByVal e As EventArgs)
        Me.hayCambios = True
        Me.idGrupoActual = Conversions.ToInteger(Me.cbGrupos.SelectedValue)
        Me.idDiccionarioActual = Conversions.ToInteger(Me.dgvDiccionarios.SelectedCells(1).Value)
        Me.Close()
    End Sub

    Private Sub frmGrupos_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs)
        Me.hayCambios = False
    End Sub

    ''' <summary>
    ''' OK
    ''' </summary>
    Private Sub btnAceptar_Click(sender As Object, e As System.EventArgs) Handles btnAceptar.Click
        If (Microsoft.VisualBasic.CompilerServices.Operators.ConditionalCompareObjectNotEqual(Me.cbGrupos.SelectedValue, 0, True)) Then
            Me.hayCambios = True
            Me.idGrupoActual = Conversions.ToInteger(Me.cbGrupos.SelectedValue)
        End If
        Me.Close()
    End Sub

    ''' <summary>
    ''' Cancel
    ''' </summary>
    Private Sub btnCancelar_Click(sender As Object, e As System.EventArgs) Handles btnCancelar.Click
        Me.hayCambios = False
        Me.Close()
    End Sub

    ''' <summary>
    ''' Load form
    ''' </summary>
    Private Sub frmGrupos_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.darFormatoCbGrupos()

        Me.cargarCbGrupos(Me.cbGrupos)

        Me.cbGrupos.SelectedIndex = 0
        Dim [integer] As Integer = Conversions.ToInteger(Me.cbGrupos.SelectedValue)

        Me.cargarDiccionariosEnElGrupo([integer], Me.dgvDiccionarios)

        Me.darFormatoDgvGrupos()

    End Sub

    ''' <summary>
    ''' Change the language group
    ''' </summary>
    Private Sub cbGrupos_SelectionChangeCommitted(sender As Object, e As System.EventArgs) Handles cbGrupos.SelectionChangeCommitted
        Dim selVal As Integer = Conversions.ToInteger(Me.cbGrupos.SelectedValue)
        Me.cargarDiccionariosEnElGrupo(selVal, Me.dgvDiccionarios)

        Me.idGrupoActual = Conversions.ToInteger(Me.cbGrupos.SelectedValue)
        Me.darFormatoDgvGrupos()
        Me.hayCambios = True
    End Sub
End Class

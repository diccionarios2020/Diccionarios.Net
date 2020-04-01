Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Collections
Imports System.Data
Imports System.Data.Common
Imports System.Data.OleDb
Imports System.Runtime.CompilerServices
Imports System.Windows.Forms

<OptionText()>
 Public Class cargadorDeDiccionarios
    Private vConexion As String

    Public Property conexion As String
        Get
            Return Me.vConexion
        End Get
        Set(ByVal value As String)
            Me.vConexion = value
        End Set
    End Property

    Public Sub New()
        MyBase.New()
    End Sub

    Friend Sub cargarDiccionarios(ByVal unaID_Grupo As Integer, ByRef cbDic As ComboBox, ByRef diccionarios As Collection)
        Dim enumerator As IEnumerator = Nothing
        Dim oleDbConnection As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(Me.conexion)
        Dim dataSet As System.Data.DataSet = New System.Data.DataSet()
        Dim oleDbDataAdapter As System.Data.OleDb.OleDbDataAdapter = New System.Data.OleDb.OleDbDataAdapter()
        Try
            Try
                oleDbConnection.Open()
                Dim str As String = String.Concat("SELECT Libros.* FROM LibrosEnGrupo AS LeC INNER JOIN Libros ON LeC.ID_Libro = Libros.ID_Libro WHERE LeC.ID_Grupo = ", Conversions.ToString(unaID_Grupo))
                oleDbDataAdapter.SelectCommand = New OleDbCommand(str, oleDbConnection)
                oleDbDataAdapter.Fill(dataSet)
                cbDic.DataSource = dataSet.Tables(0)
                cbDic.DisplayMember = "Titulo"
                cbDic.ValueMember = "ID_Libro"
                Try
                    enumerator = dataSet.Tables(0).Rows.GetEnumerator()
                    While enumerator.MoveNext()
                        Dim objectValue As Object = RuntimeHelpers.GetObjectValue(enumerator.Current)
                        Dim diccionario As Diccionario = New Diccionario()
                        Dim [integer] As Diccionario = diccionario
                        Dim objArray() As Object = {"id_libro"}
                        [integer].id_libro = Conversions.ToInteger(NewLateBinding.LateGet(objectValue, Nothing, "item", objArray, Nothing, Nothing, Nothing))
                        objArray = New Object() {"titulo"}
                        [integer].titulo = Conversions.ToString(NewLateBinding.LateGet(objectValue, Nothing, "item", objArray, Nothing, Nothing, Nothing))
                        objArray = New Object() {"idioma"}
                        [integer].idioma = Conversions.ToString(NewLateBinding.LateGet(objectValue, Nothing, "item", objArray, Nothing, Nothing, Nothing))
                        objArray = New Object() {"ruta"}
                        [integer].ruta = Conversions.ToString(NewLateBinding.LateGet(objectValue, Nothing, "item", objArray, Nothing, Nothing, Nothing))
                        objArray = New Object() {"referencia"}
                        [integer].referencia = Conversions.ToString(NewLateBinding.LateGet(objectValue, Nothing, "item", objArray, Nothing, Nothing, Nothing))
                        [integer] = Nothing
                        diccionarios.Add(diccionario, Nothing, Nothing, Nothing)
                    End While
                Finally
                    If (TypeOf enumerator Is IDisposable) Then
                        TryCast(enumerator, IDisposable).Dispose()
                    End If
                End Try
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

    Friend Sub cargarHojasDeUnDiccionario(ByVal unaIdDiccionario As Integer, ByRef dgv As DataGridView)
        Dim oleDbConnection As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(Me.conexion)
        Dim dataSet As System.Data.DataSet = New System.Data.DataSet()
        Dim oleDbDataAdapter As System.Data.OleDb.OleDbDataAdapter = New System.Data.OleDb.OleDbDataAdapter()
        Try
            Try
                oleDbConnection.Open()
                Dim str As String = String.Concat("SELECT Encabezado, numeroPagina, archivo, primeraPalabra, ultimaPalabra FROM Paginas WHERE Paginas.ID_Libro = ", Conversions.ToString(unaIdDiccionario), " ORDER BY archivo")
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
End Class

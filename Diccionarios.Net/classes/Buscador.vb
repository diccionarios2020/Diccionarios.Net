Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Collections
Imports System.Windows.Forms

<OptionText()>
 Public Class Buscador
    Public Sub New()
        MyBase.New()
    End Sub

    Friend Sub buscarPalabra(ByVal unaPalabra As String, ByRef unDGV As DataGridView)
        Dim enumerator As IEnumerator = Nothing
        unaPalabra.ToLower()
        Try
            enumerator = DirectCast(unDGV.Rows, IEnumerable).GetEnumerator()
            While enumerator.MoveNext()
                Dim current As DataGridViewRow = DirectCast(enumerator.Current, DataGridViewRow)
                current.Cells("primeraPalabra").Value.ToString().ToLower()
                If (Operators.CompareString(current.Cells("ultimaPalabra").Value.ToString().ToLower(), unaPalabra, True) < 0) Then
                    Continue While
                End If
                current.Selected = True
                unDGV.CurrentCell = unDGV.Rows(current.Index).Cells(0)
                Return
            End While
        Finally
            If (TypeOf enumerator Is IDisposable) Then
                TryCast(enumerator, IDisposable).Dispose()
            End If
        End Try
    End Sub
End Class
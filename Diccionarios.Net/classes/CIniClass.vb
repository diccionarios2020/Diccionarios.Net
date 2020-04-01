Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Runtime.InteropServices

''' <summary>
''' This appears to be dead code
''' </summary>
<OptionText()>
Public Class CIniClass
    Private m_Ini As String

    Public Property Archivo As String
        Get
            Return Me.m_Ini
        End Get
        Set(ByVal value As String)
            Me.m_Ini = value
        End Set
    End Property

    Public Sub New()
        MyBase.New()
    End Sub

    <DllImport("kernel32", CharSet:=CharSet.Ansi, EntryPoint:="GetPrivateProfileStringA", ExactSpelling:=True, SetLastError:=True)>
    Private Shared Function GetPrivateProfileStringKey(ByRef lpApplicationName As String, ByRef lpKeyName As String, ByRef lpDefault As String, ByRef lpReturnedString As String, ByVal nSize As Integer, ByRef lpFileName As String) As Integer
    End Function

    Public Sub GrabaIni(ByVal Seccion As String, ByVal Llave As String, ByVal Valor As String)
        Dim num As Long = CLng(CIniClass.WritePrivateProfileString(Seccion, Llave, Valor, Me.m_Ini))
    End Sub

    Public Function LeeIni(ByVal Seccion As String, ByVal Llave As String) As String
        Dim str As String = New String(Strings.ChrW(32), 255)
        Dim str1 As String = ""
        Dim privateProfileStringKey As Long = CLng(CIniClass.GetPrivateProfileStringKey(Seccion, Llave, str1, str, Strings.Len(str), Me.m_Ini))
        If (Strings.InStr(str, ChrW(0), CompareMethod.Text) <> 0) Then
            str = Strings.Left(str, Strings.Len(str) - 1)
        End If
        Return str
    End Function

    <DllImport("kernel32", CharSet:=CharSet.Ansi, EntryPoint:="WritePrivateProfileStringA", ExactSpelling:=True, SetLastError:=True)>
    Private Shared Function WritePrivateProfileString(ByRef lpApplicationName As String, ByRef lpKeyName As String, ByRef lpString As String, ByRef lpFileName As String) As Integer
    End Function
End Class
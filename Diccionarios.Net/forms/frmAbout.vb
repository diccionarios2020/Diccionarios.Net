Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports System.Windows.Forms

<OptionText()>
Public NotInheritable Class frmAbout
    Inherits Form

    Private loaded As Boolean

    Public Sub New()
        Me.loaded = False
        Me.InitializeComponent()
    End Sub

    Private Sub frmAbout_Click(sender As Object, e As System.EventArgs) Handles Me.Click
        Me.Dispose()
    End Sub

    Private Sub frmAbout_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.Opacity = 0
        Me.Timer1.Interval = 50
        Me.Timer1.Start()
        Me.Height = 300
        Me.Width = 498
        '498x300
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As System.EventArgs) Handles Timer1.Tick
        If (Not Me.loaded) Then
            Me.Opacity = Me.Opacity + 0.1
            If (Me.Opacity >= 1) Then
                Me.loaded = True
                Me.Timer1.Interval = 10000
            End If
        Else
            Me.Opacity = Me.Opacity - 0.05
            Me.Timer1.Interval = 50
            If (Me.Opacity <= 0) Then
                Me.Close()
            End If
        End If
    End Sub
End Class

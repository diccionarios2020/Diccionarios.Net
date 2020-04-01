<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGrupos
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGrupos))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cbGrupos = New System.Windows.Forms.ComboBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.dgvDiccionarios = New System.Windows.Forms.DataGridView()
        Me.btnAceptar = New System.Windows.Forms.Button()
        Me.btnCancelar = New System.Windows.Forms.Button()
        Me.cbPredeterminarDiccionario = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvDiccionarios, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cbGrupos)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(494, 100)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Grupo"
        '
        'cbGrupos
        '
        Me.cbGrupos.FormattingEnabled = True
        Me.cbGrupos.Location = New System.Drawing.Point(16, 35)
        Me.cbGrupos.Name = "cbGrupos"
        Me.cbGrupos.Size = New System.Drawing.Size(456, 24)
        Me.cbGrupos.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.dgvDiccionarios)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 118)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(494, 267)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Diccionarios en este grupo"
        '
        'dgvDiccionarios
        '
        Me.dgvDiccionarios.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDiccionarios.ColumnHeadersVisible = False
        Me.dgvDiccionarios.Location = New System.Drawing.Point(16, 21)
        Me.dgvDiccionarios.MultiSelect = False
        Me.dgvDiccionarios.Name = "dgvDiccionarios"
        Me.dgvDiccionarios.ReadOnly = True
        Me.dgvDiccionarios.RowHeadersVisible = False
        Me.dgvDiccionarios.RowTemplate.Height = 24
        Me.dgvDiccionarios.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvDiccionarios.Size = New System.Drawing.Size(456, 220)
        Me.dgvDiccionarios.TabIndex = 0
        '
        'btnAceptar
        '
        Me.btnAceptar.Location = New System.Drawing.Point(303, 401)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.Size = New System.Drawing.Size(93, 28)
        Me.btnAceptar.TabIndex = 2
        Me.btnAceptar.Text = "Acceptar"
        Me.btnAceptar.UseVisualStyleBackColor = True
        '
        'btnCancelar
        '
        Me.btnCancelar.Location = New System.Drawing.Point(409, 401)
        Me.btnCancelar.Name = "btnCancelar"
        Me.btnCancelar.Size = New System.Drawing.Size(93, 28)
        Me.btnCancelar.TabIndex = 3
        Me.btnCancelar.Text = "Cancelar"
        Me.btnCancelar.UseVisualStyleBackColor = True
        '
        'cbPredeterminarDiccionario
        '
        Me.cbPredeterminarDiccionario.AutoSize = True
        Me.cbPredeterminarDiccionario.Enabled = False
        Me.cbPredeterminarDiccionario.Location = New System.Drawing.Point(29, 401)
        Me.cbPredeterminarDiccionario.Name = "cbPredeterminarDiccionario"
        Me.cbPredeterminarDiccionario.Size = New System.Drawing.Size(192, 21)
        Me.cbPredeterminarDiccionario.TabIndex = 4
        Me.cbPredeterminarDiccionario.Text = "Predeterminar diccionario"
        Me.cbPredeterminarDiccionario.UseVisualStyleBackColor = True
        '
        'frmGrupos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(518, 441)
        Me.Controls.Add(Me.cbPredeterminarDiccionario)
        Me.Controls.Add(Me.btnCancelar)
        Me.Controls.Add(Me.btnAceptar)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmGrupos"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Abrir grupo de diccionarios"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.dgvDiccionarios, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnAceptar As System.Windows.Forms.Button
    Friend WithEvents btnCancelar As System.Windows.Forms.Button
    Friend WithEvents cbGrupos As System.Windows.Forms.ComboBox
    Friend WithEvents dgvDiccionarios As System.Windows.Forms.DataGridView
    Friend WithEvents cbPredeterminarDiccionario As System.Windows.Forms.CheckBox
End Class

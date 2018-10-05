<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConcentradoInfonavit
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.cboserie = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboperiodo = New System.Windows.Forms.ComboBox()
        Me.chkinter = New System.Windows.Forms.CheckBox()
        Me.chkPeriodo = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'cboserie
        '
        Me.cboserie.FormattingEnabled = True
        Me.cboserie.Items.AddRange(New Object() {"A", "B", "C", "D", "E", "F", "G", "H", "I"})
        Me.cboserie.Location = New System.Drawing.Point(343, 42)
        Me.cboserie.Name = "cboserie"
        Me.cboserie.Size = New System.Drawing.Size(59, 26)
        Me.cboserie.TabIndex = 25
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(294, 46)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(44, 18)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "Serie:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(29, 46)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 18)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "Periodo:"
        '
        'cboperiodo
        '
        Me.cboperiodo.FormattingEnabled = True
        Me.cboperiodo.Location = New System.Drawing.Point(97, 42)
        Me.cboperiodo.Name = "cboperiodo"
        Me.cboperiodo.Size = New System.Drawing.Size(191, 26)
        Me.cboperiodo.TabIndex = 22
        '
        'chkinter
        '
        Me.chkinter.AutoSize = True
        Me.chkinter.BackColor = System.Drawing.Color.Transparent
        Me.chkinter.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkinter.Location = New System.Drawing.Point(297, 74)
        Me.chkinter.Name = "chkinter"
        Me.chkinter.Size = New System.Drawing.Size(110, 22)
        Me.chkinter.TabIndex = 27
        Me.chkinter.Text = "Interbancario"
        Me.chkinter.UseVisualStyleBackColor = False
        '
        'chkPeriodo
        '
        Me.chkPeriodo.AutoSize = True
        Me.chkPeriodo.BackColor = System.Drawing.Color.Transparent
        Me.chkPeriodo.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkPeriodo.Location = New System.Drawing.Point(32, 74)
        Me.chkPeriodo.Name = "chkPeriodo"
        Me.chkPeriodo.Size = New System.Drawing.Size(235, 22)
        Me.chkPeriodo.TabIndex = 26
        Me.chkPeriodo.Text = "Tomar en cuenta todo el Bimestre"
        Me.chkPeriodo.UseVisualStyleBackColor = False
        '
        'frmConcentradoInfonavit
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(645, 397)
        Me.Controls.Add(Me.chkinter)
        Me.Controls.Add(Me.chkPeriodo)
        Me.Controls.Add(Me.cboserie)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboperiodo)
        Me.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "frmConcentradoInfonavit"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmConcentradoInfonavit"
        Me.UseWaitCursor = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboserie As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboperiodo As System.Windows.Forms.ComboBox
    Friend WithEvents chkinter As System.Windows.Forms.CheckBox
    Friend WithEvents chkPeriodo As System.Windows.Forms.CheckBox
End Class

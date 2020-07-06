<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNominaFinal
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNominaFinal))
        Me.lblRuta = New System.Windows.Forms.Label()
        Me.cmdCerrar = New System.Windows.Forms.Button()
        Me.pnlProgreso = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.pgbProgreso = New System.Windows.Forms.ProgressBar()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.tsbNuevo = New System.Windows.Forms.ToolStripButton()
        Me.tsbImportar = New System.Windows.Forms.ToolStripButton()
        Me.tsbGuardar = New System.Windows.Forms.ToolStripButton()
        Me.tsbEnviar = New System.Windows.Forms.ToolStripButton()
        Me.tsbCancelar = New System.Windows.Forms.ToolStripButton()
        Me.tsbProcesar = New System.Windows.Forms.ToolStripButton()
        Me.pnlCatalogo = New System.Windows.Forms.Panel()
        Me.chkAll = New System.Windows.Forms.CheckBox()
        Me.lsvLista = New System.Windows.Forms.ListView()
        Me.cmdVerificar = New System.Windows.Forms.Button()
        Me.cboTipoNomina = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboserie = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboperiodo = New System.Windows.Forms.ComboBox()
        Me.pnlProgreso.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.pnlCatalogo.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblRuta
        '
        Me.lblRuta.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblRuta.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRuta.Location = New System.Drawing.Point(2, 492)
        Me.lblRuta.Name = "lblRuta"
        Me.lblRuta.Size = New System.Drawing.Size(604, 39)
        Me.lblRuta.TabIndex = 33
        '
        'cmdCerrar
        '
        Me.cmdCerrar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCerrar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCerrar.Location = New System.Drawing.Point(1039, 488)
        Me.cmdCerrar.Name = "cmdCerrar"
        Me.cmdCerrar.Padding = New System.Windows.Forms.Padding(0, 0, 10, 0)
        Me.cmdCerrar.Size = New System.Drawing.Size(104, 43)
        Me.cmdCerrar.TabIndex = 32
        Me.cmdCerrar.Text = "Cerrar"
        Me.cmdCerrar.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdCerrar.UseVisualStyleBackColor = True
        '
        'pnlProgreso
        '
        Me.pnlProgreso.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.pnlProgreso.Controls.Add(Me.Label2)
        Me.pnlProgreso.Controls.Add(Me.pgbProgreso)
        Me.pnlProgreso.Location = New System.Drawing.Point(353, 225)
        Me.pnlProgreso.Name = "pnlProgreso"
        Me.pnlProgreso.Size = New System.Drawing.Size(449, 84)
        Me.pnlProgreso.TabIndex = 31
        Me.pnlProgreso.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(154, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(145, 19)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Procesando registros"
        '
        'pgbProgreso
        '
        Me.pgbProgreso.Location = New System.Drawing.Point(17, 12)
        Me.pgbProgreso.Name = "pgbProgreso"
        Me.pgbProgreso.Size = New System.Drawing.Size(413, 30)
        Me.pgbProgreso.TabIndex = 0
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbNuevo, Me.tsbImportar, Me.tsbGuardar, Me.tsbEnviar, Me.tsbCancelar, Me.tsbProcesar})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1155, 54)
        Me.ToolStrip1.TabIndex = 30
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'tsbNuevo
        '
        Me.tsbNuevo.Image = CType(resources.GetObject("tsbNuevo.Image"), System.Drawing.Image)
        Me.tsbNuevo.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbNuevo.Name = "tsbNuevo"
        Me.tsbNuevo.Size = New System.Drawing.Size(83, 51)
        Me.tsbNuevo.Text = "Agregar Excel"
        Me.tsbNuevo.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'tsbImportar
        '
        Me.tsbImportar.Enabled = False
        Me.tsbImportar.Image = CType(resources.GetObject("tsbImportar.Image"), System.Drawing.Image)
        Me.tsbImportar.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbImportar.Name = "tsbImportar"
        Me.tsbImportar.Size = New System.Drawing.Size(99, 51)
        Me.tsbImportar.Text = "Importar archivo"
        Me.tsbImportar.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'tsbGuardar
        '
        Me.tsbGuardar.AutoSize = False
        Me.tsbGuardar.Enabled = False
        Me.tsbGuardar.Image = Global.OperadoraNominas.My.Resources.Resources.if_magnifier_data_532758
        Me.tsbGuardar.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbGuardar.Name = "tsbGuardar"
        Me.tsbGuardar.Size = New System.Drawing.Size(90, 51)
        Me.tsbGuardar.Text = "Buscar"
        Me.tsbGuardar.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'tsbEnviar
        '
        Me.tsbEnviar.AutoSize = False
        Me.tsbEnviar.Enabled = False
        Me.tsbEnviar.Image = CType(resources.GetObject("tsbEnviar.Image"), System.Drawing.Image)
        Me.tsbEnviar.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbEnviar.Name = "tsbEnviar"
        Me.tsbEnviar.Size = New System.Drawing.Size(90, 51)
        Me.tsbEnviar.Text = "Guardar"
        Me.tsbEnviar.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'tsbCancelar
        '
        Me.tsbCancelar.AutoSize = False
        Me.tsbCancelar.Enabled = False
        Me.tsbCancelar.Image = CType(resources.GetObject("tsbCancelar.Image"), System.Drawing.Image)
        Me.tsbCancelar.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbCancelar.Name = "tsbCancelar"
        Me.tsbCancelar.Size = New System.Drawing.Size(90, 51)
        Me.tsbCancelar.Text = "Cancelar"
        Me.tsbCancelar.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'tsbProcesar
        '
        Me.tsbProcesar.Enabled = False
        Me.tsbProcesar.Image = CType(resources.GetObject("tsbProcesar.Image"), System.Drawing.Image)
        Me.tsbProcesar.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbProcesar.Name = "tsbProcesar"
        Me.tsbProcesar.Size = New System.Drawing.Size(98, 51)
        Me.tsbProcesar.Text = "Procesar archivo"
        Me.tsbProcesar.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.tsbProcesar.Visible = False
        '
        'pnlCatalogo
        '
        Me.pnlCatalogo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlCatalogo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlCatalogo.Controls.Add(Me.chkAll)
        Me.pnlCatalogo.Controls.Add(Me.lsvLista)
        Me.pnlCatalogo.Enabled = False
        Me.pnlCatalogo.Location = New System.Drawing.Point(0, 56)
        Me.pnlCatalogo.Name = "pnlCatalogo"
        Me.pnlCatalogo.Size = New System.Drawing.Size(1155, 426)
        Me.pnlCatalogo.TabIndex = 29
        '
        'chkAll
        '
        Me.chkAll.AutoSize = True
        Me.chkAll.BackColor = System.Drawing.Color.Transparent
        Me.chkAll.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAll.Location = New System.Drawing.Point(3, 3)
        Me.chkAll.Name = "chkAll"
        Me.chkAll.Size = New System.Drawing.Size(107, 22)
        Me.chkAll.TabIndex = 4
        Me.chkAll.Text = "Marcar todos"
        Me.chkAll.UseVisualStyleBackColor = False
        '
        'lsvLista
        '
        Me.lsvLista.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lsvLista.CheckBoxes = True
        Me.lsvLista.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lsvLista.FullRowSelect = True
        Me.lsvLista.GridLines = True
        Me.lsvLista.HideSelection = False
        Me.lsvLista.Location = New System.Drawing.Point(1, 31)
        Me.lsvLista.MultiSelect = False
        Me.lsvLista.Name = "lsvLista"
        Me.lsvLista.Size = New System.Drawing.Size(1147, 388)
        Me.lsvLista.TabIndex = 2
        Me.lsvLista.UseCompatibleStateImageBehavior = False
        Me.lsvLista.View = System.Windows.Forms.View.Details
        '
        'cmdVerificar
        '
        Me.cmdVerificar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdVerificar.Location = New System.Drawing.Point(922, 492)
        Me.cmdVerificar.Name = "cmdVerificar"
        Me.cmdVerificar.Size = New System.Drawing.Size(111, 39)
        Me.cmdVerificar.TabIndex = 34
        Me.cmdVerificar.Text = "Verificar"
        Me.cmdVerificar.UseVisualStyleBackColor = True
        '
        'cboTipoNomina
        '
        Me.cboTipoNomina.FormattingEnabled = True
        Me.cboTipoNomina.Items.AddRange(New Object() {"Abordo", "Descanso"})
        Me.cboTipoNomina.Location = New System.Drawing.Point(884, 12)
        Me.cboTipoNomina.Name = "cboTipoNomina"
        Me.cboTipoNomina.Size = New System.Drawing.Size(134, 27)
        Me.cboTipoNomina.TabIndex = 40
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(842, 17)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(41, 19)
        Me.Label4.TabIndex = 39
        Me.Label4.Text = "Tipo:"
        '
        'cboserie
        '
        Me.cboserie.FormattingEnabled = True
        Me.cboserie.Items.AddRange(New Object() {"A", "B", "C", "D", "E", "F", "G", "H", "I"})
        Me.cboserie.Location = New System.Drawing.Point(770, 12)
        Me.cboserie.Name = "cboserie"
        Me.cboserie.Size = New System.Drawing.Size(59, 27)
        Me.cboserie.TabIndex = 38
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(721, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 19)
        Me.Label3.TabIndex = 37
        Me.Label3.Text = "Serie:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(456, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 19)
        Me.Label1.TabIndex = 36
        Me.Label1.Text = "Periodo:"
        '
        'cboperiodo
        '
        Me.cboperiodo.FormattingEnabled = True
        Me.cboperiodo.Location = New System.Drawing.Point(524, 12)
        Me.cboperiodo.Name = "cboperiodo"
        Me.cboperiodo.Size = New System.Drawing.Size(191, 27)
        Me.cboperiodo.TabIndex = 35
        '
        'frmNominaFinal
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(1155, 533)
        Me.Controls.Add(Me.cboTipoNomina)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cboserie)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboperiodo)
        Me.Controls.Add(Me.cmdVerificar)
        Me.Controls.Add(Me.lblRuta)
        Me.Controls.Add(Me.cmdCerrar)
        Me.Controls.Add(Me.pnlProgreso)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.pnlCatalogo)
        Me.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "frmNominaFinal"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "NOMINA FINAL LAYOUTS"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlProgreso.ResumeLayout(False)
        Me.pnlProgreso.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.pnlCatalogo.ResumeLayout(False)
        Me.pnlCatalogo.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblRuta As Label
    Friend WithEvents cmdCerrar As Button
    Friend WithEvents pnlProgreso As Panel
    Friend WithEvents Label2 As Label
    Friend WithEvents pgbProgreso As ProgressBar
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents tsbNuevo As ToolStripButton
    Friend WithEvents tsbImportar As ToolStripButton
    Friend WithEvents tsbGuardar As ToolStripButton
    Friend WithEvents tsbProcesar As ToolStripButton
    Friend WithEvents tsbCancelar As ToolStripButton
    Friend WithEvents pnlCatalogo As Panel
    Friend WithEvents chkAll As CheckBox
    Friend WithEvents lsvLista As ListView
    Friend WithEvents cmdVerificar As System.Windows.Forms.Button
    Friend WithEvents tsbEnviar As System.Windows.Forms.ToolStripButton
    Friend WithEvents cboTipoNomina As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboserie As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboperiodo As System.Windows.Forms.ComboBox
End Class

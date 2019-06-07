<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSurvey
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSurvey))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cmb_Validacion = New System.Windows.Forms.ComboBox
        Me.Txt_PathExcel_ = New System.Windows.Forms.TextBox
        Me.btn_Importar_db = New System.Windows.Forms.Button
        Me.btn_ValidarNulos = New System.Windows.Forms.Button
        Me.btn_ImportarExcelaGrid = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.btn_AbrirExcel_ = New System.Windows.Forms.Button
        Me.tabSurvey = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dg_Excel = New System.Windows.Forms.DataGridView
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dg_Validacion = New System.Windows.Forms.DataGridView
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1.SuspendLayout()
        Me.tabSurvey.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dg_Excel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dg_Validacion, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cmb_Validacion)
        Me.GroupBox1.Controls.Add(Me.Txt_PathExcel_)
        Me.GroupBox1.Controls.Add(Me.btn_Importar_db)
        Me.GroupBox1.Controls.Add(Me.btn_ValidarNulos)
        Me.GroupBox1.Controls.Add(Me.btn_ImportarExcelaGrid)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.btn_AbrirExcel_)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupBox1.Location = New System.Drawing.Point(0, 357)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(665, 85)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'cmb_Validacion
        '
        Me.cmb_Validacion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb_Validacion.Enabled = False
        Me.cmb_Validacion.FormattingEnabled = True
        Me.cmb_Validacion.Location = New System.Drawing.Point(97, 42)
        Me.cmb_Validacion.Name = "cmb_Validacion"
        Me.cmb_Validacion.Size = New System.Drawing.Size(143, 21)
        Me.cmb_Validacion.TabIndex = 13
        '
        'Txt_PathExcel_
        '
        Me.Txt_PathExcel_.Location = New System.Drawing.Point(97, 17)
        Me.Txt_PathExcel_.Name = "Txt_PathExcel_"
        Me.Txt_PathExcel_.Size = New System.Drawing.Size(334, 20)
        Me.Txt_PathExcel_.TabIndex = 2
        '
        'btn_Importar_db
        '
        Me.btn_Importar_db.Enabled = False
        Me.btn_Importar_db.Image = CType(resources.GetObject("btn_Importar_db.Image"), System.Drawing.Image)
        Me.btn_Importar_db.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btn_Importar_db.Location = New System.Drawing.Point(587, 14)
        Me.btn_Importar_db.Name = "btn_Importar_db"
        Me.btn_Importar_db.Size = New System.Drawing.Size(75, 63)
        Me.btn_Importar_db.TabIndex = 11
        Me.btn_Importar_db.Text = "Import Excel File to DB"
        Me.btn_Importar_db.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btn_Importar_db.UseVisualStyleBackColor = True
        '
        'btn_ValidarNulos
        '
        Me.btn_ValidarNulos.Enabled = False
        Me.btn_ValidarNulos.Image = CType(resources.GetObject("btn_ValidarNulos.Image"), System.Drawing.Image)
        Me.btn_ValidarNulos.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btn_ValidarNulos.Location = New System.Drawing.Point(511, 14)
        Me.btn_ValidarNulos.Name = "btn_ValidarNulos"
        Me.btn_ValidarNulos.Size = New System.Drawing.Size(75, 63)
        Me.btn_ValidarNulos.TabIndex = 10
        Me.btn_ValidarNulos.Text = "Data Validate"
        Me.btn_ValidarNulos.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btn_ValidarNulos.UseVisualStyleBackColor = True
        '
        'btn_ImportarExcelaGrid
        '
        Me.btn_ImportarExcelaGrid.Enabled = False
        Me.btn_ImportarExcelaGrid.Image = CType(resources.GetObject("btn_ImportarExcelaGrid.Image"), System.Drawing.Image)
        Me.btn_ImportarExcelaGrid.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btn_ImportarExcelaGrid.Location = New System.Drawing.Point(435, 14)
        Me.btn_ImportarExcelaGrid.Name = "btn_ImportarExcelaGrid"
        Me.btn_ImportarExcelaGrid.Size = New System.Drawing.Size(75, 63)
        Me.btn_ImportarExcelaGrid.TabIndex = 9
        Me.btn_ImportarExcelaGrid.Text = "View Excel File"
        Me.btn_ImportarExcelaGrid.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btn_ImportarExcelaGrid.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(56, 50)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(35, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Sheet"
        '
        'btn_AbrirExcel_
        '
        Me.btn_AbrirExcel_.Image = CType(resources.GetObject("btn_AbrirExcel_.Image"), System.Drawing.Image)
        Me.btn_AbrirExcel_.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_AbrirExcel_.Location = New System.Drawing.Point(6, 14)
        Me.btn_AbrirExcel_.Name = "btn_AbrirExcel_"
        Me.btn_AbrirExcel_.Size = New System.Drawing.Size(85, 25)
        Me.btn_AbrirExcel_.TabIndex = 7
        Me.btn_AbrirExcel_.Text = "Search File"
        Me.btn_AbrirExcel_.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btn_AbrirExcel_.UseVisualStyleBackColor = True
        '
        'tabSurvey
        '
        Me.tabSurvey.Controls.Add(Me.TabPage1)
        Me.tabSurvey.Controls.Add(Me.TabPage2)
        Me.tabSurvey.Location = New System.Drawing.Point(0, 0)
        Me.tabSurvey.Name = "tabSurvey"
        Me.tabSurvey.SelectedIndex = 0
        Me.tabSurvey.Size = New System.Drawing.Size(662, 355)
        Me.tabSurvey.TabIndex = 2
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dg_Excel)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(654, 329)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "File Open"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dg_Excel
        '
        Me.dg_Excel.AllowUserToAddRows = False
        Me.dg_Excel.AllowUserToDeleteRows = False
        Me.dg_Excel.AllowUserToOrderColumns = True
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.LightYellow
        Me.dg_Excel.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dg_Excel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dg_Excel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dg_Excel.Location = New System.Drawing.Point(3, 3)
        Me.dg_Excel.Name = "dg_Excel"
        Me.dg_Excel.ReadOnly = True
        Me.dg_Excel.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.dg_Excel.RowHeadersWidth = 51
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.White
        Me.dg_Excel.RowsDefaultCellStyle = DataGridViewCellStyle2
        Me.dg_Excel.Size = New System.Drawing.Size(648, 323)
        Me.dg_Excel.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dg_Validacion)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(654, 329)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Validation Errors"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dg_Validacion
        '
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.LightYellow
        Me.dg_Validacion.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dg_Validacion.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dg_Validacion.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1})
        Me.dg_Validacion.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dg_Validacion.Location = New System.Drawing.Point(3, 3)
        Me.dg_Validacion.Name = "dg_Validacion"
        Me.dg_Validacion.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.dg_Validacion.RowHeadersWidth = 51
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.White
        Me.dg_Validacion.RowsDefaultCellStyle = DataGridViewCellStyle4
        Me.dg_Validacion.Size = New System.Drawing.Size(648, 323)
        Me.dg_Validacion.TabIndex = 1
        '
        'Column1
        '
        Me.Column1.HeaderText = "Validation of information"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Width = 550
        '
        'frmSurvey
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(665, 442)
        Me.Controls.Add(Me.tabSurvey)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmSurvey"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Survey"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.tabSurvey.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dg_Excel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dg_Validacion, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmb_Validacion As System.Windows.Forms.ComboBox
    Friend WithEvents Txt_PathExcel_ As System.Windows.Forms.TextBox
    Friend WithEvents btn_Importar_db As System.Windows.Forms.Button
    Friend WithEvents btn_ValidarNulos As System.Windows.Forms.Button
    Friend WithEvents btn_ImportarExcelaGrid As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btn_AbrirExcel_ As System.Windows.Forms.Button
    Friend WithEvents tabSurvey As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents dg_Excel As System.Windows.Forms.DataGridView
    Friend WithEvents dg_Validacion As System.Windows.Forms.DataGridView
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class

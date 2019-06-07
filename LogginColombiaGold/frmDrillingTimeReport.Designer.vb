<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDrillingTimeReport
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDrillingTimeReport))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnExcel2 = New System.Windows.Forms.PictureBox
        Me.dtpDateFin = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpDateIni = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        CType(Me.btnExcel2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnExcel2)
        Me.GroupBox1.Controls.Add(Me.dtpDateFin)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.dtpDateIni)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(333, 72)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'btnExcel2
        '
        Me.btnExcel2.Image = CType(resources.GetObject("btnExcel2.Image"), System.Drawing.Image)
        Me.btnExcel2.InitialImage = Nothing
        Me.btnExcel2.Location = New System.Drawing.Point(266, 20)
        Me.btnExcel2.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExcel2.Name = "btnExcel2"
        Me.btnExcel2.Size = New System.Drawing.Size(36, 36)
        Me.btnExcel2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.btnExcel2.TabIndex = 58
        Me.btnExcel2.TabStop = False
        Me.btnExcel2.Tag = "History Report"
        '
        'dtpDateFin
        '
        Me.dtpDateFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDateFin.Location = New System.Drawing.Point(143, 33)
        Me.dtpDateFin.Name = "dtpDateFin"
        Me.dtpDateFin.Size = New System.Drawing.Size(96, 20)
        Me.dtpDateFin.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(144, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "End Date"
        '
        'dtpDateIni
        '
        Me.dtpDateIni.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDateIni.Location = New System.Drawing.Point(19, 33)
        Me.dtpDateIni.Name = "dtpDateIni"
        Me.dtpDateIni.Size = New System.Drawing.Size(96, 20)
        Me.dtpDateIni.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Start Date"
        '
        'frmDrillingTimeReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(361, 105)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDrillingTimeReport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Drilling Time Report"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.btnExcel2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Private WithEvents btnExcel2 As System.Windows.Forms.PictureBox
    Friend WithEvents dtpDateFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpDateIni As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class

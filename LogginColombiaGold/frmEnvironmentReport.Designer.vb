<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEnvironmentReport
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEnvironmentReport))
        Me.btnExcel2 = New System.Windows.Forms.PictureBox
        Me.txtYear = New System.Windows.Forms.NumericUpDown
        Me.Label1 = New System.Windows.Forms.Label
        CType(Me.btnExcel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtYear, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnExcel2
        '
        Me.btnExcel2.Image = CType(resources.GetObject("btnExcel2.Image"), System.Drawing.Image)
        Me.btnExcel2.InitialImage = Nothing
        Me.btnExcel2.Location = New System.Drawing.Point(231, 26)
        Me.btnExcel2.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExcel2.Name = "btnExcel2"
        Me.btnExcel2.Size = New System.Drawing.Size(36, 36)
        Me.btnExcel2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.btnExcel2.TabIndex = 59
        Me.btnExcel2.TabStop = False
        Me.btnExcel2.Tag = "History Report"
        '
        'txtYear
        '
        Me.txtYear.Location = New System.Drawing.Point(115, 32)
        Me.txtYear.Maximum = New Decimal(New Integer() {2020, 0, 0, 0})
        Me.txtYear.Minimum = New Decimal(New Integer() {2011, 0, 0, 0})
        Me.txtYear.Name = "txtYear"
        Me.txtYear.Size = New System.Drawing.Size(82, 20)
        Me.txtYear.TabIndex = 61
        Me.txtYear.ThousandsSeparator = True
        Me.txtYear.Value = New Decimal(New Integer() {2011, 0, 0, 0})
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 13)
        Me.Label1.TabIndex = 62
        Me.Label1.Text = "Year to Report:"
        '
        'frmEnvironmentReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 93)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtYear)
        Me.Controls.Add(Me.btnExcel2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEnvironmentReport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Environment Report"
        CType(Me.btnExcel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtYear, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents btnExcel2 As System.Windows.Forms.PictureBox
    Friend WithEvents txtYear As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class

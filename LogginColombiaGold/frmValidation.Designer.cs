namespace LogginColombiaGold
{
    partial class frmValidation
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmValidation));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btnValidSamples = new System.Windows.Forms.PictureBox();
            this.dgData = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.cmbTypeValidation = new System.Windows.Forms.ComboBox();
            this.bgw = new System.ComponentModel.BackgroundWorker();
            this.pbLogging = new System.Windows.Forms.ProgressBar();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnValidSamples)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgData)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.pictureBox1);
            this.groupBox1.Controls.Add(this.btnValidSamples);
            this.groupBox1.Controls.Add(this.dgData);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.cmbTypeValidation);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(456, 319);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.InitialImage = null;
            this.pictureBox1.Location = new System.Drawing.Point(5, 18);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(237, 63);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 45;
            this.pictureBox1.TabStop = false;
            // 
            // btnValidSamples
            // 
            this.btnValidSamples.Image = ((System.Drawing.Image)(resources.GetObject("btnValidSamples.Image")));
            this.btnValidSamples.InitialImage = null;
            this.btnValidSamples.Location = new System.Drawing.Point(257, 86);
            this.btnValidSamples.Margin = new System.Windows.Forms.Padding(2);
            this.btnValidSamples.Name = "btnValidSamples";
            this.btnValidSamples.Size = new System.Drawing.Size(36, 36);
            this.btnValidSamples.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.btnValidSamples.TabIndex = 44;
            this.btnValidSamples.TabStop = false;
            this.btnValidSamples.Click += new System.EventHandler(this.btnValidSamples_Click);
            // 
            // dgData
            // 
            this.dgData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgData.Location = new System.Drawing.Point(22, 134);
            this.dgData.Name = "dgData";
            this.dgData.Size = new System.Drawing.Size(420, 174);
            this.dgData.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(25, 97);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(37, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Type: ";
            // 
            // cmbTypeValidation
            // 
            this.cmbTypeValidation.FormattingEnabled = true;
            this.cmbTypeValidation.Items.AddRange(new object[] {
            "Samples",
            "Lithology",
            "Weathering",
            "Box",
            "Alterations",
            "Geotech",
            "Structures",
            "Mineralizations",
            "Density"});
            this.cmbTypeValidation.Location = new System.Drawing.Point(68, 94);
            this.cmbTypeValidation.Name = "cmbTypeValidation";
            this.cmbTypeValidation.Size = new System.Drawing.Size(174, 21);
            this.cmbTypeValidation.TabIndex = 1;
            this.cmbTypeValidation.SelectionChangeCommitted += new System.EventHandler(this.cmbTypeValidation_SelectionChangeCommitted);
            this.cmbTypeValidation.SelectedIndexChanged += new System.EventHandler(this.cmbTypeValidation_SelectedIndexChanged);
            // 
            // bgw
            // 
            this.bgw.WorkerReportsProgress = true;
            this.bgw.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgw_DoWork);
            this.bgw.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bgw_RunWorkerCompleted);
            this.bgw.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bgw_ProgressChanged);
            // 
            // pbLogging
            // 
            this.pbLogging.Location = new System.Drawing.Point(341, 337);
            this.pbLogging.Name = "pbLogging";
            this.pbLogging.Size = new System.Drawing.Size(127, 16);
            this.pbLogging.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.pbLogging.TabIndex = 46;
            // 
            // frmValidation
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(481, 363);
            this.Controls.Add(this.pbLogging);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmValidation";
            this.Text = "Validation";
            this.Load += new System.EventHandler(this.frmValidation_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnValidSamples)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgData)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DataGridView dgData;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbTypeValidation;
        private System.Windows.Forms.PictureBox btnValidSamples;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.ComponentModel.BackgroundWorker bgw;
        private System.Windows.Forms.ProgressBar pbLogging;
    }
}
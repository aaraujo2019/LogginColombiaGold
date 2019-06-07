namespace LogginColombiaGold
{
    partial class frmPpal
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPpal));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.selectDBToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.passwordChangeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.logOutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.logginToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.collarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.logginToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.reportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.reportTransactionToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuValidate = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.mnuValidation = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.logginToolStripMenuItem,
            this.reportToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(876, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.selectDBToolStripMenuItem,
            this.passwordChangeToolStripMenuItem,
            this.logOutToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // selectDBToolStripMenuItem
            // 
            this.selectDBToolStripMenuItem.Name = "selectDBToolStripMenuItem";
            this.selectDBToolStripMenuItem.Size = new System.Drawing.Size(168, 22);
            this.selectDBToolStripMenuItem.Text = "Select DB";
            this.selectDBToolStripMenuItem.Visible = false;
            // 
            // passwordChangeToolStripMenuItem
            // 
            this.passwordChangeToolStripMenuItem.Name = "passwordChangeToolStripMenuItem";
            this.passwordChangeToolStripMenuItem.Size = new System.Drawing.Size(168, 22);
            this.passwordChangeToolStripMenuItem.Text = "Password Change";
            this.passwordChangeToolStripMenuItem.Click += new System.EventHandler(this.passwordChangeToolStripMenuItem_Click);
            // 
            // logOutToolStripMenuItem
            // 
            this.logOutToolStripMenuItem.Name = "logOutToolStripMenuItem";
            this.logOutToolStripMenuItem.Size = new System.Drawing.Size(168, 22);
            this.logOutToolStripMenuItem.Text = "Log out";
            this.logOutToolStripMenuItem.Click += new System.EventHandler(this.logOutToolStripMenuItem_Click);
            // 
            // logginToolStripMenuItem
            // 
            this.logginToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.collarToolStripMenuItem,
            this.logginToolStripMenuItem1,
            this.mnuValidation});
            this.logginToolStripMenuItem.Name = "logginToolStripMenuItem";
            this.logginToolStripMenuItem.Size = new System.Drawing.Size(73, 20);
            this.logginToolStripMenuItem.Text = "Data Entry";
            // 
            // collarToolStripMenuItem
            // 
            this.collarToolStripMenuItem.Name = "collarToolStripMenuItem";
            this.collarToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.collarToolStripMenuItem.Text = "Collar";
            this.collarToolStripMenuItem.Click += new System.EventHandler(this.collarToolStripMenuItem_Click);
            // 
            // logginToolStripMenuItem1
            // 
            this.logginToolStripMenuItem1.Name = "logginToolStripMenuItem1";
            this.logginToolStripMenuItem1.Size = new System.Drawing.Size(152, 22);
            this.logginToolStripMenuItem1.Text = "Logging";
            this.logginToolStripMenuItem1.Click += new System.EventHandler(this.logginToolStripMenuItem1_Click);
            // 
            // reportToolStripMenuItem
            // 
            this.reportToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.reportTransactionToolStripMenuItem,
            this.mnuValidate});
            this.reportToolStripMenuItem.Name = "reportToolStripMenuItem";
            this.reportToolStripMenuItem.Size = new System.Drawing.Size(54, 20);
            this.reportToolStripMenuItem.Text = "Report";
            // 
            // reportTransactionToolStripMenuItem
            // 
            this.reportTransactionToolStripMenuItem.Name = "reportTransactionToolStripMenuItem";
            this.reportTransactionToolStripMenuItem.Size = new System.Drawing.Size(174, 22);
            this.reportTransactionToolStripMenuItem.Text = "Report Transaction";
            this.reportTransactionToolStripMenuItem.Click += new System.EventHandler(this.reportTransactionToolStripMenuItem_Click);
            // 
            // mnuValidate
            // 
            this.mnuValidate.Name = "mnuValidate";
            this.mnuValidate.Size = new System.Drawing.Size(174, 22);
            this.mnuValidate.Text = "Validation Logging";
            this.mnuValidate.Click += new System.EventHandler(this.validationLoggingToolStripMenuItem_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Location = new System.Drawing.Point(0, 460);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(876, 22);
            this.statusStrip1.TabIndex = 2;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // mnuValidation
            // 
            this.mnuValidation.Name = "mnuValidation";
            this.mnuValidation.Size = new System.Drawing.Size(152, 22);
            this.mnuValidation.Text = "Validation";
            // 
            // frmPpal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.ClientSize = new System.Drawing.Size(876, 482);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.IsMdiContainer = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmPpal";
            this.Text = "DataInGranColombiaGold";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.frmPpal_Load);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmPpal_FormClosed);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem selectDBToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem logOutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem logginToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem collarToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem logginToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem passwordChangeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem reportToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem reportTransactionToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mnuValidate;
        private System.Windows.Forms.ToolStripMenuItem mnuValidation;
    }
}
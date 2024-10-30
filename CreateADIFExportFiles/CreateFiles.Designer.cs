namespace CreateADIFExportFiles
{
    partial class CreateFiles
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CreateFiles));
            this.LblSpecificationFile = new System.Windows.Forms.Label();
            this.BtnChooseFile = new System.Windows.Forms.Button();
            this.OfdSpecificationFile = new System.Windows.Forms.OpenFileDialog();
            this.BtnCreateFiles = new System.Windows.Forms.Button();
            this.BtnClose = new System.Windows.Forms.Button();
            this.SsProgress = new System.Windows.Forms.StatusStrip();
            this.TsslProgress = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.TxtSpecificationFile = new System.Windows.Forms.TextBox();
            this.SsProgress.SuspendLayout();
            this.SuspendLayout();
            // 
            // LblSpecificationFile
            // 
            this.LblSpecificationFile.AutoSize = true;
            this.LblSpecificationFile.Location = new System.Drawing.Point(14, 40);
            this.LblSpecificationFile.Name = "LblSpecificationFile";
            this.LblSpecificationFile.Size = new System.Drawing.Size(98, 13);
            this.LblSpecificationFile.TabIndex = 0;
            this.LblSpecificationFile.Text = "ADIF Specification:";
            this.toolTip.SetToolTip(this.LblSpecificationFile, "The ADIF Specification XHTML (.htm) file.  Selecting the non-annotated version is" +
        " recommended.");
            // 
            // BtnChooseFile
            // 
            this.BtnChooseFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnChooseFile.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.BtnChooseFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnChooseFile.Location = new System.Drawing.Point(498, 36);
            this.BtnChooseFile.Name = "BtnChooseFile";
            this.BtnChooseFile.Size = new System.Drawing.Size(30, 23);
            this.BtnChooseFile.TabIndex = 2;
            this.BtnChooseFile.Text = "...";
            this.toolTip.SetToolTip(this.BtnChooseFile, "Browse for the ADIF XHTML (.htm) file");
            this.BtnChooseFile.UseVisualStyleBackColor = true;
            this.BtnChooseFile.Click += new System.EventHandler(this.BtnChooseFile_Click);
            // 
            // BtnCreateFiles
            // 
            this.BtnCreateFiles.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.BtnCreateFiles.Location = new System.Drawing.Point(196, 111);
            this.BtnCreateFiles.Name = "BtnCreateFiles";
            this.BtnCreateFiles.Size = new System.Drawing.Size(75, 23);
            this.BtnCreateFiles.TabIndex = 3;
            this.BtnCreateFiles.Text = "Create Files";
            this.toolTip.SetToolTip(this.BtnCreateFiles, "Export files from the ADIF Specification file.");
            this.BtnCreateFiles.UseVisualStyleBackColor = true;
            this.BtnCreateFiles.Click += new System.EventHandler(this.BtnCreateFiles_Click);
            // 
            // BtnClose
            // 
            this.BtnClose.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.BtnClose.Location = new System.Drawing.Point(277, 111);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(75, 23);
            this.BtnClose.TabIndex = 4;
            this.BtnClose.Text = "Close";
            this.toolTip.SetToolTip(this.BtnClose, "Close the application.");
            this.BtnClose.UseVisualStyleBackColor = true;
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // SsProgress
            // 
            this.SsProgress.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.TsslProgress});
            this.SsProgress.Location = new System.Drawing.Point(0, 151);
            this.SsProgress.Name = "SsProgress";
            this.SsProgress.Size = new System.Drawing.Size(549, 22);
            this.SsProgress.TabIndex = 5;
            this.SsProgress.Text = "statusStrip1";
            // 
            // TsslProgress
            // 
            this.TsslProgress.Name = "TsslProgress";
            this.TsslProgress.Size = new System.Drawing.Size(0, 17);
            // 
            // TxtSpecificationFile
            // 
            this.TxtSpecificationFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TxtSpecificationFile.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::CreateADIFExportFiles.Properties.Settings.Default, "TxtSpecificationFile", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.TxtSpecificationFile.Location = new System.Drawing.Point(118, 37);
            this.TxtSpecificationFile.Name = "TxtSpecificationFile";
            this.TxtSpecificationFile.Size = new System.Drawing.Size(382, 20);
            this.TxtSpecificationFile.TabIndex = 1;
            this.TxtSpecificationFile.Text = global::CreateADIFExportFiles.Properties.Settings.Default.TxtSpecificationFile;
            this.toolTip.SetToolTip(this.TxtSpecificationFile, "The ADIF Specification XHTML (.htm) file.  Selecting the non-annotated version is" +
        " recommended.");
            this.TxtSpecificationFile.TextChanged += new System.EventHandler(this.TxtSpecificationFile_TextChanged);
            // 
            // CreateFiles
            // 
            this.AcceptButton = this.BtnCreateFiles;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.BtnClose;
            this.ClientSize = new System.Drawing.Size(549, 173);
            this.Controls.Add(this.SsProgress);
            this.Controls.Add(this.BtnClose);
            this.Controls.Add(this.BtnCreateFiles);
            this.Controls.Add(this.BtnChooseFile);
            this.Controls.Add(this.LblSpecificationFile);
            this.Controls.Add(this.TxtSpecificationFile);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "CreateFiles";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Create ADIF Export Files";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.CreateFiles_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.CreateFiles_FormClosed);
            this.Load += new System.EventHandler(this.CreateFiles_Load);
            this.SsProgress.ResumeLayout(false);
            this.SsProgress.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TxtSpecificationFile;
        private System.Windows.Forms.Label LblSpecificationFile;
        private System.Windows.Forms.Button BtnChooseFile;
        private System.Windows.Forms.OpenFileDialog OfdSpecificationFile;
        private System.Windows.Forms.Button BtnCreateFiles;
        private System.Windows.Forms.Button BtnClose;
        private System.Windows.Forms.StatusStrip SsProgress;
        private System.Windows.Forms.ToolStripStatusLabel TsslProgress;
        private System.Windows.Forms.ToolTip toolTip;
    }
}


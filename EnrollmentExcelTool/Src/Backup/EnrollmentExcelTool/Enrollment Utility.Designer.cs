namespace EnrollmentExcelTool
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btnRun = new System.Windows.Forms.Button();
            this.grpMaster = new System.Windows.Forms.GroupBox();
            this.txtSourceFile = new System.Windows.Forms.TextBox();
            this.btnSource = new System.Windows.Forms.Button();
            this.txtMasterFile = new System.Windows.Forms.TextBox();
            this.btnSelectMaster = new System.Windows.Forms.Button();
            this.cmbSelectOption = new System.Windows.Forms.ComboBox();
            this.lblSelect = new System.Windows.Forms.Label();
            this.grpMaster.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnRun
            // 
            this.btnRun.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRun.Location = new System.Drawing.Point(371, 216);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 23);
            this.btnRun.TabIndex = 0;
            this.btnRun.Text = "Run";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // grpMaster
            // 
            this.grpMaster.Controls.Add(this.txtSourceFile);
            this.grpMaster.Controls.Add(this.btnSource);
            this.grpMaster.Controls.Add(this.txtMasterFile);
            this.grpMaster.Controls.Add(this.btnSelectMaster);
            this.grpMaster.Location = new System.Drawing.Point(26, 71);
            this.grpMaster.Name = "grpMaster";
            this.grpMaster.Size = new System.Drawing.Size(828, 139);
            this.grpMaster.TabIndex = 3;
            this.grpMaster.TabStop = false;
            this.grpMaster.Text = "Master To Source";
            // 
            // txtSourceFile
            // 
            this.txtSourceFile.BackColor = System.Drawing.SystemColors.HighlightText;
            this.txtSourceFile.Location = new System.Drawing.Point(53, 91);
            this.txtSourceFile.Name = "txtSourceFile";
            this.txtSourceFile.ReadOnly = true;
            this.txtSourceFile.Size = new System.Drawing.Size(573, 20);
            this.txtSourceFile.TabIndex = 6;
            // 
            // btnSource
            // 
            this.btnSource.Location = new System.Drawing.Point(632, 89);
            this.btnSource.Name = "btnSource";
            this.btnSource.Size = new System.Drawing.Size(141, 23);
            this.btnSource.TabIndex = 5;
            this.btnSource.Text = "Select Source File";
            this.btnSource.UseVisualStyleBackColor = true;
            this.btnSource.Click += new System.EventHandler(this.btnSource_Click);
            // 
            // txtMasterFile
            // 
            this.txtMasterFile.BackColor = System.Drawing.SystemColors.HighlightText;
            this.txtMasterFile.Location = new System.Drawing.Point(54, 41);
            this.txtMasterFile.Name = "txtMasterFile";
            this.txtMasterFile.ReadOnly = true;
            this.txtMasterFile.Size = new System.Drawing.Size(573, 20);
            this.txtMasterFile.TabIndex = 4;
            // 
            // btnSelectMaster
            // 
            this.btnSelectMaster.Location = new System.Drawing.Point(633, 39);
            this.btnSelectMaster.Name = "btnSelectMaster";
            this.btnSelectMaster.Size = new System.Drawing.Size(141, 23);
            this.btnSelectMaster.TabIndex = 3;
            this.btnSelectMaster.Text = "Select Master File";
            this.btnSelectMaster.UseVisualStyleBackColor = true;
            this.btnSelectMaster.Click += new System.EventHandler(this.btnSelectMaster_Click);
            // 
            // cmbSelectOption
            // 
            this.cmbSelectOption.FormattingEnabled = true;
            this.cmbSelectOption.Items.AddRange(new object[] {
            "Master To Source",
            "Source To Profile"});
            this.cmbSelectOption.Location = new System.Drawing.Point(371, 41);
            this.cmbSelectOption.Name = "cmbSelectOption";
            this.cmbSelectOption.Size = new System.Drawing.Size(162, 21);
            this.cmbSelectOption.TabIndex = 4;
            this.cmbSelectOption.SelectedIndexChanged += new System.EventHandler(this.cmbSelectOption_SelectedIndexChanged);
            // 
            // lblSelect
            // 
            this.lblSelect.AutoSize = true;
            this.lblSelect.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSelect.Location = new System.Drawing.Point(216, 44);
            this.lblSelect.Name = "lblSelect";
            this.lblSelect.Size = new System.Drawing.Size(149, 13);
            this.lblSelect.TabIndex = 5;
            this.lblSelect.Text = "Select Encoding Option :";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(871, 271);
            this.Controls.Add(this.lblSelect);
            this.Controls.Add(this.cmbSelectOption);
            this.Controls.Add(this.grpMaster);
            this.Controls.Add(this.btnRun);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Enrollment Utility";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.grpMaster.ResumeLayout(false);
            this.grpMaster.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.GroupBox grpMaster;
        private System.Windows.Forms.TextBox txtMasterFile;
        private System.Windows.Forms.Button btnSelectMaster;
        private System.Windows.Forms.ComboBox cmbSelectOption;
        private System.Windows.Forms.Label lblSelect;
        private System.Windows.Forms.TextBox txtSourceFile;
        private System.Windows.Forms.Button btnSource;
    }
}


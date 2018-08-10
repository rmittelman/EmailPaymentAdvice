namespace EmailPaymentAdvice
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
            this.lblFrom = new System.Windows.Forms.Label();
            this.dtpFromDate = new System.Windows.Forms.DateTimePicker();
            this.dtpToDate = new System.Windows.Forms.DateTimePicker();
            this.lblTo = new System.Windows.Forms.Label();
            this.clbSchools = new System.Windows.Forms.CheckedListBox();
            this.lblSchools = new System.Windows.Forms.Label();
            this.btnProcess = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.txtStatus = new System.Windows.Forms.TextBox();
            this.lblVersion = new System.Windows.Forms.Label();
            this.lblSettings = new System.Windows.Forms.Label();
            this.btnPreview = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblFrom
            // 
            this.lblFrom.AutoSize = true;
            this.lblFrom.Location = new System.Drawing.Point(35, 50);
            this.lblFrom.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblFrom.Name = "lblFrom";
            this.lblFrom.Size = new System.Drawing.Size(194, 17);
            this.lblFrom.TabIndex = 0;
            this.lblFrom.Text = "Include Payments Made From";
            this.lblFrom.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dtpFromDate
            // 
            this.dtpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpFromDate.Location = new System.Drawing.Point(239, 47);
            this.dtpFromDate.Margin = new System.Windows.Forms.Padding(4);
            this.dtpFromDate.Name = "dtpFromDate";
            this.dtpFromDate.Size = new System.Drawing.Size(132, 22);
            this.dtpFromDate.TabIndex = 3;
            this.dtpFromDate.ValueChanged += new System.EventHandler(this.dtpFromDate_ValueChanged);
            // 
            // dtpToDate
            // 
            this.dtpToDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpToDate.Location = new System.Drawing.Point(239, 79);
            this.dtpToDate.Margin = new System.Windows.Forms.Padding(4);
            this.dtpToDate.Name = "dtpToDate";
            this.dtpToDate.Size = new System.Drawing.Size(132, 22);
            this.dtpToDate.TabIndex = 5;
            this.dtpToDate.ValueChanged += new System.EventHandler(this.dtpToDate_ValueChanged);
            // 
            // lblTo
            // 
            this.lblTo.AutoSize = true;
            this.lblTo.Location = new System.Drawing.Point(12, 82);
            this.lblTo.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblTo.Name = "lblTo";
            this.lblTo.Size = new System.Drawing.Size(216, 17);
            this.lblTo.TabIndex = 4;
            this.lblTo.Text = "Include Payments Made Through";
            this.lblTo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // clbSchools
            // 
            this.clbSchools.CheckOnClick = true;
            this.clbSchools.FormattingEnabled = true;
            this.clbSchools.Location = new System.Drawing.Point(239, 127);
            this.clbSchools.Margin = new System.Windows.Forms.Padding(4);
            this.clbSchools.Name = "clbSchools";
            this.clbSchools.Size = new System.Drawing.Size(132, 106);
            this.clbSchools.TabIndex = 6;
            // 
            // lblSchools
            // 
            this.lblSchools.AutoSize = true;
            this.lblSchools.Location = new System.Drawing.Point(100, 130);
            this.lblSchools.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblSchools.Name = "lblSchools";
            this.lblSchools.Size = new System.Drawing.Size(129, 17);
            this.lblSchools.TabIndex = 7;
            this.lblSchools.Text = "Schools to Process";
            this.lblSchools.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnProcess
            // 
            this.btnProcess.Location = new System.Drawing.Point(380, 203);
            this.btnProcess.Margin = new System.Windows.Forms.Padding(4);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(193, 28);
            this.btnProcess.TabIndex = 8;
            this.btnProcess.Text = "Process Selected Schools";
            this.btnProcess.UseVisualStyleBackColor = true;
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(27, 265);
            this.lblStatus.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(48, 17);
            this.lblStatus.TabIndex = 9;
            this.lblStatus.Text = "Status";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtStatus
            // 
            this.txtStatus.Font = new System.Drawing.Font("Lucida Console", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtStatus.Location = new System.Drawing.Point(84, 262);
            this.txtStatus.Margin = new System.Windows.Forms.Padding(4);
            this.txtStatus.Multiline = true;
            this.txtStatus.Name = "txtStatus";
            this.txtStatus.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtStatus.Size = new System.Drawing.Size(692, 194);
            this.txtStatus.TabIndex = 10;
            // 
            // lblVersion
            // 
            this.lblVersion.AutoSize = true;
            this.lblVersion.Location = new System.Drawing.Point(16, 11);
            this.lblVersion.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(46, 17);
            this.lblVersion.TabIndex = 11;
            this.lblVersion.Text = "label1";
            // 
            // lblSettings
            // 
            this.lblSettings.AutoSize = true;
            this.lblSettings.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblSettings.Font = new System.Drawing.Font("Wingdings 2", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.lblSettings.ForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.lblSettings.Location = new System.Drawing.Point(745, 7);
            this.lblSettings.Name = "lblSettings";
            this.lblSettings.Size = new System.Drawing.Size(41, 31);
            this.lblSettings.TabIndex = 12;
            this.lblSettings.Text = "ß";
            this.lblSettings.Click += new System.EventHandler(this.lblSettings_Click);
            // 
            // btnPreview
            // 
            this.btnPreview.Location = new System.Drawing.Point(380, 128);
            this.btnPreview.Name = "btnPreview";
            this.btnPreview.Size = new System.Drawing.Size(193, 28);
            this.btnPreview.TabIndex = 13;
            this.btnPreview.Text = "Preview Selected Schools";
            this.btnPreview.UseVisualStyleBackColor = true;
            this.btnPreview.Click += new System.EventHandler(this.btnPreview_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(793, 469);
            this.Controls.Add(this.btnPreview);
            this.Controls.Add(this.lblSettings);
            this.Controls.Add(this.lblVersion);
            this.Controls.Add(this.txtStatus);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnProcess);
            this.Controls.Add(this.lblSchools);
            this.Controls.Add(this.clbSchools);
            this.Controls.Add(this.dtpToDate);
            this.Controls.Add(this.lblTo);
            this.Controls.Add(this.dtpFromDate);
            this.Controls.Add(this.lblFrom);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Email Payment Advice";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblFrom;
        private System.Windows.Forms.DateTimePicker dtpFromDate;
        private System.Windows.Forms.DateTimePicker dtpToDate;
        private System.Windows.Forms.Label lblTo;
        private System.Windows.Forms.CheckedListBox clbSchools;
        private System.Windows.Forms.Label lblSchools;
        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.TextBox txtStatus;
        private System.Windows.Forms.Label lblVersion;
        private System.Windows.Forms.Label lblSettings;
        private System.Windows.Forms.Button btnPreview;
    }
}
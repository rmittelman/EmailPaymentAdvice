namespace EmailPaymentAdvice
{
    partial class SettingsForm
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
            if(disposing && (components != null))
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
            this.txtStatus = new System.Windows.Forms.TextBox();
            this.lblStatus = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtFromDate = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtToDate = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtToEmail = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtFromEmail = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtFromName = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtCCs = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtBCCs = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.clbSchools = new System.Windows.Forms.CheckedListBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.txtErrorEmails = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtStatus
            // 
            this.txtStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtStatus.Location = new System.Drawing.Point(79, 440);
            this.txtStatus.Margin = new System.Windows.Forms.Padding(4);
            this.txtStatus.Multiline = true;
            this.txtStatus.Name = "txtStatus";
            this.txtStatus.Size = new System.Drawing.Size(675, 60);
            this.txtStatus.TabIndex = 8;
            // 
            // lblStatus
            // 
            this.lblStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(22, 443);
            this.lblStatus.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(48, 17);
            this.lblStatus.TabIndex = 11;
            this.lblStatus.Text = "Status";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(30, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(255, 17);
            this.label1.TabIndex = 13;
            this.label1.Text = "Schools selected to process by default:";
            // 
            // txtFromDate
            // 
            this.txtFromDate.Location = new System.Drawing.Point(699, 40);
            this.txtFromDate.Name = "txtFromDate";
            this.txtFromDate.Size = new System.Drawing.Size(55, 22);
            this.txtFromDate.TabIndex = 1;
            this.txtFromDate.TextChanged += new System.EventHandler(this.txtFromDate_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(432, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(261, 34);
            this.label2.TabIndex = 15;
            this.label2.Text = "Number of days ago to start processing:\r\n(prior to current date when processing)";
            // 
            // txtToDate
            // 
            this.txtToDate.Location = new System.Drawing.Point(699, 92);
            this.txtToDate.Name = "txtToDate";
            this.txtToDate.Size = new System.Drawing.Size(55, 22);
            this.txtToDate.TabIndex = 2;
            this.txtToDate.TextChanged += new System.EventHandler(this.txtToDate_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(433, 92);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(260, 34);
            this.label3.TabIndex = 17;
            this.label3.Text = "Number of days ago to stop processing:\r\n(prior to current date when processing)";
            // 
            // txtToEmail
            // 
            this.txtToEmail.Location = new System.Drawing.Point(288, 144);
            this.txtToEmail.Name = "txtToEmail";
            this.txtToEmail.Size = new System.Drawing.Size(465, 22);
            this.txtToEmail.TabIndex = 3;
            this.txtToEmail.TextChanged += new System.EventHandler(this.txtToEmail_TextChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(65, 144);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(217, 34);
            this.label4.TabIndex = 19;
            this.label4.Text = "Email address to send notices to:\r\n( a valid email or {student} )";
            // 
            // txtFromEmail
            // 
            this.txtFromEmail.Location = new System.Drawing.Point(288, 189);
            this.txtFromEmail.Name = "txtFromEmail";
            this.txtFromEmail.Size = new System.Drawing.Size(465, 22);
            this.txtFromEmail.TabIndex = 4;
            this.txtFromEmail.TextChanged += new System.EventHandler(this.txtFromEmail_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(4, 189);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(281, 34);
            this.label5.TabIndex = 21;
            this.label5.Text = "\"From\" email address:\r\n( {school} will be replaced by school name )";
            // 
            // txtFromName
            // 
            this.txtFromName.Location = new System.Drawing.Point(288, 236);
            this.txtFromName.Name = "txtFromName";
            this.txtFromName.Size = new System.Drawing.Size(465, 22);
            this.txtFromName.TabIndex = 5;
            this.txtFromName.TextChanged += new System.EventHandler(this.txtFromName_TextChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(1, 236);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(281, 34);
            this.label6.TabIndex = 23;
            this.label6.Text = "\"From\" email name:\r\n( {school} will be replaced by school name )";
            // 
            // txtCCs
            // 
            this.txtCCs.Location = new System.Drawing.Point(288, 281);
            this.txtCCs.Name = "txtCCs";
            this.txtCCs.Size = new System.Drawing.Size(465, 22);
            this.txtCCs.TabIndex = 6;
            this.txtCCs.TextChanged += new System.EventHandler(this.txtCCs_TextChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(28, 281);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(254, 34);
            this.label7.TabIndex = 25;
            this.label7.Text = "Send CC\'s to:\r\n(seperate multiple emails with commas)";
            // 
            // txtBCCs
            // 
            this.txtBCCs.Location = new System.Drawing.Point(289, 330);
            this.txtBCCs.Name = "txtBCCs";
            this.txtBCCs.Size = new System.Drawing.Size(465, 22);
            this.txtBCCs.TabIndex = 7;
            this.txtBCCs.TextChanged += new System.EventHandler(this.txtBCCs_TextChanged);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(29, 330);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(254, 34);
            this.label8.TabIndex = 27;
            this.label8.Text = "Send BCC\'s to:\r\n(seperate multiple emails with commas)";
            // 
            // clbSchools
            // 
            this.clbSchools.CheckOnClick = true;
            this.clbSchools.FormattingEnabled = true;
            this.clbSchools.Location = new System.Drawing.Point(288, 20);
            this.clbSchools.Margin = new System.Windows.Forms.Padding(4);
            this.clbSchools.Name = "clbSchools";
            this.clbSchools.Size = new System.Drawing.Size(132, 106);
            this.clbSchools.TabIndex = 0;
            this.clbSchools.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.clbSchools_ItemCheck);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(699, 9);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(55, 25);
            this.btnSave.TabIndex = 28;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // txtErrorEmails
            // 
            this.txtErrorEmails.Location = new System.Drawing.Point(289, 378);
            this.txtErrorEmails.Name = "txtErrorEmails";
            this.txtErrorEmails.Size = new System.Drawing.Size(465, 22);
            this.txtErrorEmails.TabIndex = 29;
            this.txtErrorEmails.TextChanged += new System.EventHandler(this.txtErrorEmails_TextChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(29, 378);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(254, 34);
            this.label9.TabIndex = 30;
            this.label9.Text = "Send error emails to:\r\n(seperate multiple emails with commas)";
            // 
            // SettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(777, 519);
            this.Controls.Add(this.txtErrorEmails);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.clbSchools);
            this.Controls.Add(this.txtBCCs);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.txtCCs);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtFromName);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtFromEmail);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtToEmail);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtToDate);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtFromDate);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtStatus);
            this.Controls.Add(this.lblStatus);
            this.Name = "SettingsForm";
            this.Text = "Settings";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SettingsForm_FormClosing);
            this.Load += new System.EventHandler(this.SettingsForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtStatus;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtFromDate;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtToDate;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtToEmail;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtFromEmail;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtFromName;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtCCs;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtBCCs;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.CheckedListBox clbSchools;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TextBox txtErrorEmails;
        private System.Windows.Forms.Label label9;
    }
}
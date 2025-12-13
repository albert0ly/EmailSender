namespace MailSenderLibTester
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.lblTenant = new System.Windows.Forms.Label();
            this.txtTenant = new System.Windows.Forms.TextBox();
            this.lblClientId = new System.Windows.Forms.Label();
            this.txtClientId = new System.Windows.Forms.TextBox();
            this.lblClientSecret = new System.Windows.Forms.Label();
            this.txtClientSecret = new System.Windows.Forms.TextBox();
            this.lblMailbox = new System.Windows.Forms.Label();
            this.txtMailbox = new System.Windows.Forms.TextBox();
            this.lblTo = new System.Windows.Forms.Label();
            this.txtTo = new System.Windows.Forms.TextBox();
            this.lblCc = new System.Windows.Forms.Label();
            this.txtCc = new System.Windows.Forms.TextBox();
            this.lblBcc = new System.Windows.Forms.Label();
            this.txtBcc = new System.Windows.Forms.TextBox();
            this.lblSubject = new System.Windows.Forms.Label();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.chkIsHtml = new System.Windows.Forms.CheckBox();
            this.lblBody = new System.Windows.Forms.Label();
            this.txtBody = new System.Windows.Forms.RichTextBox();
            this.lblAttachments = new System.Windows.Forms.Label();
            this.btnAddAttachment = new System.Windows.Forms.Button();
            this.lstAttachments = new System.Windows.Forms.ListBox();
            this.lblStatus = new System.Windows.Forms.Label();
            this.checkSaveInSent = new System.Windows.Forms.CheckBox();
            this.btnSend2 = new System.Windows.Forms.Button();
            this.btnDeleteAttachments = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblTenant
            // 
            this.lblTenant.AutoSize = true;
            this.lblTenant.Location = new System.Drawing.Point(12, 15);
            this.lblTenant.Name = "lblTenant";
            this.lblTenant.Size = new System.Drawing.Size(56, 13);
            this.lblTenant.TabIndex = 0;
            this.lblTenant.Text = "Tenant Id:";
            // 
            // txtTenant
            // 
            this.txtTenant.Location = new System.Drawing.Point(130, 12);
            this.txtTenant.Name = "txtTenant";
            this.txtTenant.Size = new System.Drawing.Size(540, 20);
            this.txtTenant.TabIndex = 1;
            // 
            // lblClientId
            // 
            this.lblClientId.AutoSize = true;
            this.lblClientId.Location = new System.Drawing.Point(12, 41);
            this.lblClientId.Name = "lblClientId";
            this.lblClientId.Size = new System.Drawing.Size(48, 13);
            this.lblClientId.TabIndex = 2;
            this.lblClientId.Text = "Client Id:";
            // 
            // txtClientId
            // 
            this.txtClientId.Location = new System.Drawing.Point(130, 38);
            this.txtClientId.Name = "txtClientId";
            this.txtClientId.Size = new System.Drawing.Size(540, 20);
            this.txtClientId.TabIndex = 3;
            // 
            // lblClientSecret
            // 
            this.lblClientSecret.AutoSize = true;
            this.lblClientSecret.Location = new System.Drawing.Point(12, 67);
            this.lblClientSecret.Name = "lblClientSecret";
            this.lblClientSecret.Size = new System.Drawing.Size(70, 13);
            this.lblClientSecret.TabIndex = 4;
            this.lblClientSecret.Text = "Client Secret:";
            // 
            // txtClientSecret
            // 
            this.txtClientSecret.Location = new System.Drawing.Point(130, 64);
            this.txtClientSecret.Name = "txtClientSecret";
            this.txtClientSecret.Size = new System.Drawing.Size(540, 20);
            this.txtClientSecret.TabIndex = 5;
            this.txtClientSecret.UseSystemPasswordChar = true;
            // 
            // lblMailbox
            // 
            this.lblMailbox.AutoSize = true;
            this.lblMailbox.Location = new System.Drawing.Point(12, 93);
            this.lblMailbox.Name = "lblMailbox";
            this.lblMailbox.Size = new System.Drawing.Size(113, 13);
            this.lblMailbox.TabIndex = 6;
            this.lblMailbox.Text = "Mailbox (sender UPN):";
            // 
            // txtMailbox
            // 
            this.txtMailbox.Location = new System.Drawing.Point(130, 90);
            this.txtMailbox.Name = "txtMailbox";
            this.txtMailbox.Size = new System.Drawing.Size(540, 20);
            this.txtMailbox.TabIndex = 7;
            // 
            // lblTo
            // 
            this.lblTo.AutoSize = true;
            this.lblTo.Location = new System.Drawing.Point(12, 125);
            this.lblTo.Name = "lblTo";
            this.lblTo.Size = new System.Drawing.Size(23, 13);
            this.lblTo.TabIndex = 8;
            this.lblTo.Text = "To:";
            // 
            // txtTo
            // 
            this.txtTo.Location = new System.Drawing.Point(130, 122);
            this.txtTo.Name = "txtTo";
            this.txtTo.Size = new System.Drawing.Size(540, 20);
            this.txtTo.TabIndex = 9;
            this.txtTo.Text = "albert.lyubarsky@albertly01.onmicrosoft.com";
            // 
            // lblCc
            // 
            this.lblCc.AutoSize = true;
            this.lblCc.Location = new System.Drawing.Point(12, 151);
            this.lblCc.Name = "lblCc";
            this.lblCc.Size = new System.Drawing.Size(23, 13);
            this.lblCc.TabIndex = 10;
            this.lblCc.Text = "Cc:";
            // 
            // txtCc
            // 
            this.txtCc.Location = new System.Drawing.Point(130, 148);
            this.txtCc.Name = "txtCc";
            this.txtCc.Size = new System.Drawing.Size(540, 20);
            this.txtCc.TabIndex = 11;
            // 
            // lblBcc
            // 
            this.lblBcc.AutoSize = true;
            this.lblBcc.Location = new System.Drawing.Point(12, 177);
            this.lblBcc.Name = "lblBcc";
            this.lblBcc.Size = new System.Drawing.Size(29, 13);
            this.lblBcc.TabIndex = 12;
            this.lblBcc.Text = "Bcc:";
            // 
            // txtBcc
            // 
            this.txtBcc.Location = new System.Drawing.Point(130, 174);
            this.txtBcc.Name = "txtBcc";
            this.txtBcc.Size = new System.Drawing.Size(540, 20);
            this.txtBcc.TabIndex = 13;
            // 
            // lblSubject
            // 
            this.lblSubject.AutoSize = true;
            this.lblSubject.Location = new System.Drawing.Point(12, 203);
            this.lblSubject.Name = "lblSubject";
            this.lblSubject.Size = new System.Drawing.Size(46, 13);
            this.lblSubject.TabIndex = 14;
            this.lblSubject.Text = "Subject:";
            // 
            // txtSubject
            // 
            this.txtSubject.Location = new System.Drawing.Point(130, 200);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(479, 20);
            this.txtSubject.TabIndex = 15;
            this.txtSubject.Text = "Subject";
            // 
            // chkIsHtml
            // 
            this.chkIsHtml.AutoSize = true;
            this.chkIsHtml.Checked = true;
            this.chkIsHtml.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIsHtml.Location = new System.Drawing.Point(615, 202);
            this.chkIsHtml.Name = "chkIsHtml";
            this.chkIsHtml.Size = new System.Drawing.Size(67, 17);
            this.chkIsHtml.TabIndex = 16;
            this.chkIsHtml.Text = "Is HTML";
            this.chkIsHtml.UseVisualStyleBackColor = true;
            // 
            // lblBody
            // 
            this.lblBody.AutoSize = true;
            this.lblBody.Location = new System.Drawing.Point(12, 278);
            this.lblBody.Name = "lblBody";
            this.lblBody.Size = new System.Drawing.Size(34, 13);
            this.lblBody.TabIndex = 17;
            this.lblBody.Text = "Body:";
            // 
            // txtBody
            // 
            this.txtBody.Location = new System.Drawing.Point(15, 294);
            this.txtBody.Name = "txtBody";
            this.txtBody.Size = new System.Drawing.Size(655, 180);
            this.txtBody.TabIndex = 18;
            this.txtBody.Text = "Body";
            // 
            // lblAttachments
            // 
            this.lblAttachments.AutoSize = true;
            this.lblAttachments.Location = new System.Drawing.Point(12, 485);
            this.lblAttachments.Name = "lblAttachments";
            this.lblAttachments.Size = new System.Drawing.Size(69, 13);
            this.lblAttachments.TabIndex = 19;
            this.lblAttachments.Text = "Attachments:";
            // 
            // btnAddAttachment
            // 
            this.btnAddAttachment.Location = new System.Drawing.Point(130, 480);
            this.btnAddAttachment.Name = "btnAddAttachment";
            this.btnAddAttachment.Size = new System.Drawing.Size(120, 23);
            this.btnAddAttachment.TabIndex = 20;
            this.btnAddAttachment.Text = "Add Attachment";
            this.btnAddAttachment.UseVisualStyleBackColor = true;
            this.btnAddAttachment.Click += new System.EventHandler(this.btnAddAttachment_Click);
            // 
            // lstAttachments
            // 
            this.lstAttachments.FormattingEnabled = true;
            this.lstAttachments.Location = new System.Drawing.Point(256, 480);
            this.lstAttachments.Name = "lstAttachments";
            this.lstAttachments.Size = new System.Drawing.Size(414, 82);
            this.lstAttachments.TabIndex = 21;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(12, 587);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(10, 13);
            this.lblStatus.TabIndex = 23;
            this.lblStatus.Text = "-";
            // 
            // checkSaveInSent
            // 
            this.checkSaveInSent.AutoSize = true;
            this.checkSaveInSent.Location = new System.Drawing.Point(130, 246);
            this.checkSaveInSent.Name = "checkSaveInSent";
            this.checkSaveInSent.Size = new System.Drawing.Size(115, 17);
            this.checkSaveInSent.TabIndex = 25;
            this.checkSaveInSent.Text = "Save in Sent Items";
            this.checkSaveInSent.UseVisualStyleBackColor = true;
            // 
            // btnSend2
            // 
            this.btnSend2.Location = new System.Drawing.Point(15, 538);
            this.btnSend2.Name = "btnSend2";
            this.btnSend2.Size = new System.Drawing.Size(120, 23);
            this.btnSend2.TabIndex = 26;
            this.btnSend2.Text = "Send 2";
            this.btnSend2.UseVisualStyleBackColor = true;
            this.btnSend2.Click += new System.EventHandler(this.btnSend2_Click);
            // 
            // btnDeleteAttachments
            // 
            this.btnDeleteAttachments.Location = new System.Drawing.Point(130, 509);
            this.btnDeleteAttachments.Name = "btnDeleteAttachments";
            this.btnDeleteAttachments.Size = new System.Drawing.Size(120, 23);
            this.btnDeleteAttachments.TabIndex = 27;
            this.btnDeleteAttachments.Text = "Delete Attachments";
            this.btnDeleteAttachments.UseVisualStyleBackColor = true;
            this.btnDeleteAttachments.Click += new System.EventHandler(this.btnDeleteAttachments_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(695, 699);
            this.Controls.Add(this.btnDeleteAttachments);
            this.Controls.Add(this.btnSend2);
            this.Controls.Add(this.checkSaveInSent);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.lstAttachments);
            this.Controls.Add(this.btnAddAttachment);
            this.Controls.Add(this.lblAttachments);
            this.Controls.Add(this.txtBody);
            this.Controls.Add(this.lblBody);
            this.Controls.Add(this.chkIsHtml);
            this.Controls.Add(this.txtSubject);
            this.Controls.Add(this.lblSubject);
            this.Controls.Add(this.txtBcc);
            this.Controls.Add(this.lblBcc);
            this.Controls.Add(this.txtCc);
            this.Controls.Add(this.lblCc);
            this.Controls.Add(this.txtTo);
            this.Controls.Add(this.lblTo);
            this.Controls.Add(this.txtMailbox);
            this.Controls.Add(this.lblMailbox);
            this.Controls.Add(this.txtClientSecret);
            this.Controls.Add(this.lblClientSecret);
            this.Controls.Add(this.txtClientId);
            this.Controls.Add(this.lblClientId);
            this.Controls.Add(this.txtTenant);
            this.Controls.Add(this.lblTenant);
            this.Name = "MainForm";
            this.Text = "MailSenderLib Tester";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblTenant;
        private System.Windows.Forms.TextBox txtTenant;
        private System.Windows.Forms.Label lblClientId;
        private System.Windows.Forms.TextBox txtClientId;
        private System.Windows.Forms.Label lblClientSecret;
        private System.Windows.Forms.TextBox txtClientSecret;
        private System.Windows.Forms.Label lblMailbox;
        private System.Windows.Forms.TextBox txtMailbox;
        private System.Windows.Forms.Label lblTo;
        private System.Windows.Forms.TextBox txtTo;
        private System.Windows.Forms.Label lblCc;
        private System.Windows.Forms.TextBox txtCc;
        private System.Windows.Forms.Label lblBcc;
        private System.Windows.Forms.TextBox txtBcc;
        private System.Windows.Forms.Label lblSubject;
        private System.Windows.Forms.TextBox txtSubject;
        private System.Windows.Forms.CheckBox chkIsHtml;
        private System.Windows.Forms.Label lblBody;
        private System.Windows.Forms.RichTextBox txtBody;
        private System.Windows.Forms.Label lblAttachments;
        private System.Windows.Forms.Button btnAddAttachment;
        private System.Windows.Forms.ListBox lstAttachments;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.CheckBox checkSaveInSent;
        private System.Windows.Forms.Button btnSend2;
        private System.Windows.Forms.Button btnDeleteAttachments;
    }
}

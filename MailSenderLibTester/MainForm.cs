using MailSenderLib;
using MailSenderLib.Models;
using MailSenderLib.Options;
using MailSenderLib.Services;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MailSenderLibTester
{
    public partial class MainForm : Form
    {
        private readonly List<string> _attachmentPaths = new List<string>();

        // Controls for receive tab
        private TabControl _tabControl;
        private TabPage _tabSend;
        private TabPage _tabReceive;
        private TextBox _txtRecvMailbox;
        private Button _btnGet;
        private DataGridView _dgvMessages;
        private TextBox _txtDetailsSubject;
        private TextBox _txtDetailsBody;
        private ListBox _lstRecvAttachments;
        private PictureBox _pbPreview;
        private Label _lblRecvStatus;

        public MainForm()
        {
            InitializeComponent();
            LoadConfigIntoFields();

            // Build runtime TabControl and move existing send controls into first tab
            SetupTabs();
        }

        private void SetupTabs()
        {
            // Create TabControl and tabs
            _tabControl = new TabControl { Dock = DockStyle.Fill };
            _tabSend = new TabPage("Send Email");
            _tabReceive = new TabPage("Receive Emails");

            // Move existing top-level controls into Send tab. We assume designer created these controls with known names.
            var sendControlNames = new[] {
                "txtTenant", "lblTenant", "txtClientId","txtClientSecret", "lblClientSecret", "lblClientId", "txtMailbox",
                "txtTo","txtCc","txtBcc","txtSubject","txtBody",
                "chkIsHtml","btnSend","btnAddAttachment","lstAttachments","lblStatus","btnSend", "checkSaveInSent","lblMailbox",
                "lblTo","lblCc","lblBcc","lblSubject","lblBody"
            };

            foreach (var name in sendControlNames)
            {
                var ctrls = this.Controls.Find(name, true);
                foreach (Control c in ctrls)
                {
                    // reparent to send tab
                    _tabSend.Controls.Add(c);
                }
            }

            // Some other controls (labels etc.) are left as-is by designer; try to move all child controls from main form's container panel if exists
            // Add tab pages to control
            _tabControl.TabPages.Add(_tabSend);
            _tabControl.TabPages.Add(_tabReceive);

            // Add TabControl to form and dock
            // Place tab control at top-level of form
            this.Controls.Add(_tabControl);
            _tabControl.BringToFront();

            // Build Receive tab UI
            BuildReceiveTabUi();
        }

        private void BuildReceiveTabUi()
        {
            var padding = 8;
            int x = padding, y = padding, w = _tabReceive.ClientSize.Width - padding * 2;

            // Mailbox textbox and Get button
            _txtRecvMailbox = new TextBox { Name = "txtRecvMailbox", Width = 300, Left = x, Top = y };
            _txtRecvMailbox.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            _txtRecvMailbox.Text = txtTo != null ? txtTo.Text : string.Empty; // if designer control exists
            _btnGet = new Button { Text = "Get", Left = x + 310, Top = y - 2, Width = 80 };
            _btnGet.Click += BtnGet_Click;

            _lblRecvStatus = new Label { Text = "", Left = x + 400, Top = y + 3, AutoSize = true };

            _tabReceive.Controls.Add(_txtRecvMailbox);
            _tabReceive.Controls.Add(_btnGet);
            _tabReceive.Controls.Add(_lblRecvStatus);

            y += 30;

            // DataGridView for messages
            _dgvMessages = new DataGridView { Left = x, Top = y, Width = _tabReceive.ClientSize.Width - 2 * padding, Height = 200, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            _dgvMessages.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _dgvMessages.ReadOnly = true;
            _dgvMessages.AutoGenerateColumns = false;
            _dgvMessages.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Id", DataPropertyName = "Id", Visible = false });
            _dgvMessages.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Subject", DataPropertyName = "Subject", Width = 400 });
            _dgvMessages.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Received", DataPropertyName = "ReceivedDateTime", Width = 160 });
            _dgvMessages.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "HasAttachments", DataPropertyName = "HasAttachments", Width = 80 });
            _dgvMessages.SelectionChanged += DgvMessages_SelectionChanged;

            _tabReceive.Controls.Add(_dgvMessages);

            y += 210;

            // Details: Subject label and textbox
            var lblSub = new Label { Text = "Subject:", Left = x, Top = y + 6, AutoSize = true };
            _txtDetailsSubject = new TextBox { Left = x + 60, Top = y, Width = 700, ReadOnly = true, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            _tabReceive.Controls.Add(lblSub);
            _tabReceive.Controls.Add(_txtDetailsSubject);

            y += 30;

            // Body
            var lblBody = new Label { Text = "Body:", Left = x, Top = y + 6, AutoSize = true };
            _txtDetailsBody = new TextBox { Left = x + 60, Top = y, Width = 700, Height = 120, Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            _tabReceive.Controls.Add(lblBody);
            _tabReceive.Controls.Add(_txtDetailsBody);

            // Attachments list
            _lstRecvAttachments = new ListBox { Left = x + 770, Top = padding + 60, Width = 240, Height = 300, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            _lstRecvAttachments.SelectedIndexChanged += LstRecvAttachments_SelectedIndexChanged;
            _tabReceive.Controls.Add(_lstRecvAttachments);

            // PictureBox preview
            _pbPreview = new PictureBox { Left = x + 770, Top = padding + 370, Width = 240, Height = 180, SizeMode = PictureBoxSizeMode.Zoom, BorderStyle = BorderStyle.FixedSingle, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            _tabReceive.Controls.Add(_pbPreview);

            // Resize handling
            _tabReceive.Resize += (s, e) =>
            {
                _dgvMessages.Width = _tabReceive.ClientSize.Width - 2 * padding - 260;
                _txtDetailsSubject.Width = _dgvMessages.Width - 60;
                _txtDetailsBody.Width = _dgvMessages.Width - 60;
                _lstRecvAttachments.Left = _dgvMessages.Right + 10;
                _pbPreview.Left = _lstRecvAttachments.Left;
            };
        }

        private async void BtnGet_Click(object sender, EventArgs e)
        {
            _btnGet.Enabled = false;
            _lblRecvStatus.Text = "Fetching...";
            try
            {
                var options = new GraphMailOptionsAuth
                {
                    TenantId = txtTenant != null ? txtTenant.Text.Trim() : string.Empty,
                    ClientId = txtClientId != null ? txtClientId.Text.Trim() : string.Empty,
                    ClientSecret = txtClientSecret != null ? txtClientSecret.Text.Trim() : string.Empty,
                    MailboxAddress = txtMailbox != null ? txtMailbox.Text.Trim() : string.Empty
                };

                var receiver = new GraphMailReceiver(options);
                var list = await receiver.ReceiveEmailsAsync(_txtRecvMailbox.Text.Trim(), CancellationToken.None);

                _dgvMessages.DataSource = list;
                _lblRecvStatus.Text = string.Format("{0} messages", list.Count);
            }
            catch (Exception ex)
            {
                _lblRecvStatus.Text = ex.Message;
            }
            finally
            {
                _btnGet.Enabled = true;
            }
        }

        private void DgvMessages_SelectionChanged(object sender, EventArgs e)
        {
            if (_dgvMessages.SelectedRows.Count == 0) return;
            var row = _dgvMessages.SelectedRows[0];
            if (row.DataBoundItem is MailMessageDto msg)
            {
                _txtDetailsSubject.Text = msg.Subject ?? string.Empty;
                _txtDetailsBody.Text = msg.Body ?? string.Empty;

                _lstRecvAttachments.Items.Clear();
                _pbPreview.Image = null;
                foreach (var a in msg.Attachments)
                {
                    _lstRecvAttachments.Items.Add(a);
                }
            }
        }

        private void LstRecvAttachments_SelectedIndexChanged(object sender, EventArgs e)
        {
            _pbPreview.Image = null;
            if (_lstRecvAttachments.SelectedItem is MailAttachmentDto a)
            {
                if (!string.IsNullOrEmpty(a.ContentBase64) && !string.IsNullOrEmpty(a.ContentType) && a.ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        var bytes = Convert.FromBase64String(a.ContentBase64);
                        using (var ms = new MemoryStream(bytes))
                        {
                            _pbPreview.Image = Image.FromStream(ms);
                        }
                    }
                    catch
                    {
                        // ignore image errors
                    }
                }
            }
        }

        private void btnAddAttachment_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Multiselect = true;
                if (ofd.ShowDialog(this) == DialogResult.OK)
                {
                    foreach (var file in ofd.FileNames)
                    {
                        _attachmentPaths.Add(file);
                        lstAttachments.Items.Add(file);
                    }
                }
            }
        }

        private async void btnSend_Click(object sender, EventArgs e)
        {
            btnSend.Enabled = false;
            lblStatus.Text = "Sending...";
            try
            {
                var optionsAuth = new GraphMailOptionsAuth
                {
                    TenantId = txtTenant.Text.Trim(),
                    ClientId = txtClientId.Text.Trim(),
                    ClientSecret = txtClientSecret.Text.Trim(),
                    MailboxAddress = txtMailbox.Text.Trim()
                };

                var options = new GraphMailOptions
                {
                    MoveToSentFolder = checkSaveInSent.Checked,
                    MarkAsRead = false  
                };

                var senderLib = new GraphMailSender(optionsAuth, options);

                var to = SplitEmails(txtTo.Text);
                var cc = SplitEmails(txtCc.Text);
                var bcc = SplitEmails(txtBcc.Text);
                var subject = txtSubject.Text;
                var body = txtBody.Text;
                var isHtml = chkIsHtml.Checked;

                var streams = new List<Stream>();
                try
                {
                    //var atts = _attachmentPaths.Select(p =>
                    //{
                    //    var s = (Stream)File.OpenRead(p);
                    //    streams.Add(s);
                    //    return (FileName: Path.GetFileName(p), ContentType: GetMimeType(p), ContentStream: s);
                    //}).ToList();


                    var attachments = new List<EmailAttachment>
                    {
                        new EmailAttachment
                        {
                            FileName = "large_file.pdf",
                            FilePath = @"C:\path\to\large_file.pdf"
                        },
                        new EmailAttachment
                        {
                            FileName = "document.docx",
                            FilePath = @"C:\path\to\document.docx"
                        }
                    };

                    var attachments1 = _attachmentPaths.Select(a => new EmailAttachment
                    {
                        FileName = Path.GetFileName(a),
                        FilePath = a
                    }).ToList();

                 //   await senderLib.SendEmailAsync(to, cc, bcc, subject, body, isHtml, atts);

                    var mailService = new GraphMailService(optionsAuth.TenantId, optionsAuth.ClientId, optionsAuth.ClientSecret);

                    await mailService.SendMailWithLargeAttachmentsAsync(
                        fromEmail: optionsAuth.MailboxAddress,
                        toEmails: to,
                        subject: subject,
                        body: body,
                        attachments: attachments1,
                        ccEmails:cc,
                        isHtml: isHtml
                    );
                }
                finally
                {
                    foreach (var s in streams)
                    {
                        try { s.Dispose(); } catch { }
                    }
                }
                lblStatus.Text = "Sent";
            }
            catch (Exception ex)
            {
                lblStatus.Text = ex.Message;
            }
            finally
            {
                btnSend.Enabled = true;
            }
        }

        private void LoadConfigIntoFields()
        {
            try
            {
                // Prefer development-only secrets file if present
                var secretsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MailSenderLibTester-secrets.config");
                NameValueCollection appSettings = null;
                if (File.Exists(secretsPath))
                {
                    var map = new ExeConfigurationFileMap { ExeConfigFilename = secretsPath };
                    var cfg = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
                    appSettings = cfg.AppSettings.Settings.Count > 0 ? ToNameValueCollection(cfg.AppSettings.Settings) : null;
                }

                string tenant = null, clientId = null, clientSecret = null, mailbox = null;
                if (appSettings != null)
                {
                    tenant = appSettings["TenantId"];
                    clientId = appSettings["ClientId"];
                    clientSecret = appSettings["ClientSecret"];
                    mailbox = appSettings["MailboxAddress"];
                }

                // Fallback to App.config if not provided in secrets
                tenant = tenant ?? ConfigurationManager.AppSettings["TenantId"];
                clientId = clientId ?? ConfigurationManager.AppSettings["ClientId"];
                clientSecret = clientSecret ?? ConfigurationManager.AppSettings["ClientSecret"];
                mailbox = mailbox ?? ConfigurationManager.AppSettings["MailboxAddress"];

                txtTenant.Text = tenant ?? string.Empty;
                txtClientId.Text = clientId ?? string.Empty;
                txtClientSecret.Text = clientSecret ?? string.Empty;
                txtMailbox.Text = mailbox ?? string.Empty;
            }
            catch
            {
                // ignore config errors and leave fields empty
            }
        }

        private static NameValueCollection ToNameValueCollection(KeyValueConfigurationCollection settings)
        {
            var nvc = new NameValueCollection();
            foreach (KeyValueConfigurationElement e in settings)
            {
                nvc[e.Key] = e.Value;
            }
            return nvc;
        }

        private static List<string> SplitEmails(string text)
        {
            return string.IsNullOrWhiteSpace(text)
                ? new List<string>()
                : text.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToList();
        }

        private static string GetMimeType(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            switch (ext)
            {
                case ".txt": return "text/plain";
                case ".htm": return "text/html";
                case ".html": return "text/html";
                case ".pdf": return "application/pdf";
                case ".jpg": return "image/jpeg";
                case ".jpeg": return "image/jpeg";
                case ".png": return "image/png";
                default: return "application/octet-stream";
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}

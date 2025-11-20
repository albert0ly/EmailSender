using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using MailSenderLib;
using System.Collections.Specialized;

namespace MailSenderLibTester
{
    public partial class MainForm : Form
    {
        private readonly List<string> _attachmentPaths = new List<string>();

        public MainForm()
        {
            InitializeComponent();
            LoadConfigIntoFields();
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
                var options = new GraphMailOptions
                {
                    TenantId = txtTenant.Text.Trim(),
                    ClientId = txtClientId.Text.Trim(),
                    ClientSecret = txtClientSecret.Text.Trim(),
                    MailboxAddress = txtMailbox.Text.Trim()
                };
                var senderLib = new GraphMailSender(options);

                var to = SplitEmails(txtTo.Text);
                var cc = SplitEmails(txtCc.Text);
                var bcc = SplitEmails(txtBcc.Text);
                var subject = txtSubject.Text;
                var body = txtBody.Text;
                var isHtml = chkIsHtml.Checked;

                var streams = new List<Stream>();
                try
                {
                    var atts = _attachmentPaths.Select(p =>
                    {
                        var s = (Stream)File.OpenRead(p);
                        streams.Add(s);
                        return (FileName: Path.GetFileName(p), ContentType: GetMimeType(p), ContentStream: s);
                    }).ToList();

                    await senderLib.SendEmailAsync(to, cc, bcc, subject, body, isHtml, atts);
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
    }
}

using MailSenderLib.Models;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace MailSenderLib.Interfaces
{
    public interface IGraphMailSender
    {
        void Dispose();
        Task SendEmailAsync(List<string> toRecipients,
                            List<string>? ccRecipients,
                            List<string>? bccRecipients,
                            string subject,
                            string body,
                            bool isHtml=true,
                            List<EmailAttachment>? attachments=null,
                            string? fromEmail = null,
                            CancellationToken ct = default);
    }
}
using MailSenderLib.Models;
using MailSenderLib.Options;
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
                            string? correlationId = null,
                            GraphMailOptions? options = null,
                            CancellationToken ct = default);
    }
}
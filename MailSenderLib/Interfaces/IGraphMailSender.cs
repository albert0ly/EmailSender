using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace MailSenderLib.Interfaces
{
    public interface IGraphMailSender
    {
        Task SendEmailAsync(
            IEnumerable<string> toRecipients,
            IEnumerable<string>? ccRecipients,
            IEnumerable<string>? bccRecipients,
            string subject,
            string body,
            bool isHtml,
            IEnumerable<(string FileName, string ContentType, Stream ContentStream)>? attachments,
            CancellationToken cancellationToken = default);
    }
}

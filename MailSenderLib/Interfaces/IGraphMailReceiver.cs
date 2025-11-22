using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using MailSenderLib.Models;

namespace MailSenderLib.Interfaces
{
    public interface IGraphMailReceiver
    {
        Task<List<MailMessageDto>> ReceiveEmailsAsync(string? mailbox, CancellationToken ct = default);
    }
}

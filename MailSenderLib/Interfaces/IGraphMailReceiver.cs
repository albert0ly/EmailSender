using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using MailSenderLib.Models;

namespace MailSenderLib.Interfaces
{
    /// <summary>
    /// Receives mail through Microsoft Graph.
    /// </summary>
    public interface IGraphMailReceiver
    {
        Task<List<MailMessageDto>> ReceiveEmailsAsync(string? mailbox, CancellationToken ct = default);
    }
}

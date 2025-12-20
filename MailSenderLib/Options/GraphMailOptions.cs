using System;

namespace MailSenderLib.Options
{
    public class GraphMailOptionsAuth
    {
        public string TenantId { get; set; } = string.Empty;
        public string ClientId { get; set; } = string.Empty;
        public string ClientSecret { get; set; } = string.Empty;
        public string MailboxAddress { get; set; } = string.Empty;
    }


    public class GraphMailOptions
    {
#pragma warning disable CA1805 // Do not initialize unnecessarily
        public bool MarkAsRead { get; set; } = false;
        public bool MoveToSentFolder { get; set; } = false;
#pragma warning restore CA1805 // Do not initialize unnecessarily
        public TimeSpan? HttpClientTimeout { get; set; } = TimeSpan.Zero;
        public long LargeAttachmentThreshold { get; set; }  = 3 * 1024 * 1024; // 3MB
        public  int ChunkSize { get; set; }  = 5 * 1024 * 1024; // 5MB
        public  long MaxTotalAttachmentSize { get; set; } = 35 * 1024 * 1024; // 35MB - protect against memory issues with huge attachments
    }
}

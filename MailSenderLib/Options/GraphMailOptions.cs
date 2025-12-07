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
    }
}

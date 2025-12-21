using System;
using System.Collections.Generic;
using System.Text;

namespace MailSenderLib.Models
{
    // Custom contract resolver for @odata.type
    internal sealed class ODataContractResolver : Newtonsoft.Json.Serialization.DefaultContractResolver
    {
        protected override string ResolvePropertyName(string propertyName)
        {
            if (propertyName == "odataType")
                return "@odata.type";
            return base.ResolvePropertyName(propertyName);
        }
    }

    // Strongly-typed payload classes for better performance and type safety
    public sealed class MessagePayload
    {
        public string? Subject { get; set; }
        public BodyPayload? Body { get; set; }
        public List<RecipientPayload>? ToRecipients { get; set; }
        public List<RecipientPayload>? CcRecipients { get; set; }
        public List<RecipientPayload>? BccRecipients { get; set; }
    }

    public sealed class BodyPayload
    {
        public string ContentType { get; set; } = string.Empty;
        public string Content { get; set; } = string.Empty;
    }

    public sealed class RecipientPayload
    {
        public EmailAddressPayload? EmailAddress { get; set; }
    }

    public sealed class EmailAddressPayload
    {
        public string Address { get; set; } = string.Empty;
    }

    public sealed class EmailAttachment
    {
        public string FileName { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
#pragma warning disable CA1805 // Do not initialize unnecessarily
        public bool IsInline { get; set; } = false;          
#pragma warning restore CA1805 // Do not initialize unnecessarily
        public string? ContentId { get; set; }                                                                           
        public string? ContentType { get; set; }
    }
}

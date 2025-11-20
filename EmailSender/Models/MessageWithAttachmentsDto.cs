using System;
using System.Collections.Generic;

namespace EmailSender.Models
{
    public class AttachmentDto
    {
        public string? Id { get; set; }
        public string? Name { get; set; }
        public string? ContentType { get; set; }
        public long? Size { get; set; }
        public bool? IsInline { get; set; }
        // Base64-encoded content for file attachments. Null for non-file or if not retrieved.
        public string? ContentBase64 { get; set; }
    }

    public class MessageDto
    {
        public string? Id { get; set; }
        public string? Subject { get; set; }
        public string? Body { get; set; }
        public DateTimeOffset? ReceivedDateTime { get; set; }
        public bool? IsRead { get; set; }
        public bool? HasAttachments { get; set; }
        public List<AttachmentDto> Attachments { get; set; } = new List<AttachmentDto>();
        public string? WebLink { get; set; }

        // New fields
        public List<string> To { get; set; } = new List<string>();
        public List<string> Cc { get; set; } = new List<string>();
        public List<string> Bcc { get; set; } = new List<string>();
        // Internet message headers (name -> value). Multiple headers with same name will be concatenated with ','.
        public Dictionary<string, string?> Headers { get; set; } = new Dictionary<string, string?>();
    }
}

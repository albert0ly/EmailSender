using System;
using System.Collections.Generic;

namespace MailSenderLib.Models
{
    /// <summary>
    /// DTO representing a mail message with metadata and attachments.
    /// </summary>
    public class MailMessageDto
    {
        /// <summary>
        /// Message identifier.
        /// </summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>
        /// Message subject.
        /// </summary>
        public string? Subject { get; set; }

        /// <summary>
        /// Message body (HTML or text).
        /// </summary>
        public string? Body { get; set; }

        /// <summary>
        /// Received date/time.
        /// </summary>
        public DateTimeOffset? ReceivedDateTime { get; set; }

        /// <summary>
        /// True if message is marked as read.
        /// </summary>
        public bool? IsRead { get; set; }

        /// <summary>
        /// True if the message has attachments.
        /// </summary>
        public bool? HasAttachments { get; set; }

        /// <summary>
        /// Link to view the message in Outlook web.
        /// </summary>
        public string? WebLink { get; set; }

        /// <summary>
        /// To recipients as email addresses.
        /// </summary>
        public List<string> To { get; set; } = new List<string>();

        /// <summary>
        /// Cc recipients as email addresses.
        /// </summary>
        public List<string> Cc { get; set; } = new List<string>();

        /// <summary>
        /// Bcc recipients as email addresses.
        /// </summary>
        public List<string> Bcc { get; set; } = new List<string>();

        /// <summary>
        /// Internet message headers (name -> value).
        /// </summary>
        public Dictionary<string, string?> Headers { get; set; } = new Dictionary<string, string?>();

        /// <summary>
        /// Attachment metadata and optional content.
        /// </summary>
        public List<MailAttachmentDto> Attachments { get; set; } = new List<MailAttachmentDto>();
    }
}

namespace MailSenderLib.Models
{
    /// <summary>
    /// DTO representing a message attachment returned by Graph.
    /// </summary>
    public class MailAttachmentDto
    {
        /// <summary>
        /// Attachment identifier.
        /// </summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>
        /// File name of the attachment.
        /// </summary>
        public string? Name { get; set; }

        /// <summary>
        /// Content type (MIME) of the attachment.
        /// </summary>
        public string? ContentType { get; set; }

        /// <summary>
        /// Size in bytes.
        /// </summary>
        public long? Size { get; set; }

        /// <summary>
        /// True if the attachment is inline.
        /// </summary>
        public bool? IsInline { get; set; }

        /// <summary>
        /// Base64-encoded content for file attachments (when retrieved).
        /// </summary>
        public string? ContentBase64 { get; set; }

        /// <summary>
        /// ID for In-Line attachments (when retrieved).
        /// </summary>
        public string? ContentId { get; set; }

        /// <summary>
        /// Returns a friendly string for UI lists.
        /// </summary>
        public override string ToString() => Name ?? Id;
    }
}

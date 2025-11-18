using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace EmailSender.Models;

public class SendEmailRequest
{
    [Required]
    public List<string> To { get; set; } = new();

    public List<string>? Cc { get; set; }

    public List<string>? Bcc { get; set; }

    [Required]
    public string Subject { get; set; } = string.Empty;

    // HTML or Text. Defaults to HTML for flexibility.
    public string Body { get; set; } = string.Empty;

    // When using multipart/form-data, files will come via IFormFile.
    public List<IFormFile>? Attachments { get; set; }

    // Optional: treat body as HTML. Default true.
    public bool IsHtml { get; set; } = true;
}

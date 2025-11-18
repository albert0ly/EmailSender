using EmailSender.Models;
using EmailSender.Services;
using Microsoft.AspNetCore.Mvc;

namespace EmailSender.Controllers;

[ApiController]
[Route("api/[controller]")]
public class EmailController(IEmailSender emailSender) : ControllerBase
{
    private readonly IEmailSender _emailSender = emailSender;

    [HttpPost("send")] // multipart/form-data to support large attachments via streaming
    [DisableRequestSizeLimit]
    [RequestFormLimits(MultipartBodyLengthLimit = long.MaxValue, ValueLengthLimit = int.MaxValue)]
    public async Task<IActionResult> Send([FromForm] SendEmailRequest request, CancellationToken ct)
    {
        if (request.To == null || request.To.Count == 0)
            return BadRequest("At least one recipient is required.");

        var attachments = new List<(string FileName, string ContentType, Stream ContentStream)>();
        if (request.Attachments != null)
        {
            foreach (var file in request.Attachments)
            {
                if (file.Length > 0)
                {
                    var stream = file.OpenReadStream();
                    attachments.Add((file.FileName, string.IsNullOrWhiteSpace(file.ContentType) ? "application/octet-stream" : file.ContentType, stream));
                }
            }
        }

        await _emailSender.SendEmailAsync(
            request.To,
            request.Cc,
            request.Bcc,
            request.Subject,
            request.Body,
            request.IsHtml,
            attachments,
            ct);

        // Dispose streams after send
        foreach (var a in attachments)
        {
            a.ContentStream.Dispose();
        }

        return Accepted();
    }
}

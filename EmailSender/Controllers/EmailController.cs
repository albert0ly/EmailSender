using EmailSender.Models;
using EmailSender.Services;
using Microsoft.AspNetCore.Mvc;
using Azure.Identity;
using Azure.Core;
using Microsoft.Graph;
using Microsoft.Kiota.Authentication.Azure;
using Microsoft.Kiota.Http.HttpClientLibrary;
using Microsoft.Extensions.Options;

namespace EmailSender.Controllers;

[ApiController]
[Route("api/[controller]")]
public class EmailController(IEmailSender emailSender, IOptions<GraphOptions> graphOptions) : ControllerBase
{
    private readonly IEmailSender _emailSender = emailSender;
    private readonly GraphOptions _graphOptions = graphOptions.Value;

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

    /// <summary>
    /// Read new (unread) messages from the specified mailbox Inbox.
    /// Returns DTOs with all attachments (file content base64 encoded for file attachments).
    /// Marks messages as read so they are not returned again.
    /// </summary>
    /// <param name="mailbox">Mailbox UPN or id. If not provided, uses configured Graph:MailboxAddress.</param>
    [HttpPost("read")]
    public async Task<IActionResult> Read([FromQuery] string? mailbox, CancellationToken ct)
    {
        var user = string.IsNullOrWhiteSpace(mailbox) ? _graphOptions.MailboxAddress : mailbox!;
        if (string.IsNullOrWhiteSpace(user)) return BadRequest("Mailbox must be provided either as query parameter or in configuration.");

        // Build GraphServiceClient using same approach as GraphEmailSender
        var credential = new ClientSecretCredential(_graphOptions.TenantId, _graphOptions.ClientId, _graphOptions.ClientSecret);
        var authProvider = new AzureIdentityAuthenticationProvider(credential, new[] { "graph.microsoft.com" }, null, false, new[] { "https://graph.microsoft.com/.default" });
        var httpClient = GraphClientFactory.Create();
        var requestAdapter = new HttpClientRequestAdapter(authProvider, httpClient: httpClient);
        var graph = new GraphServiceClient(requestAdapter);

        // Query unread messages from Inbox
        var messagesResponse = await graph.Users[user].MailFolders["inbox"].Messages.GetAsync(rc =>
        {
            rc.QueryParameters.Filter = "isRead eq false";
            rc.QueryParameters.Select = new[] { "id", "subject", "body", "receivedDateTime", "isRead", "hasAttachments", "webLink", "toRecipients", "ccRecipients", "bccRecipients", "internetMessageHeaders" };
            rc.QueryParameters.Top = 100;
        }, cancellationToken: ct);

        var messages = messagesResponse?.Value ?? new List<Microsoft.Graph.Models.Message>();

        var result = new List<EmailSender.Models.MessageDto>();

        foreach (var msg in messages)
        {
            var dto = new EmailSender.Models.MessageDto
            {
                Id = msg.Id,
                Subject = msg.Subject,
                Body = msg.Body?.Content,
                ReceivedDateTime = msg.ReceivedDateTime,
                IsRead = msg.IsRead,
                HasAttachments = msg.HasAttachments,
                WebLink = msg.WebLink
            };

            // Populate recipients
            if (msg.ToRecipients != null)
            {
                foreach (var r in msg.ToRecipients)
                {
                    var addr = r?.EmailAddress?.Address;
                    if (!string.IsNullOrWhiteSpace(addr)) dto.To.Add(addr);
                }
            }
            if (msg.CcRecipients != null)
            {
                foreach (var r in msg.CcRecipients)
                {
                    var addr = r?.EmailAddress?.Address;
                    if (!string.IsNullOrWhiteSpace(addr)) dto.Cc.Add(addr);
                }
            }
            if (msg.BccRecipients != null)
            {
                foreach (var r in msg.BccRecipients)
                {
                    var addr = r?.EmailAddress?.Address;
                    if (!string.IsNullOrWhiteSpace(addr)) dto.Bcc.Add(addr);
                }
            }

            // Populate headers
            if (msg.InternetMessageHeaders != null)
            {
                foreach (var h in msg.InternetMessageHeaders)
                {
                    if (h == null) continue;
                    var name = h.Name ?? string.Empty;
                    var value = h.Value;
                    if (string.IsNullOrEmpty(name)) continue;
                    if (dto.Headers.ContainsKey(name))
                    {
                        dto.Headers[name] = string.Concat(dto.Headers[name], ",", value);
                    }
                    else
                    {
                        dto.Headers[name] = value;
                    }
                }
            }

            if (msg.HasAttachments == true)
            {
                try
                {
                    var atts = await graph.Users[user].Messages[msg.Id].Attachments.GetAsync(c => c.QueryParameters.Top = 200, cancellationToken: ct);
                    if (atts?.Value != null)
                    {
                        foreach (var a in atts.Value)
                        {
                            var adto = new EmailSender.Models.AttachmentDto
                            {
                                Id = a.Id,
                                Name = a.Name,
                                ContentType = a.AdditionalData != null && a.AdditionalData.ContainsKey("@odata.mediaContentType") ? a.AdditionalData["@odata.mediaContentType"]?.ToString() : a is Microsoft.Graph.Models.FileAttachment fa2 ? fa2.ContentType : null,
                                Size = a.Size,
                                IsInline = a.IsInline
                            };

                            if (a is Microsoft.Graph.Models.FileAttachment fa)
                            {
                                // ContentBytes is a byte[] property on FileAttachment
                                try
                                {
                                    if (fa.ContentBytes != null)
                                    {
                                        adto.ContentBase64 = Convert.ToBase64String(fa.ContentBytes);
                                    }
                                }
                                catch
                                {
                                    // ignore conversion errors
                                }
                            }

                            dto.Attachments.Add(adto);
                        }
                    }
                }
                catch
                {
                    // ignore per-message attachment errors
                }
            }

            // Mark message as read so it won't be returned again
            if (msg.IsRead != true)
            {
                try
                {
                    var update = new Microsoft.Graph.Models.Message { IsRead = true };
                    await graph.Users[user].Messages[msg.Id].PatchAsync(update, cancellationToken: ct);
                }
                catch
                {
                    // ignore errors marking as read
                }
            }

            result.Add(dto);
        }

        return Ok(result.ToArray());
    }
}

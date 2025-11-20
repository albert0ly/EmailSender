using Azure.Identity;
using Azure.Core;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.Messages.Item.Attachments.CreateUploadSession;
using Microsoft.Kiota.Authentication.Azure;
using Microsoft.Kiota.Http.HttpClientLibrary;
using System.Net.Http;
using System.Net.Http.Headers;

namespace EmailSender.Services;

public class GraphOptions
{
    public string TenantId { get; set; } = string.Empty;
    public string ClientId { get; set; } = string.Empty;
    public string ClientSecret { get; set; } = string.Empty;
    public string MailboxAddress { get; set; } = string.Empty; // sender mailbox
}

public interface IEmailSender
{
    Task SendEmailAsync(
        IEnumerable<string> to,
        IEnumerable<string>? cc,
        IEnumerable<string>? bcc,
        string subject,
        string body,
        bool isHtml,
        IEnumerable<(string FileName, string ContentType, Stream ContentStream)> attachments,
        CancellationToken cancellationToken = default);
}

public class GraphEmailSender : IEmailSender
{
    private readonly GraphOptions _options;
    private readonly GraphServiceClient _graph;

    public GraphEmailSender(IOptions<GraphOptions> options)
    {
        _options = options.Value;

        var credential = new ClientSecretCredential(
            _options.TenantId,
            _options.ClientId,
            _options.ClientSecret);
        var token = credential.GetToken(new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" }));
        Console.WriteLine($"Token acquired, length={token.Token?.Length}");

        // For Microsoft.Graph v5 with Kiota pipeline
        // AzureIdentityAuthenticationProvider constructor expects allowed hosts (without http/https) as the second argument.
        var authProvider = new AzureIdentityAuthenticationProvider(credential, new[] { "graph.microsoft.com" }, null, false, new[] { "https://graph.microsoft.com/.default" });
        var httpClient = GraphClientFactory.Create();
        var requestAdapter = new HttpClientRequestAdapter(authProvider, httpClient: httpClient);
        _graph = new GraphServiceClient(requestAdapter);
    }

    public async Task SendEmailAsync(
        IEnumerable<string> to,
        IEnumerable<string>? cc,
        IEnumerable<string>? bcc,
        string subject,
        string body,
        bool isHtml,
        IEnumerable<(string FileName, string ContentType, Stream ContentStream)> attachments,
        CancellationToken cancellationToken = default)
    {
        var message = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                ContentType = isHtml ? BodyType.Html : BodyType.Text,
                Content = body
            },
            ToRecipients = new List<Recipient>(to.Select(x => new Recipient { EmailAddress = new EmailAddress { Address = x } }))
        };

        if (cc != null)
            message.CcRecipients = cc.Select(x => new Recipient { EmailAddress = new EmailAddress { Address = x } }).ToList();
        if (bcc != null)
            message.BccRecipients = bcc.Select(x => new Recipient { EmailAddress = new EmailAddress { Address = x } }).ToList();

        // Create draft first
        var draft = await _graph.Users[_options.MailboxAddress].Messages.PostAsync(message, cancellationToken: cancellationToken);
        if (draft == null) throw new InvalidOperationException("Failed to create draft message.");

        // Attach large files via upload session if needed
        foreach (var att in attachments)
        {
            // If content length is unknown, force upload session approach
            var uploadSession = await _graph.Users[_options.MailboxAddress]
                .Messages[draft.Id]
                .Attachments
                .CreateUploadSession
                .PostAsync(new CreateUploadSessionPostRequestBody
                {
                    AttachmentItem = new AttachmentItem
                    {
                        AttachmentType = AttachmentType.File,
                        Name = att.FileName,
                        Size = att.ContentStream.CanSeek ? att.ContentStream.Length : (long?)null,
                        ContentType = att.ContentType
                    }
                }, cancellationToken: cancellationToken);

            if (uploadSession == null)
                throw new InvalidOperationException("Failed to create upload session for attachment.");

            await UploadToSessionAsync(uploadSession, att.ContentStream, cancellationToken);
        }

        // Send the message
        await _graph.Users[_options.MailboxAddress].Messages[draft.Id].Send.PostAsync(cancellationToken: cancellationToken);
    }

    private static async Task UploadToSessionAsync(UploadSession session, Stream content, CancellationToken ct)
    {
        if (session == null || string.IsNullOrEmpty(session.UploadUrl))
            throw new ArgumentException("Invalid upload session.");

        using var http = new HttpClient();

        const int chunkSize = 5 * 1024 * 1024; // 5 MB
        long start = 0;
        long? total = content.CanSeek ? content.Length : null;

        var buffer = new byte[chunkSize];
        int read;
        while ((read = await content.ReadAsync(buffer.AsMemory(0, chunkSize), ct)) > 0)
        {
            var end = start + read - 1;

            using var request = new HttpRequestMessage(HttpMethod.Put, session.UploadUrl);
            request.Content = new ByteArrayContent(buffer, 0, read);
            request.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            request.Content.Headers.ContentLength = read;
            var totalText = total.HasValue ? total.Value.ToString() : "*";
            // Move Content-Range header to content headers so it is sent correctly.
            request.Content.Headers.TryAddWithoutValidation("Content-Range", $"bytes {start}-{end}/{totalText}");

            var response = await http.SendAsync(request, ct);
            if ((int)response.StatusCode == 201 || (int)response.StatusCode == 200)
            {
                // Completed
                break;
            }
            if ((int)response.StatusCode != 202)
            {
                var body = await response.Content.ReadAsStringAsync(ct);
                throw new InvalidOperationException($"Chunk upload failed: {(int)response.StatusCode} {response.ReasonPhrase} - {body}");
            }

            start = end + 1;
        }
    }
}

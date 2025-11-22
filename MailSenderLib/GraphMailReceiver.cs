using Azure.Core;
using Azure.Identity;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace MailSenderLib
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
        /// Returns a friendly string for UI lists.
        /// </summary>
        public override string ToString() => Name ?? Id;
    }

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

    /// <summary>
    /// Simple Graph mail receiver using client credentials and REST APIs.
    /// </summary>
    public class GraphMailReceiver
    {
        private readonly GraphMailOptions _options;
        private readonly ClientSecretCredential _credential;
        private static readonly Uri GraphBaseUri = new Uri("https://graph.microsoft.com/v1.0/");
        private static readonly string[] GraphScopes = new[] { "https://graph.microsoft.com/.default" };

        /// <summary>
        /// Creates a new instance with the provided Graph options.
        /// </summary>
        /// <param name="options">Graph authentication and mailbox options.</param>
        public GraphMailReceiver(GraphMailOptions options)
        {
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _credential = new ClientSecretCredential(_options.TenantId, _options.ClientId, _options.ClientSecret);
        }

        /// <summary>
        /// Receives unread messages from the mailbox inbox, fetches attachments and marks messages as read.
        /// </summary>
        /// <param name="mailbox">Mailbox UPN or id. If null or empty, uses configured MailboxAddress.</param>
        /// <param name="ct">Cancellation token.</param>
        /// <returns>List of messages with attachments.</returns>
        public async Task<List<MailMessageDto>> ReceiveEmailsAsync(string? mailbox, CancellationToken ct = default)
        {
            var user = string.IsNullOrWhiteSpace(mailbox) ? _options.MailboxAddress : mailbox!;
            if (string.IsNullOrWhiteSpace(user)) throw new ArgumentException("Mailbox must be provided.", nameof(mailbox));

            var token = await _credential.GetTokenAsync(new TokenRequestContext(GraphScopes), ct);

            using var http = new HttpClient() { BaseAddress = GraphBaseUri };
            http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

            var select = "id,subject,body,receivedDateTime,isRead,hasAttachments,webLink,toRecipients,ccRecipients,bccRecipients,internetMessageHeaders";
            var url = $"users/{Uri.EscapeDataString(user)}/mailFolders/inbox/messages?$filter=isRead eq false&$select={Uri.EscapeDataString(select)}&$top=100";

            var resp = await http.GetAsync(url, ct);
            if (!resp.IsSuccessStatusCode)
            {
                var body = await resp.Content.ReadAsStringAsync();
                throw new InvalidOperationException($"Failed to list messages: {(int)resp.StatusCode} {resp.ReasonPhrase} - {body}");
            }

            var json = await resp.Content.ReadAsStringAsync();
            var root = JObject.Parse(json);
            var array = root.Value<JArray>("value");
            var result = new List<MailMessageDto>();

            if (array == null) return result;

            foreach (var item in array)
            {
                var id = item.Value<string>("id") ?? string.Empty;

                // parse receivedDateTime safely
                DateTimeOffset? received = null;
                var receivedToken = item.SelectToken("receivedDateTime");
                if (receivedToken != null)
                {
                    var s = receivedToken.Type == JTokenType.String ? receivedToken.ToString() : receivedToken.ToString(Newtonsoft.Json.Formatting.None);
                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        if (DateTimeOffset.TryParse(s, out var dto))
                        {
                            received = dto;
                        }
                        else if (DateTime.TryParse(s, out var dt))
                        {
                            // convert DateTime to DateTimeOffset assuming unspecified kind as local
                            received = new DateTimeOffset(dt);
                        }
                    }
                }

                var msg = new MailMessageDto
                {
                    Id = id,
                    Subject = item.Value<string>("subject"),
                    Body = item.SelectToken("body.content")?.ToString(),
                    ReceivedDateTime = received,
                    IsRead = item.Value<bool?>("isRead"),
                    HasAttachments = item.Value<bool?>("hasAttachments"),
                    WebLink = item.Value<string>("webLink")
                };

                // recipients
                var toArr = item.Value<JArray>("toRecipients");
                if (toArr != null)
                {
                    foreach (var r in toArr)
                    {
                        var addr = r.SelectToken("emailAddress.address")?.ToString();
                        if (!string.IsNullOrWhiteSpace(addr)) msg.To.Add(addr!);
                    }
                }
                var ccArr = item.Value<JArray>("ccRecipients");
                if (ccArr != null)
                {
                    foreach (var r in ccArr)
                    {
                        var addr = r.SelectToken("emailAddress.address")?.ToString();
                        if (!string.IsNullOrWhiteSpace(addr)) msg.Cc.Add(addr!);
                    }
                }
                var bccArr = item.Value<JArray>("bccRecipients");
                if (bccArr != null)
                {
                    foreach (var r in bccArr)
                    {
                        var addr = r.SelectToken("emailAddress.address")?.ToString();
                        if (!string.IsNullOrWhiteSpace(addr)) msg.Bcc.Add(addr!);
                    }
                }

                // headers
                var headersArr = item.Value<JArray>("internetMessageHeaders");
                if (headersArr != null)
                {
                    foreach (var h in headersArr)
                    {
                        var name = h.Value<string>("name");
                        var value = h.Value<string>("value");
                        if (string.IsNullOrWhiteSpace(name)) continue;
                        var key = name!;
                        // Use TryGetValue to avoid duplicate dictionary lookup (fix CA1854)
                        if (msg.Headers.TryGetValue(key, out var existing))
                        {
                            msg.Headers[key] = string.Concat(existing, ",", value);
                        }
                        else
                        {
                            msg.Headers[key] = value;
                        }
                    }
                }

                // attachments
                if (msg.HasAttachments == true)
                {
                    try
                    {
                        var attUrl = $"users/{Uri.EscapeDataString(user)}/messages/{Uri.EscapeDataString(id)}/attachments";
                        var attResp = await http.GetAsync(attUrl, ct);
                        if (attResp.IsSuccessStatusCode)
                        {
                            var attJson = await attResp.Content.ReadAsStringAsync();
                            var attRoot = JObject.Parse(attJson);
                            var attArray = attRoot.Value<JArray>("value");
                            if (attArray != null)
                            {
                                foreach (var a in attArray)
                                {
                                    var adto = new MailAttachmentDto
                                    {
                                        Id = a.Value<string>("id") ?? string.Empty,
                                        Name = a.Value<string>("name"),
                                        ContentType = a.Value<string>("contentType") ?? a.SelectToken("@odata.mediaContentType")?.ToString(),
                                        Size = a.Value<long?>("size"),
                                        IsInline = a.Value<bool?>("isInline")
                                    };

                                    var contentBytes = a.Value<string>("contentBytes");
                                    if (!string.IsNullOrEmpty(contentBytes))
                                    {
                                        // contentBytes is base64 string already
                                        adto.ContentBase64 = contentBytes;
                                    }

                                    msg.Attachments.Add(adto);
                                }
                            }
                        }
                    }
                    catch
                    {
                        // ignore attachment fetch errors per-message
                    }
                }

                // mark as read
                if (msg.IsRead != true)
                {
                    try
                    {
                        var patch = new JObject { ["isRead"] = true };
                        var patchContent = new StringContent(patch.ToString(), System.Text.Encoding.UTF8, "application/json");
                        using (var patchReq = new HttpRequestMessage(new HttpMethod("PATCH"), $"users/{Uri.EscapeDataString(user)}/messages/{Uri.EscapeDataString(id)}") { Content = patchContent })
                        {
                            // send and ignore
                            await http.SendAsync(patchReq, ct);
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                }

                result.Add(msg);
            }

            return result;
        }
    }
}

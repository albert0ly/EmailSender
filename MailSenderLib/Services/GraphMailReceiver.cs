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
using MailSenderLib.Models;

namespace MailSenderLib.Services
{
    /// <summary>
    /// Simple Graph mail receiver using client credentials and REST APIs.
    /// </summary>
    public class GraphMailReceiver : Interfaces.IGraphMailReceiver
    {
        private readonly MailSenderLib.Options.GraphMailOptions _options;
        private readonly ClientSecretCredential _credential;
        private readonly HttpClient? _httpClient;
        private static readonly Uri GraphBaseUri = new Uri("https://graph.microsoft.com/v1.0/");
        private static readonly string[] GraphScopes = new[] { "https://graph.microsoft.com/.default" };

        /// <summary>
        /// Creates a new instance with the provided Graph options.
        /// </summary>
        /// <param name="options">Graph authentication and mailbox options.</param>
        /// <param name="httpClient">Optional HttpClient for testing/DI. If not provided a new HttpClient will be created per call.</param>
        public GraphMailReceiver(MailSenderLib.Options.GraphMailOptions options, HttpClient? httpClient = null)
        {
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _credential = new ClientSecretCredential(_options.TenantId, _options.ClientId, _options.ClientSecret);
            _httpClient = httpClient;
        }

        /// <inheritdoc />
        public async Task<List<MailMessageDto>> ReceiveEmailsAsync(string? mailbox, CancellationToken ct = default)
        {
            var user = string.IsNullOrWhiteSpace(mailbox) ? _options.MailboxAddress : mailbox!;
            if (string.IsNullOrWhiteSpace(user)) throw new ArgumentException("Mailbox must be provided.", nameof(mailbox));

            var token = await _credential.GetTokenAsync(new TokenRequestContext(GraphScopes), ct);

            using var http = _httpClient ?? new HttpClient();
            http.BaseAddress = GraphBaseUri;
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

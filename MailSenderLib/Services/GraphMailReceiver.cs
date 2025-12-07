using Azure.Core;
using Azure.Identity;
using MailSenderLib.Models;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace MailSenderLib.Services
{
    /// <summary>
    /// Simple Graph mail receiver using client credentials and REST APIs.
    /// </summary>
    public class GraphMailReceiver : Interfaces.IGraphMailReceiver, IDisposable
    {
        private readonly MailSenderLib.Options.GraphMailOptionsAuth _options;
        private readonly ClientSecretCredential _credential;
        private readonly HttpClient? _httpClient;
        private readonly ILogger<GraphMailReceiver>? _logger;
        private static readonly Uri GraphBaseUri = new Uri("https://graph.microsoft.com/v1.0/");
        private static readonly string[] GraphScopes = new[] { "https://graph.microsoft.com/.default" };

        // Cached token and lock for refresh
        private AccessToken _cachedToken;
        private readonly SemaphoreSlim _tokenLock = new SemaphoreSlim(1, 1);
        private static readonly TimeSpan TokenExpiryBuffer = TimeSpan.FromSeconds(60);

        // LoggerMessage delegates
        private static readonly Action<ILogger, Exception?> _logFailedToAcquireToken =
            LoggerMessage.Define(LogLevel.Error, new EventId(2000, nameof(_logFailedToAcquireToken)), "Failed to acquire access token for GraphMailReceiver");
        private static readonly Action<ILogger, Exception?> _logRefreshingToken =
            LoggerMessage.Define(LogLevel.Debug, new EventId(2001, nameof(_logRefreshingToken)), "Refreshing access token for GraphMailReceiver");
        private static readonly Action<ILogger, DateTimeOffset, Exception?> _logTokenAcquired =
            LoggerMessage.Define<DateTimeOffset>(LogLevel.Debug, new EventId(2002, nameof(_logTokenAcquired)), "Access token acquired, expires on {ExpiresOn}");
        private static readonly Action<ILogger, int, string, string, Exception?> _logFailedToListMessages =
            LoggerMessage.Define<int, string, string>(LogLevel.Error, new EventId(2003, nameof(_logFailedToListMessages)), "Failed to list messages: {Status} {Reason} - {Body}");
        private static readonly Action<ILogger, string, Exception?> _logFailedToFetchAttachments =
            LoggerMessage.Define<string>(LogLevel.Warning, new EventId(2004, nameof(_logFailedToFetchAttachments)), "Failed to fetch attachments for message {MessageId}");

        /// <summary>
        /// Creates a new instance with the provided Graph options.
        /// </summary>
        /// <param name="options">Graph authentication and mailbox options.</param>
        /// <param name="httpClient">Optional HttpClient for testing/DI. If not provided a new HttpClient will be created per call.</param>
        public GraphMailReceiver(MailSenderLib.Options.GraphMailOptionsAuth options, HttpClient? httpClient = null, object? logger = null)
        {
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _credential = new ClientSecretCredential(_options.TenantId, _options.ClientId, _options.ClientSecret);
            _httpClient = httpClient;
            _logger = logger as ILogger<GraphMailReceiver>;
        }

        // Return a cached token if valid, otherwise refresh in a thread-safe manner
        private async Task<AccessToken> GetAccessTokenAsync(CancellationToken ct)
        {
            if (!string.IsNullOrEmpty(_cachedToken.Token) && _cachedToken.ExpiresOn > DateTimeOffset.UtcNow.Add(TokenExpiryBuffer))
            {
                return _cachedToken;
            }

            await _tokenLock.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                if (!string.IsNullOrEmpty(_cachedToken.Token) && _cachedToken.ExpiresOn > DateTimeOffset.UtcNow.Add(TokenExpiryBuffer))
                {
                    return _cachedToken;
                }

                if (_logger != null) _logRefreshingToken(_logger, null);
                var token = await _credential.GetTokenAsync(new TokenRequestContext(GraphScopes), ct).ConfigureAwait(false);
                _cachedToken = token;
                if (_logger != null) _logTokenAcquired(_logger, _cachedToken.ExpiresOn, null);
                return _cachedToken;
            }
            catch (Exception ex)
            {
                if (_logger != null) _logFailedToAcquireToken(_logger, ex);
                throw;
            }
            finally
            {
                _tokenLock.Release();
            }
        }

        /// <inheritdoc />
        public async Task<List<MailMessageDto>> ReceiveEmailsAsync(string? mailbox, CancellationToken ct = default)
        {
            var user = string.IsNullOrWhiteSpace(mailbox) ? _options.MailboxAddress : mailbox!;
            if (string.IsNullOrWhiteSpace(user)) throw new ArgumentException("Mailbox must be provided.", nameof(mailbox));

            // Acquire token (cached)
            var token = await GetAccessTokenAsync(ct).ConfigureAwait(false);

            using var http = _httpClient ?? new HttpClient();
            http.BaseAddress = GraphBaseUri;
            http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

            var select = "id,subject,body,receivedDateTime,isRead,hasAttachments,webLink,toRecipients,ccRecipients,bccRecipients,internetMessageHeaders";
            var url = $"users/{Uri.EscapeDataString(user)}/mailFolders/inbox/messages?$filter=isRead eq false&$select={Uri.EscapeDataString(select)}&$top=100";

            var resp = await http.GetAsync(url, ct).ConfigureAwait(false);
            if (!resp.IsSuccessStatusCode)
            {
                var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (_logger != null) _logFailedToListMessages(_logger, (int)resp.StatusCode, resp.ReasonPhrase, body, null);
                throw new InvalidOperationException($"Failed to list messages: {(int)resp.StatusCode} {resp.ReasonPhrase} - {body}");
            }

            var json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
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
                        var attResp = await http.GetAsync(attUrl, ct).ConfigureAwait(false);
                        if (attResp.IsSuccessStatusCode)
                        {
                            var attJson = await attResp.Content.ReadAsStringAsync().ConfigureAwait(false);
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
                    catch (Exception ex)
                    {
                        if (_logger != null) _logFailedToFetchAttachments(_logger, id, ex);
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
                            await http.SendAsync(patchReq, ct).ConfigureAwait(false);
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




        public void Dispose()
        {
            try
            {
                _tokenLock?.Dispose();
                GC.SuppressFinalize(this);
            }
            catch
            {
                // ignore
            }
        }
    }
}

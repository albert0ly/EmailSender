using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace MailSenderLib.Services
{
    /// <summary>
    /// Lightweight Graph email sender usable from .NET Standard2.0 via Graph REST APIs and Azure.Identity.
    /// </summary>
    public class GraphMailSender : Interfaces.IGraphMailSender, IDisposable
    {
        private readonly MailSenderLib.Options.GraphMailOptions _options;
        private readonly ClientSecretCredential _credential;
        private readonly ILogger<GraphMailSender>? _logger;
        private static readonly Uri GraphBaseUri = new Uri("https://graph.microsoft.com/v1.0/");
        private static readonly string[] scopes = { "https://graph.microsoft.com/.default" };

        // Cached token and lock for refresh
        private AccessToken _cachedToken;
        private readonly SemaphoreSlim _tokenLock = new SemaphoreSlim(1, 1);
        // safety buffer before expiry to force refresh
        private static readonly TimeSpan TokenExpiryBuffer = TimeSpan.FromSeconds(60);

        // LoggerMessage delegates (avoid allocation-heavy LoggerExtensions calls)
        private static readonly Action<ILogger, Exception?> _logFailedToAcquireToken =
            LoggerMessage.Define(LogLevel.Error, new EventId(1000, nameof(_logFailedToAcquireToken)), "Failed to acquire access token for GraphMailSender");
        private static readonly Action<ILogger, Exception?> _logRefreshingToken =
            LoggerMessage.Define(LogLevel.Debug, new EventId(1001, nameof(_logRefreshingToken)), "Refreshing access token for GraphMailSender");
        private static readonly Action<ILogger, DateTimeOffset, Exception?> _logTokenAcquired =
            LoggerMessage.Define<DateTimeOffset>(LogLevel.Debug, new EventId(1002, nameof(_logTokenAcquired)), "Access token acquired, expires on {ExpiresOn}");
        private static readonly Action<ILogger, string, Exception?> _logUploadSessionUrl =
            LoggerMessage.Define<string>(LogLevel.Debug, new EventId(1013, nameof(_logUploadSessionUrl)), "Upload session URL: {Url}");
        private static readonly Action<ILogger, long, long, long, int, Exception?> _logChunkStatus =
            LoggerMessage.Define<long, long, long, int>(LogLevel.Debug, new EventId(1010, nameof(_logChunkStatus)), "Chunk {Start}-{End}/{Total}, Status {Status}");
        private static readonly Action<ILogger, string, Exception?> _logResponseBodyTrace =
            LoggerMessage.Define<string>(LogLevel.Trace, new EventId(1011, nameof(_logResponseBodyTrace)), "{Body}");
        private static readonly Action<ILogger, string, Exception?> _logUploadComplete =
            LoggerMessage.Define<string>(LogLevel.Debug, new EventId(1012, nameof(_logUploadComplete)), "Upload complete for {FileName}");
        private static readonly Action<ILogger, int, string, string, Exception?> _logChunkFailed =
            LoggerMessage.Define<int, string, string>(LogLevel.Error, new EventId(1014, nameof(_logChunkFailed)), "Chunk upload failed {Status} {Reason} - {Body}");

        public GraphMailSender(MailSenderLib.Options.GraphMailOptions options, object? logger = null)
        {
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _credential = new ClientSecretCredential(_options.TenantId, _options.ClientId, _options.ClientSecret);
            _logger = logger as ILogger<GraphMailSender>;
        }

        private static async Task EnsureSuccess(HttpResponseMessage resp, string action, CancellationToken ct)
        {
            if (!resp.IsSuccessStatusCode)
            {
                var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                throw new InvalidOperationException($"Failed to {action}: {(int)resp.StatusCode} {resp.ReasonPhrase} - {body}");
            }
        }

        // Return a cached token if valid, otherwise refresh in a thread-safe manner
        private async Task<AccessToken> GetAccessTokenAsync(CancellationToken ct)
        {
            // fast-path without locking
            if (!string.IsNullOrEmpty(_cachedToken.Token) && _cachedToken.ExpiresOn > DateTimeOffset.UtcNow.Add(TokenExpiryBuffer))
            {
                return _cachedToken;
            }

            await _tokenLock.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                // re-check after acquiring lock
                if (!string.IsNullOrEmpty(_cachedToken.Token) && _cachedToken.ExpiresOn > DateTimeOffset.UtcNow.Add(TokenExpiryBuffer))
                {
                    return _cachedToken;
                }

                if (_logger != null) _logRefreshingToken(_logger, null);
                var token = await _credential.GetTokenAsync(new TokenRequestContext(scopes), ct).ConfigureAwait(false);
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

        public async Task SendEmailAsync(
            IEnumerable<string> toRecipients,
            IEnumerable<string>? ccRecipients,
            IEnumerable<string>? bccRecipients,
            string subject,
            string body,
            bool isHtml,
            IEnumerable<(string FileName, string ContentType, Stream ContentStream)>? attachments,
            CancellationToken cancellationToken = default)
        {
            if (toRecipients == null) throw new ArgumentNullException(nameof(toRecipients));
            var toList = new List<string>(toRecipients);
            var ccList = ccRecipients != null ? new List<string>(ccRecipients) : new List<string>();
            var bccList = bccRecipients != null ? new List<string>(bccRecipients) : new List<string>();
            if (toList.Count == 0) throw new ArgumentException("At least one recipient is required.", nameof(toRecipients));

            // Acquire token (cached)
            var token = await GetAccessTokenAsync(cancellationToken).ConfigureAwait(false);

            using (var http = new HttpClient() { BaseAddress = GraphBaseUri })
            {
                http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

                //1. Create draft message
                var draftPayload = BuildCreateMessagePayload(toList, ccList, bccList, subject, body, isHtml);
                var draftResp = await http.PostAsync($"users/{Uri.EscapeDataString(_options.MailboxAddress)}/messages", new StringContent(draftPayload, System.Text.Encoding.UTF8, "application/json"), cancellationToken).ConfigureAwait(false);
                await EnsureSuccess(draftResp, "create draft", cancellationToken).ConfigureAwait(false);
                var draftJson = await draftResp.Content.ReadAsStringAsync().ConfigureAwait(false);
                var draftIdMaybe = Newtonsoft.Json.Linq.JObject.Parse(draftJson).Value<string>("id");
                if (string.IsNullOrWhiteSpace(draftIdMaybe)) throw new InvalidOperationException("Failed to obtain draft id.");
                string draftId = draftIdMaybe!;

                //2. Upload attachments using upload session (handles large files and unknown lengths)
                if (attachments != null)
                {
                    foreach (var att in attachments)
                    {
                        await UploadAttachmentAsync(http, _options.MailboxAddress, draftId!, att.FileName, att.ContentType, att.ContentStream, cancellationToken).ConfigureAwait(false);
                    }
                }

                //3. Send the message
                var sendResp = await http.PostAsync($"users/{Uri.EscapeDataString(_options.MailboxAddress)}/messages/{Uri.EscapeDataString(draftId)}/send", new StringContent("{}", System.Text.Encoding.UTF8, "application/json"), cancellationToken).ConfigureAwait(false);
                await EnsureSuccess(sendResp, "send message", cancellationToken).ConfigureAwait(false);
            }
        }

        private static string BuildCreateMessagePayload(List<string> to, List<string> cc, List<string> bcc, string subject, string body, bool isHtml)
        {
            string Escape(string s) => s?.Replace("\\", "\\\\").Replace("\"", "\\\"") ?? string.Empty;

            string Recipients(IEnumerable<string> addrs)
                => string.Join(",", addrs.Where(a => !string.IsNullOrWhiteSpace(a)).Select(a => "{\"emailAddress\":{\"address\":\"" + Escape(a.Trim()) + "\"}}"));

            var contentType = isHtml ? "HTML" : "Text";
            var json = $"{{\n \"subject\": \"{Escape(subject)}\",\n \"body\": {{ \"contentType\": \"{contentType}\", \"content\": \"{Escape(body)}\" }},\n \"toRecipients\": [ {Recipients(to)} ],\n \"ccRecipients\": [ {Recipients(cc)} ],\n \"bccRecipients\": [ {Recipients(bcc)} ]\n}}";
            return json;
        }

        private async Task UploadAttachmentAsync(
            HttpClient http,
            string mailbox,
            string draftId,
            string fileName,
            string contentType,
            Stream content,
            CancellationToken ct)
        {
            if (!content.CanSeek)
                throw new InvalidOperationException("Stream must support seeking for chunked upload.");

            long totalSize = content.Length;
            content.Position = 0;

            // -------------------------------
            // 1. Create upload session
            // -------------------------------
            var payload = new
            {
                attachmentItem = new
                {
                    attachmentType = "file",
                    name = fileName,
                    size = totalSize,
                    contentType = string.IsNullOrWhiteSpace(contentType)
                        ? "application/octet-stream"
                        : contentType
                }
            };

            var startSessionJson = Newtonsoft.Json.JsonConvert.SerializeObject(payload);
            var sessionResp = await http.PostAsync(
                $"users/{Uri.EscapeDataString(mailbox)}/messages/{Uri.EscapeDataString(draftId)}/attachments/createUploadSession",
                new StringContent(startSessionJson, System.Text.Encoding.UTF8, "application/json"),
                ct
            ).ConfigureAwait(false);

            await EnsureSuccess(sessionResp, "create upload session", ct).ConfigureAwait(false);

            var sessionJson = await sessionResp.Content.ReadAsStringAsync().ConfigureAwait(false);
            var uploadUrl = Newtonsoft.Json.Linq.JObject.Parse(sessionJson).Value<string>("uploadUrl");

            if (string.IsNullOrWhiteSpace(uploadUrl))
                throw new InvalidOperationException("Upload session URL not found.");

            if (_logger != null) _logUploadSessionUrl(_logger, uploadUrl, null);

            // -------------------------------
            // 2. Upload chunks
            // -------------------------------
            const int chunkSize = 5 * 1024 * 1024; // 5 MB
            byte[] buffer = new byte[chunkSize];
            long start = 0;

            using (var uploadClient = new HttpClient()) // fresh client
            {
                while (start < totalSize)
                {
                    int read = await content.ReadAsync(buffer, 0, buffer.Length, ct).ConfigureAwait(false);
                    if (read <= 0) break;

                    long end = start + read - 1;

                    using (var put = new HttpRequestMessage(HttpMethod.Put, new Uri(uploadUrl, UriKind.Absolute)))
                    {
                        put.Content = new ByteArrayContent(buffer, 0, read);
                        put.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                        put.Content.Headers.ContentLength = read;
                        // IMPORTANT: Content-Range must be a content header so proxies/HttpClient don't strip it
                        put.Content.Headers.TryAddWithoutValidation("Content-Range", $"bytes {start}-{end}/{totalSize}");

                        var resp = await uploadClient.SendAsync(put, ct).ConfigureAwait(false);
                        var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);

                        if (_logger != null) _logChunkStatus(_logger, start, end, totalSize, (int)resp.StatusCode, null);
                        if (_logger != null) _logResponseBodyTrace(_logger, body, null);

                        if ((int)resp.StatusCode == 200 || (int)resp.StatusCode == 201)
                        {
                            if (_logger != null) _logUploadComplete(_logger, fileName, null);
                            return;
                        }
                        else if ((int)resp.StatusCode != 202)
                        {
                            if (_logger != null) _logChunkFailed(_logger, (int)resp.StatusCode, resp.ReasonPhrase, body, null);
                            throw new InvalidOperationException(
                                $"Chunk upload failed {(int)resp.StatusCode} {resp.ReasonPhrase} - {body}"
                            );
                        }

                        start = end + 1;
                    }
                }
            }

            if (_logger != null) _logUploadComplete(_logger, fileName, null);
        }

        public void Dispose()
        {
            try
            {
                _tokenLock?.Dispose();
            }
            catch
            {
                // ignore
            }
        }
    }
}

using Azure.Core;
using Azure.Identity;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace MailSenderLib
{
    /// <summary>
    /// Options for authenticating and sending email via Microsoft Graph.
    /// </summary>
    public class GraphMailOptions
    {
        /// <summary>
        /// Azure AD tenant ID (GUID).
        /// </summary>
        public string TenantId { get; set; } = string.Empty;
        /// <summary>
        /// Application (client) ID of the app registration.
        /// </summary>
        public string ClientId { get; set; } = string.Empty;
        /// <summary>
        /// Client secret for the app registration.
        /// </summary>
        public string ClientSecret { get; set; } = string.Empty;
        /// <summary>
        /// Sender mailbox UPN or user ID used for the send action.
        /// </summary>
        public string MailboxAddress { get; set; } = string.Empty; // sender mailbox UPN or id
    }

    /// <summary>
    /// Abstraction for sending emails using Microsoft Graph.
    /// </summary>
    public interface IGraphMailSender
    {
        /// <summary>
        /// Sends an email using Microsoft Graph with optional CC, BCC and attachments.
        /// </summary>
        /// <param name="to">One or more primary recipients.</param>
        /// <param name="cc">Optional list of CC recipients.</param>
        /// <param name="bcc">Optional list of BCC recipients.</param>
        /// <param name="subject">Email subject.</param>
        /// <param name="body">Email body content.</param>
        /// <param name="isHtml">If true, body is treated as HTML; otherwise plain text.</param>
        /// <param name="attachments">Zero or more attachments as filename, content type and stream.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        Task SendEmailAsync(
            IEnumerable<string> to,
            IEnumerable<string> cc,
            IEnumerable<string> bcc,
            string subject,
            string body,
            bool isHtml,
            IEnumerable<(string FileName, string ContentType, Stream ContentStream)> attachments,
            CancellationToken cancellationToken = default);
    }

    /// <summary>
    /// Lightweight Graph email sender usable from .NET Standard2.0 via Graph REST APIs and Azure.Identity.
    /// </summary>
    public class GraphMailSender : IGraphMailSender
    {
        private readonly GraphMailOptions _options;
        private readonly ClientSecretCredential _credential;
        private static readonly Uri GraphBaseUri = new Uri("https://graph.microsoft.com/v1.0/");
        private static readonly string[] scopes = new[] { "https://graph.microsoft.com/.default" };

        /// <summary>
        /// Creates a new GraphMailSender.
        /// </summary>
        /// <param name="options">Graph configuration including tenant, client and sender mailbox.</param>
        public GraphMailSender(GraphMailOptions options)
        {
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _credential = new ClientSecretCredential(_options.TenantId, _options.ClientId, _options.ClientSecret);
        }

        // EnsureSuccess moved above SendEmailAsync so the method is in scope at call sites.
        private static async Task EnsureSuccess(HttpResponseMessage resp, CancellationToken ct, string action)
        {
            if (!resp.IsSuccessStatusCode)
            {
                var body = await resp.Content.ReadAsStringAsync();
                throw new InvalidOperationException($"Failed to {action}: {(int)resp.StatusCode} {resp.ReasonPhrase} - {body}");
            }
        }

        /// <inheritdoc />
        public async Task SendEmailAsync(
            IEnumerable<string> to,
            IEnumerable<string> cc,
            IEnumerable<string> bcc,
            string subject,
            string body,
            bool isHtml,
            IEnumerable<(string FileName, string ContentType, Stream ContentStream)> attachments,
            CancellationToken cancellationToken = default)
        {
            if (to == null) throw new ArgumentNullException(nameof(to));
            var toList = new List<string>(to);
            var ccList = cc != null ? new List<string>(cc) : new List<string>();
            var bccList = bcc != null ? new List<string>(bcc) : new List<string>();
            if (toList.Count == 0) throw new ArgumentException("At least one recipient is required.", nameof(to));

            // Acquire token
            var token = await _credential.GetTokenAsync(new TokenRequestContext(scopes), cancellationToken);

            using (var http = new HttpClient() { BaseAddress = GraphBaseUri })
            {
                http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

                //1. Create draft message
                var draftPayload = BuildCreateMessagePayload(toList, ccList, bccList, subject, body, isHtml);
                var draftResp = await http.PostAsync($"users/{Uri.EscapeDataString(_options.MailboxAddress)}/messages", new StringContent(draftPayload, System.Text.Encoding.UTF8, "application/json"), cancellationToken);
                await EnsureSuccess(draftResp, cancellationToken, "create draft");
                var draftJson = await draftResp.Content.ReadAsStringAsync();
                var draftIdMaybe = Newtonsoft.Json.Linq.JObject.Parse(draftJson).Value<string>("id");
                if (string.IsNullOrWhiteSpace(draftIdMaybe)) throw new InvalidOperationException("Failed to obtain draft id.");
                string draftId = draftIdMaybe!;

                //2. Upload attachments using upload session (handles large files and unknown lengths)
                if (attachments != null)
                {
                    foreach (var att in attachments)
                    {
                        await UploadAttachmentAsync(http, _options.MailboxAddress, draftId!, att.FileName, att.ContentType, att.ContentStream, cancellationToken);
                    }
                }

                //3. Send the message
                var sendResp = await http.PostAsync($"users/{Uri.EscapeDataString(_options.MailboxAddress)}/messages/{Uri.EscapeDataString(draftId)}/send", new StringContent("{}", System.Text.Encoding.UTF8, "application/json"), cancellationToken);
                await EnsureSuccess(sendResp, cancellationToken, "send message");
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

        private static async Task UploadAttachmentAsync(
            HttpClient http,
            string mailbox,
            string draftId,
            string fileName,
            string contentType,
            Stream content,
            CancellationToken ct)
        {
            if (content == null) throw new ArgumentNullException(nameof(content));
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
            );

            await EnsureSuccess(sessionResp, ct, "create upload session");

            var sessionJson = await sessionResp.Content.ReadAsStringAsync();
            var uploadUrl = Newtonsoft.Json.Linq.JObject.Parse(sessionJson).Value<string>("uploadUrl");

            if (string.IsNullOrWhiteSpace(uploadUrl))
                throw new InvalidOperationException("Upload session URL not found.");

            Console.WriteLine("Upload session URL: " + uploadUrl);

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
                    int read = await content.ReadAsync(buffer, 0, buffer.Length, ct);
                    if (read <= 0) break;

                    long end = start + read - 1;

                    using (var put = new HttpRequestMessage(HttpMethod.Put, new Uri(uploadUrl, UriKind.Absolute)))
                    {
                        put.Content = new ByteArrayContent(buffer, 0, read);
                        put.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                        put.Content.Headers.ContentLength = read;
                        // IMPORTANT: Content-Range must be a content header so proxies/HttpClient don't strip it
                        put.Content.Headers.TryAddWithoutValidation("Content-Range", $"bytes {start}-{end}/{totalSize}");

                        var resp = await uploadClient.SendAsync(put, ct);
                        var body = await resp.Content.ReadAsStringAsync();

                        Console.WriteLine($"Chunk {start}-{end}/{totalSize}, Status: {(int)resp.StatusCode}");
                        Console.WriteLine(body);

                        if ((int)resp.StatusCode == 200 || (int)resp.StatusCode == 201)
                        {
                            Console.WriteLine("Upload complete!");
                            return;
                        }
                        else if ((int)resp.StatusCode != 202)
                        {
                            throw new InvalidOperationException(
                                $"Chunk upload failed {(int)resp.StatusCode} {resp.ReasonPhrase} - {body}"
                            );
                        }

                        start = end + 1;
                    }
                }
            }

            Console.WriteLine("All chunks uploaded successfully.");
        }
    }
}

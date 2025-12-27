using Azure.Core;
using Azure.Identity;
using Codeuctivity;
using MailSenderLib.Exceptions;
using MailSenderLib.Extensions;
using MailSenderLib.Interfaces;
using MailSenderLib.Logging;
using MailSenderLib.Models;
using MailSenderLib.Options;
using MailSenderLib.Utils;
using Microsoft.AspNetCore.StaticFiles;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Polly;
using Polly.Contrib.WaitAndRetry;
using System;
using System.Buffers;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MailSenderLib.Services
{
    public class GraphMailSender : IDisposable, IGraphMailSender
    {
        private const int JsonStreamBufferSize = 8192;
        private const int FileStreamBufferSize = 8192;
        private const long LargeAttachmentThreshold = 3 * 1024 * 1024; // 3MB
        private const int ChunkSize = 5 * 1024 * 1024; // 5MB        
        private long MaxTotalAttachmentSize { get; set; } = 35 * 1024 * 1024; // 35MB - protect against memory issues with huge attachments
        private readonly GraphMailOptionsAuth _optionsAuth;
        private readonly ClientSecretCredential _credential;
        private readonly ILogger<GraphMailSender>? _logger;
        private readonly HttpClient _httpClient;
        private readonly bool _ownsHttpClient;
        private static readonly string[] scopes = { "https://graph.microsoft.com/.default" };
        private static readonly FileExtensionContentTypeProvider _provider = new FileExtensionContentTypeProvider();
        private int _disposed;

        // Centralized Polly retry policy
        private readonly AsyncPolicy<HttpResponseMessage> _retryPolicy;

        /// <summary>
        /// Initializes a new instance of the GraphMailSender class using IHttpClientFactory (recommended).
        /// </summary>
        /// <param name="optionsAuth">Graph authentication options (required).</param>
        /// <param name="httpClientFactory">IHttpClientFactory for creating HttpClient instances.</param>
        /// <param name="logger">Optional logger instance.</param>
        /// <remarks>
        /// Using IHttpClientFactory is recommended to avoid socket exhaustion issues.
        /// The factory manages the HttpClient lifecycle automatically.
        /// </remarks>
        public GraphMailSender(
            GraphMailOptionsAuth optionsAuth,
            IHttpClientFactory httpClientFactory,
            GraphMailOptions? options=null,
            ILogger<GraphMailSender>? logger = null)
        {
            _optionsAuth = optionsAuth ?? throw new ArgumentNullException(nameof(optionsAuth));
            _credential = new ClientSecretCredential(_optionsAuth.TenantId, _optionsAuth.ClientId, _optionsAuth.ClientSecret);
            _logger = logger;
            _httpClient = (httpClientFactory ?? throw new ArgumentNullException(nameof(httpClientFactory))).CreateClient();
            if (options?.HttpClientTimeout != null && options.HttpClientTimeout > TimeSpan.Zero)
            {
                _httpClient.Timeout = options.HttpClientTimeout.Value;
            }
            MaxTotalAttachmentSize = options?.MaxTotalAttachmentSize ?? MaxTotalAttachmentSize;
            _ownsHttpClient = false; // Never own HttpClient from factory
            _retryPolicy = CreateRetryPolicy();
        }

        /// <summary>
        /// Initializes a new instance of the GraphMailSender class using a direct HttpClient instance.
        /// </summary>
        /// <param name="optionsAuth">Graph authentication options (required).</param>
        /// <param name="httpClient">HttpClient instance to use.</param>
        /// <param name="logger">Optional logger instance.</param>
        /// <remarks>
        /// Note: The provided HttpClient will not be disposed by this class.
        /// For production use, prefer the constructor with IHttpClientFactory.
        /// </remarks>
        public GraphMailSender(
            GraphMailOptionsAuth optionsAuth,
            HttpClient httpClient,
            GraphMailOptions? options = null,
            ILogger<GraphMailSender>? logger = null)
        {
            _optionsAuth = optionsAuth ?? throw new ArgumentNullException(nameof(optionsAuth));
            _credential = new ClientSecretCredential(_optionsAuth.TenantId, _optionsAuth.ClientId, _optionsAuth.ClientSecret);
            _logger = logger;
            _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            if (options?.HttpClientTimeout != null && options.HttpClientTimeout > TimeSpan.Zero)
            {
                _httpClient.Timeout = options.HttpClientTimeout.Value;
            }
            MaxTotalAttachmentSize = options?.MaxTotalAttachmentSize ?? MaxTotalAttachmentSize;
            _ownsHttpClient = false; // Don't own injected HttpClient
            _retryPolicy = CreateRetryPolicy();
        }

        /// <summary>
        /// Initializes a new instance of the GraphMailSender class.
        /// Creates a new HttpClient instance internally (not recommended for production).
        /// </summary>
        /// <param name="optionsAuth">Graph authentication options (required).</param>
        /// <param name="logger">Optional logger instance.</param>
        /// <remarks>
        /// This constructor creates a new HttpClient instance which can lead to socket exhaustion.
        /// For production use, prefer the constructor with IHttpClientFactory.
        /// </remarks>
        public GraphMailSender(
            GraphMailOptionsAuth optionsAuth,
            GraphMailOptions? options = null,
            ILogger<GraphMailSender>? logger = null)
        {
            _optionsAuth = optionsAuth ?? throw new ArgumentNullException(nameof(optionsAuth));
            _credential = new ClientSecretCredential(_optionsAuth.TenantId, _optionsAuth.ClientId, _optionsAuth.ClientSecret);
            _logger = logger;
            _httpClient = new HttpClient();
            if (options?.HttpClientTimeout != null && options.HttpClientTimeout > TimeSpan.Zero)
            {
                _httpClient.Timeout = options.HttpClientTimeout.Value;
            }
            MaxTotalAttachmentSize = options?.MaxTotalAttachmentSize ?? MaxTotalAttachmentSize;
            _ownsHttpClient = true; // We own this one
            _retryPolicy = CreateRetryPolicy();
        }

        /// <summary> 
        /// Get access token using client credentials flow with proper caching and expiration handling
        /// </summary>
        private async Task<string> GetAccessTokenAsync(CancellationToken ct)
        {
            try
            {
                var token = await _credential.GetTokenAsync(
                    new TokenRequestContext(scopes), ct).ConfigureAwait(false);
                return token.Token;
            }
            catch (AuthenticationFailedException ex)
            {
                _logger?.LogAuthenticationFailed(ex);
                throw; // Rethrow - don't retry
            }
        }

        private Task<HttpResponseMessage> SendWithRetryAsync(
                        Func<HttpRequestMessage> requestFactory,
                        CancellationToken ct)
        {
            return _retryPolicy.ExecuteAsync(async () =>
            {
                using (var request = requestFactory())
                {
                    return await _httpClient
                        .SendAsync(request, ct)
                        .ConfigureAwait(false);
                }
            });
        }

        private Task<HttpResponseMessage> SendWithRetryAsync(
                        Func<Task<HttpRequestMessage>> requestFactory, 
                        CancellationToken ct)
        {
            return _retryPolicy.ExecuteAsync(async () =>
            {
                using (var request = await requestFactory().ConfigureAwait(false))  // Await the factory
                {
                    return await _httpClient
                        .SendAsync(request, ct)
                        .ConfigureAwait(false);
                }
            });
        }

        private AsyncPolicy<HttpResponseMessage> CreateRetryPolicy()
        {
            // Pre-generate jittered delays (Polly recommends doing this once)
            var jitterDelays = Backoff.DecorrelatedJitterBackoffV2( medianFirstRetryDelay: TimeSpan.FromSeconds(1),
                                                                    retryCount: 5).ToArray();

            return Policy<HttpResponseMessage>
                .Handle<HttpRequestException>()
                .Or<TaskCanceledException>()
                .OrResult(r =>
                    r.StatusCode == HttpStatusCode.RequestTimeout ||
                    r.StatusCode == (HttpStatusCode)429 ||
                    (int)r.StatusCode >= 500)
                .WaitAndRetryAsync(
                    retryCount: 5,
                    sleepDurationProvider: (retryAttempt, outcome, context) =>
                    {
                        // 1️. Honor Retry-After header if present
                        var response = outcome.Result;
                        if (response?.Headers?.RetryAfter?.Delta != null)
                        {
                            return response.Headers.RetryAfter.Delta.Value;
                        }

                        // 2️. Otherwise use jittered backoff
                        return jitterDelays[retryAttempt - 1];
                    },
                    onRetryAsync: async (outcome, delay, retryAttempt, context) =>
                    {
                        Exception? contentReadException = null;
                        string? body = null;
                        try
                        {
                            body = outcome.Result?.Content != null
                                ? await outcome.Result.Content.ReadAsStringAsync().ConfigureAwait(false)
                                : null;
                        }
                        catch (Exception ex)
                        {
                            contentReadException = ex;
                        }

                        _logger?.LogRetrying(
                            retryAttempt,
                            delay.TotalSeconds,
                            outcome.Result?.StatusCode ?? 0,
                            body ?? "No response body",
                            contentReadException);
                    });
        }


        /// <summary>
        /// Sends an email using Microsoft Graph API with support for large attachments.
        /// </summary>
        /// <param name="toRecipients">
        /// Required. A list of recipient email addresses to send the message to.
        /// Must contain at least one valid recipient.
        /// </param>
        /// <param name="ccRecipients">
        /// Optional. A list of CC (carbon copy) recipient email addresses.
        /// </param>
        /// <param name="bccRecipients">
        /// Optional. A list of BCC (blind carbon copy) recipient email addresses.
        /// </param>
        /// <param name="subject">
        /// Required. The subject line of the email. Subject is sanitized before sending.
        /// </param>
        /// <param name="body">
        /// Required. The body content of the email. Content is sanitized before sending.
        /// </param>
        /// <param name="isHtml">
        /// Optional. Indicates whether the body content is HTML (true) or plain text (false).
        /// Default is true.
        /// </param>
        /// <param name="attachments">
        /// Optional. A list of file attachments to include with the email.
        /// Files larger than 3 MB are uploaded in chunks using an upload session.
        /// Smaller files are uploaded directly as base64 content.
        /// </param>
        /// <param name="fromEmail">
        /// Optional. The sender's email address. If not provided, defaults to the mailbox
        /// address configured in <see cref="GraphMailOptionsAuth"/>.
        /// </param>
        /// <param name="ct">
        /// Optional. A <see cref="CancellationToken"/> to cancel the operation.
        /// Cancellation is checked during long-running operations such as large file uploads.
        /// </param>
        /// <remarks>
        /// <para>
        /// The method performs the following steps:
        /// 1. Creates a draft message in the sender's mailbox.
        /// 2. Attaches files (small attachments directly, large attachments via chunked upload).
        /// 3. Retrieves the complete message with attachments.
        /// 4. Sends the message using the <c>sendMail</c> endpoint with <c>saveToSentItems = false</c>.
        /// 5. Deletes the draft message after sending.
        /// </para>
        /// <para>
        /// Access tokens are acquired using <see cref="ClientSecretCredential"/> before each major Graph call.
        /// This ensures tokens are refreshed if they expire during long-running operations.
        /// </para>
        /// </remarks>
        /// <exception cref="ArgumentException">
        /// Thrown if <paramref name="toRecipients"/> is null or empty.
        /// </exception>
        /// <exception cref="FileNotFoundException">
        /// Thrown if an attachment file path does not exist.
        /// </exception>
        /// <exception cref="GraphMailAttachmentException">
        /// Thrown if an attachment is empty, fails to upload, or upload is incomplete.
        /// </exception>
        /// <exception cref="GraphMailFailedCreateMessageException">
        /// Thrown if the draft message cannot be created.
        /// </exception>
        /// <exception cref="GraphMailFailedSendMessageException">
        /// Thrown if the message cannot be sent.
        /// </exception>
        /// <exception cref="GraphMailFailedDeleteDraftMessageException">
        /// Thrown if the draft message cannot be deleted after sending.
        /// </exception>
        /// <exception cref="AggregateException">
        /// Thrown if multiple errors occur (e.g., send failure and draft cleanup failure).
        /// </exception>
        /// <exception cref="OperationCanceledException">
        /// Thrown if the operation is canceled via <paramref name="ct"/>.
        /// </exception>
        /// <example>
        /// Example usage:
        /// <code>
        /// var sender = new GraphMailSender(optionsAuth, httpClientFactory, logger);
        /// await sender.SendEmailAsync(
        ///     toRecipients: new List<string> { "user@example.com" },
        ///     ccRecipients: null,
        ///     bccRecipients: null,
        ///     subject: "Quarterly Report",
        ///     body: "<p>Please find the report attached.</p>",
        ///     isHtml: true,
        ///     attachments: new List<EmailAttachment>
        ///     {
        ///         new EmailAttachment { FileName = "report.pdf", FilePath = "C:\\Reports\\Q1.pdf" }
        ///     });
        /// </code>
        /// </example>
        [SuppressMessage("Usage", "CA2219:Do not raise exceptions in finally clauses",
                        Justification = "Exception is stored and thrown after finally block completes")]
        public async Task SendEmailAsync(
            List<string> toRecipients,
            List<string>? ccRecipients,
            List<string>? bccRecipients,
            string subject,
            string body,
            bool isHtml = true,
            List<EmailAttachment>? attachments = null,
            string? fromEmail = null,
            string? correlationId = null,
            CancellationToken ct = default)
        {
            bool draftCreated = false;
            string messageId = string.Empty;
            Exception? originalException = null;
            string userEncoded = string.Empty;
            var sw = Stopwatch.StartNew();
            IDisposable? scope = null;

            if (Volatile.Read(ref _disposed) == 1)
                throw new ObjectDisposedException(nameof(GraphMailSender));

            var effectiveCorrelationId =
                correlationId
                ?? Activity.Current?.Id
                ?? Guid.NewGuid().ToString("N");

            if (_logger != null)
            {             
                scope = _logger.BeginScope(new Dictionary<string, object> { ["CorrelationId"] = effectiveCorrelationId });
            }

            try
            {
                try
                {
                    // ... validation code ...                   
                    if (toRecipients == null || !(toRecipients.Count > 0))
                        throw new ArgumentException("At least one recipient is required", nameof(toRecipients));

                    fromEmail ??= _optionsAuth.MailboxAddress;
                    if (!EmailValidator.IsValidEmail(fromEmail))
                              throw new ArgumentException("'From' email address is not valid", nameof(fromEmail));

                    userEncoded = Uri.EscapeDataString(fromEmail);

                    List<string> invalidRecipients = toRecipients.Where(email => !EmailValidator.IsValidEmail(email)).ToList();
                    if (invalidRecipients.Count > 0)
                        throw new ArgumentException(
                            $"Invalid email addresses: {string.Join(", ", invalidRecipients)}",
                            nameof(toRecipients));

                    if (ccRecipients != null && ccRecipients.Count > 0)
                    {
                        List<string> invalidCc = ccRecipients.Where(email => !EmailValidator.IsValidEmail(email)).ToList();
                        if (invalidCc.Count > 0)
                            throw new ArgumentException(
                                $"Invalid CC email addresses: {string.Join(", ", invalidCc)}",
                                nameof(ccRecipients));
                    }

                    if (bccRecipients != null && bccRecipients.Count > 0)
                    {
                        List<string> invalidBcc = bccRecipients.Where(email => !EmailValidator.IsValidEmail(email)).ToList();
                        if (invalidBcc.Count > 0)
                            throw new ArgumentException(
                                $"Invalid BCC email addresses: {string.Join(", ", invalidBcc)}",
                                nameof(bccRecipients));
                    }

                    _logger?.LogSendingEmail(fromEmail, toRecipients.Count);

                    body = EmailSanitizer.SanitizeBody(body);
                    subject = EmailSanitizer.SanitizeSubject(subject);

                    _logger?.LogExecutionStep("Step 1: Create draft message", sw.ElapsedMilliseconds);

                    var messageUrl = $"https://graph.microsoft.com/v1.0/users/{userEncoded}/messages";

                    var message = new Message
                    {
                        Subject = subject,
                        Body = new Body
                        {
                            ContentType = isHtml ? "HTML" : "Text",
                            Content = body
                        },
                        ToRecipients = toRecipients.Select(email => new Recipient
                        {
                            EmailAddress = new EmailAddress { Address = email }
                        }).ToList()
                    };

                    if (ccRecipients?.Count > 0)
                        message.CcRecipients = ccRecipients.Select(email => new Recipient
                        {
                            EmailAddress = new EmailAddress { Address = email }
                        }).ToList();

                    if (bccRecipients?.Count > 0)
                        message.BccRecipients = bccRecipients.Select(email => new Recipient
                        {
                            EmailAddress = new EmailAddress { Address = email }
                        }).ToList();

                    var messageJson = JsonConvert.SerializeObject(message, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore,
                        ContractResolver = new Newtonsoft.Json.Serialization.CamelCasePropertyNamesContractResolver()
                    });

                    // Step 1: Create draft message
                    var messageResponse = await SendWithRetryAsync(async () =>
                    {
                        var token = await GetAccessTokenAsync(ct).ConfigureAwait(false); 
                        var req = new HttpRequestMessage(HttpMethod.Post, messageUrl);
                        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                        req.Content = new StringContent(messageJson, Encoding.UTF8, "application/json");
                        return req;
                    }, ct).ConfigureAwait(false);

                    if (!messageResponse.IsSuccessStatusCode)
                    {
                        var error = await GetErrorDetailsAsync(messageResponse).ConfigureAwait(false);
                        _logger?.LogFailedToCreateMessage(error);
                        throw new GraphMailFailedCreateMessageException($"Failed to create message: {error}");
                    }

                    // If targeting .NET 5+ or .NET Core 3.1+
#if NET5_0_OR_GREATER
                    var messageResponseBody = await messageResponse.Content
                          .ReadAsStringAsync(ct).ConfigureAwait(false);
#else
                    var messageResponseBody = await messageResponse.Content
                        .ReadAsStringAsync().ConfigureAwait(false);
#endif

                    var createdMessage = JObject.Parse(messageResponseBody);
                    messageId = createdMessage["id"]?.ToString()
                        ?? throw new GraphMailFailedCreateMessageException($"Message ID not found in response: {messageResponseBody}");

                    _logger?.LogDraftCreated(messageId);
                    draftCreated = true;

                    // Step 2: Attach files
                    _logger?.LogExecutionStep("Step 2: Attach files", sw.ElapsedMilliseconds);
                    if (attachments?.Count > 0)
                    {
                        // Validate total attachment size to prevent memory issues
                        long totalSize = 0;
                        foreach (var attachment in attachments)
                        {
                            if (!File.Exists(attachment.FilePath))
                            {
                                throw new FileNotFoundException($"Attachment file not found: {attachment.FilePath}", attachment.FilePath);
                            }

                            var fileInfo = new FileInfo(attachment.FilePath);
                            if (fileInfo.Length == 0)
                            {
                                throw new GraphMailAttachmentException($"Attachment file is empty: {attachment.FileName}");
                            }

                            if (attachment.IsInline && string.IsNullOrWhiteSpace(attachment.ContentId))
                            {
                                throw new ArgumentException($"Inline attachment {attachment.FileName} must have a ContentId");
                            }

                            totalSize += fileInfo.Length;
                        }

                        if (totalSize > MaxTotalAttachmentSize)
                        {
                            throw new GraphMailAttachmentException(
                                $"Total attachment size ({totalSize / 1024 / 1024}MB) exceeds limit ({MaxTotalAttachmentSize / 1024 / 1024}MB). " +
                                $"This protects against memory issues when retrieving the message with attachments.");
                        }

                        foreach (var attachment in attachments)
                        {
                            var fileInfo = new FileInfo(attachment.FilePath);
                            var fileSize = fileInfo.Length;
                            var contentType = GetMimeType(attachment.FileName);

                            _logger?.LogAttachingFile(attachment.FileName, fileSize);
                            if (fileSize > LargeAttachmentThreshold) // > 3MB
                            {
                                await UploadLargeAttachmentStreamAsync(userEncoded, messageId, attachment.FileName, attachment.FilePath, 
                                                                       fileSize, contentType, attachment.IsInline, 
                                                                       attachment.ContentId ?? string.Empty, ct).ConfigureAwait(false);
                            }
                            else
                            {
                                await AddSmallAttachmentAsync(userEncoded, messageId, attachment.FileName, attachment.FilePath,
                                                              contentType, attachment.IsInline,
                                                              attachment.ContentId ?? string.Empty, ct).ConfigureAwait(false);
                            }
                        }
                    }

                    // Step 3: Get the complete message with attachments
                    _logger?.LogExecutionStep("Step 3: Get the complete message with attachments", sw.ElapsedMilliseconds);
                    var getMessageUrl = $"https://graph.microsoft.com/v1.0/users/{userEncoded}/messages/{messageId}?$expand=attachments";
                    var getResponse = await SendWithRetryAsync(async () =>
                    {
                        var token = await GetAccessTokenAsync(ct).ConfigureAwait(false);  
                        var req = new HttpRequestMessage(HttpMethod.Get, getMessageUrl);
                        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                        return req;
                    }, ct).ConfigureAwait(false);

                    if (!getResponse.IsSuccessStatusCode)
                    {
                        var error = await GetErrorDetailsAsync(getResponse).ConfigureAwait(false);
                        _logger?.LogFailedToGetMessage(messageId, error);
                        throw new GraphMailFailedCreateMessageException($"Failed to retrieve message with attachments: {error}");
                    }

                    // Stream the JSON response directly to avoid loading entire response into memory
                    // This is critical for large attachments with base64-encoded contentBytes
                    JObject completeMessage;
                    using (var stream = await getResponse.Content.ReadAsStreamAsync().ConfigureAwait(false))
                    using (var streamReader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: JsonStreamBufferSize))
                    using (var jsonReader = new JsonTextReader(streamReader))
                    {
                        var serializer = new JsonSerializer();
                        completeMessage = serializer.Deserialize<JObject>(jsonReader)
                            ?? throw new GraphMailFailedCreateMessageException("Failed to parse message response");
                    }

                    // Remove read-only fields
                    var cleanMessage = CleanMessageForSending(completeMessage);

                    // Step 4: Send using sendMail endpoint
                    _logger?.LogExecutionStep("Step 4: Send using sendMail endpoint", sw.ElapsedMilliseconds);
                    var sendUrl = $"https://graph.microsoft.com/v1.0/users/{userEncoded}/sendMail";
                    var sendPayload = new
                    {
                        message = cleanMessage,
                        saveToSentItems = false
                    };
                    var sendJson = JsonConvert.SerializeObject(sendPayload);

                    var sendResponse = await SendWithRetryAsync(async () =>
                    {
                        var token = await GetAccessTokenAsync(ct).ConfigureAwait(false);  // Fresh token
                        var req = new HttpRequestMessage(HttpMethod.Post, sendUrl);
                        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                        req.Content = new StringContent(sendJson, Encoding.UTF8, "application/json");
                        return req;
                    }, ct).ConfigureAwait(false);

                    if (!sendResponse.IsSuccessStatusCode)
                    {
                        var error = await GetErrorDetailsAsync(sendResponse).ConfigureAwait(false);
                        _logger?.LogFailedToSendMessage(error);
                        throw new GraphMailFailedSendMessageException($"Failed to send message: {error}");
                    }

                    _logger?.LogMessageSent(messageId);
                }
                catch (Exception ex)
                {
                    originalException = ex;
                    _logger?.LogFailedToSendMessage("", ex);

                    // Don't throw yet - we'll handle it in finally after cleanup attempt
                }
                finally
                {
                    // Step 5: Delete the draft message
                    _logger?.LogExecutionStep("Step 5: Delete the draft message if it was created", sw.ElapsedMilliseconds);
                    if (draftCreated && !string.IsNullOrEmpty(messageId))
                    {
                        try
                        {
                            var deleteUrl = $"https://graph.microsoft.com/v1.0/users/{userEncoded}/messages/{messageId}";
                            var deleteResponse = await SendWithRetryAsync(async () =>
                            {
                                var token = await GetAccessTokenAsync(ct).ConfigureAwait(false);  
                                var req = new HttpRequestMessage(HttpMethod.Delete, deleteUrl);
                                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                                return req;
                            }, ct).ConfigureAwait(false);

                            if (!deleteResponse.IsSuccessStatusCode)
                            {
                                var error = await GetErrorDetailsAsync(deleteResponse).ConfigureAwait(false);
                                _logger?.LogFailedToDeleteDraft(messageId, error);
                                var cleanupEx = new GraphMailFailedDeleteDraftMessageException($"Failed to delete draft message {messageId}, Error {error}");
                                if (originalException != null)
                                {
                                    originalException = new AggregateException("Email operation failed with multiple errors", originalException, cleanupEx);
                                }
                                else
                                {
                                    originalException = cleanupEx;
                                }
                            }
                        }
                        catch (Exception cleanupEx)
                        {
                            originalException = originalException != null
                                ? new AggregateException("Email operation failed with multiple errors", originalException, cleanupEx)
                                : cleanupEx;
                        }
                    }

                    if (originalException != null)
                        throw originalException;
                }
            }
            finally
            {
                scope?.Dispose();
            }
        }

        private static JObject CleanMessageForSending(JObject completeMessage)
        {
            var cleanMessage = new JObject();

            // Copy only the fields needed for sending
            if (completeMessage["subject"] != null)
                cleanMessage["subject"] = completeMessage["subject"];
            if (completeMessage["body"] != null)
                cleanMessage["body"] = completeMessage["body"];
            if (completeMessage["toRecipients"] != null)
                cleanMessage["toRecipients"] = completeMessage["toRecipients"];
            if (completeMessage["ccRecipients"] != null)
                cleanMessage["ccRecipients"] = completeMessage["ccRecipients"];
            if (completeMessage["bccRecipients"] != null)
                cleanMessage["bccRecipients"] = completeMessage["bccRecipients"];
            if (completeMessage["replyTo"] != null)
                cleanMessage["replyTo"] = completeMessage["replyTo"];
            if (completeMessage["from"] != null)
                cleanMessage["from"] = completeMessage["from"];
            if (completeMessage["importance"] != null)
                cleanMessage["importance"] = completeMessage["importance"];

            // Clean attachments - remove metadata fields
            if (completeMessage["attachments"] != null)
            {
                if (completeMessage["attachments"] is JArray attachmentsArray && attachmentsArray.Count > 0)
                {
                    var cleanAttachments = new JArray();
                    foreach (var item in attachmentsArray)
                    {
                        if (item is JObject att)  // Explicit cast with pattern matching
                        {
                            var cleanAtt = new JObject();
                            if (att["@odata.type"] != null)
                                cleanAtt["@odata.type"] = att["@odata.type"];
                            if (att["name"] != null)
                                cleanAtt["name"] = att["name"];
                            if (att["contentType"] != null)
                                cleanAtt["contentType"] = att["contentType"];
                            if (att["contentBytes"] != null)
                                cleanAtt["contentBytes"] = att["contentBytes"];
                            if (att["size"] != null)
                                cleanAtt["size"] = att["size"];
                            if (att["isInline"] != null)
                                cleanAtt["isInline"] = att["isInline"];
                            if (att["contentId"] != null)
                                cleanAtt["contentId"] = att["contentId"];

                            cleanAttachments.Add(cleanAtt);
                        }
                    }
                    cleanMessage["attachments"] = cleanAttachments;
                }
            }

            return cleanMessage;
        }

        private async Task AddSmallAttachmentAsync(string fromEmail, string messageId,
            string fileName, string filePath, string contentType, bool isInline, string? contentId, CancellationToken ct)
        {
            fileName = fileName.SanitizeFilename();
            var attachUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages/{messageId}/attachments";

            var base64Content = await filePath.StreamFileAsBase64Async(ct).ConfigureAwait(false);
            
            var attachment = new
            {
                odataType = "#microsoft.graph.fileAttachment",
                name = fileName,
                contentType,
                contentBytes = base64Content,
                isInline,
                contentId  // Will be null if not inline
            };

            var json = JsonConvert.SerializeObject(attachment, new JsonSerializerSettings
            {
                ContractResolver = new ODataContractResolver()
            });

            var response = await SendWithRetryAsync(async () =>  
            {
                var token = await GetAccessTokenAsync(ct).ConfigureAwait(false);
                var req = new HttpRequestMessage(HttpMethod.Post, attachUrl);
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                req.Content = new StringContent(json, Encoding.UTF8, "application/json");
                return req;
            }, ct).ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
            {
                var error = await GetErrorDetailsAsync(response).ConfigureAwait(false);
                _logger?.LogFailedToAddAttachment(fileName, error);
                throw new GraphMailAttachmentException($"Failed to add small attachment '{fileName}': {error}");
            }

            _logger?.LogSmallAttachmentAdded(fileName);
        }


        private async Task UploadLargeAttachmentStreamAsync(
            string fromEmail, string messageId, string fileName,
            string filePath, long fileSize, string contentType,
            bool isInline, string? contentId, CancellationToken ct)
        {
            fileName = fileName.SanitizeFilename();
            const int maxSessionRetries = 3;
            var jitterDelays = Backoff.DecorrelatedJitterBackoffV2(medianFirstRetryDelay: TimeSpan.FromSeconds(1),
                                                        retryCount: maxSessionRetries).ToArray();

            for (int sessionAttempt = 0; sessionAttempt < maxSessionRetries; sessionAttempt++)
            {
                try
                {
                    var sessionUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages/{messageId}/attachments/createUploadSession";
                    var sessionData = new
                    {
                        AttachmentItem = new
                        {
                            attachmentType = "file",
                            name = fileName,
                            size = fileSize,
                            isInline,  
                            contentId // an be null if not inline
                        }
                    };
                    var sessionJson = JsonConvert.SerializeObject(sessionData);

                    var sessionResponse = await SendWithRetryAsync(async () =>
                    {
                        var token = await GetAccessTokenAsync(ct).ConfigureAwait(false);
                        var req = new HttpRequestMessage(HttpMethod.Post, sessionUrl);
                        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                        req.Content = new StringContent(sessionJson, Encoding.UTF8, "application/json");
                        return req;
                    }, ct).ConfigureAwait(false);

                    if (!sessionResponse.IsSuccessStatusCode)
                    {
                        var error = await GetErrorDetailsAsync(sessionResponse).ConfigureAwait(false);
                        _logger?.LogFailedToCreateUploadSession(fileName, error);
                        throw new GraphMailAttachmentException($"Failed to create upload session for '{fileName}': {error}");
                    }

                    var sessionBody = await sessionResponse.Content.ReadAsStringAsync().ConfigureAwait(false);
                    var uploadUrl = JObject.Parse(sessionBody)["uploadUrl"]?.ToString()
                        ?? throw new GraphMailAttachmentException($"uploadUrl not found for '{fileName}'");

                    _logger?.LogUploadSessionUrl(uploadUrl.StripAfter('?'), fileName, sessionAttempt + 1, maxSessionRetries, messageId);                    

                    var buffer = ArrayPool<byte>.Shared.Rent(ChunkSize);
                    long offset = 0;

                    try
                    {
                        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, FileStreamBufferSize, true);
                        while (offset < fileSize)
                        {
                            ct.ThrowIfCancellationRequested();


                            int bytesToRead = Math.Min(buffer.Length, (int)(fileSize - offset));
                            int totalBytesRead = 0;

                            while (totalBytesRead < bytesToRead)
                            {
                                int bytesRead = await fs.ReadAsync(
                                    buffer, totalBytesRead, bytesToRead - totalBytesRead, ct);

                                if (bytesRead == 0)
                                    throw new GraphMailAttachmentException(
                                    $"Unexpected end of file while uploading '{fileName}'. " +
                                    $"Expected to read from offset {offset} but file stream returned {bytesRead} bytes. " +
                                    $"File size: {fileSize}, bytes uploaded: {offset}");

                                totalBytesRead += bytesRead;
                            }

                            long end = offset + totalBytesRead - 1;

                            // Note: uploadUrl from Graph API is pre-authenticated, so we don't need to set Authorization header
                            var response = await SendWithRetryAsync(() =>
                            {
                                var req = new HttpRequestMessage(HttpMethod.Put, uploadUrl);
                                var content = new ByteArrayContent(buffer, 0, totalBytesRead);
                                req.Content = content;
                                req.Content.Headers.ContentType = new MediaTypeHeaderValue(contentType);
                                req.Content.Headers.ContentLength = totalBytesRead;
                                req.Content.Headers.ContentRange = new ContentRangeHeaderValue(offset, end, fileSize);
                                return req;
                            }, ct).ConfigureAwait(false);

                            // Check for 404 - session expired or not found (Graph API backend issue)
                            if (response.StatusCode == HttpStatusCode.NotFound)
                            {
                                var errorBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                                _logger?.LogChunkFailed((int)response.StatusCode,
                                    $"Chunk upload failed for '{fileName}' at offset {offset}: " +
                                    $"SESSION_INVALID_404: {response.ReasonPhrase ?? ""}", errorBody);

                                // Throw special exception to trigger session retry
                                throw new GraphMailAttachmentException("SESSION_INVALID_404",
                                    new HttpRequestException("Upload session not found (404) during chunk upload – likely backend issue"), fileName);
                            }

                            if (!response.IsSuccessStatusCode)
                            {
                                var errorBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                                _logger?.LogChunkFailed((int)response.StatusCode, response.ReasonPhrase ?? "", errorBody);
                                throw new GraphMailAttachmentException(
                                    $"Chunk upload failed for '{fileName}' at offset {offset}: " +
                                    $"{(int)response.StatusCode} {response.ReasonPhrase} - {errorBody}");
                            }

                            var responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                            offset = end + 1;
                            _logger?.LogChunkStatus(offset, fileSize, fileName, (int)response.StatusCode);

                            // Check nextExpectedRanges to see if we can break early
                            if (!string.IsNullOrWhiteSpace(responseBody))
                            {
                                var responseJson = JObject.Parse(responseBody);
                                if (responseJson["nextExpectedRanges"] is JArray nextRanges && nextRanges.Count > 0)
                                {
                                    _logger?.LogResponseBodyTrace($"Next expected ranges: {string.Join(", ", nextRanges)}");
                                    continue;
                                }
                                // No nextExpectedRanges means upload is complete
                                break;
                            }
                        }

                        if (offset != fileSize)
                        {
                            throw new GraphMailAttachmentException(
                                $"Incomplete file upload for '{fileName}'. " +
                                $"Expected to upload {fileSize} bytes but only uploaded {offset} bytes.");
                        }

                        _logger?.LogUploadComplete(fileName);

                        // Success - return from method
                        return;
                    }
                    catch (OperationCanceledException ex)
                    {
                        _logger?.LogUploadCancelled(fileName, offset, fileSize, ex);
                        throw;
                    }
                    catch (IOException ex)
                    {
                        throw new GraphMailAttachmentException(
                            $"IO error while reading file '{fileName}' at offset {offset}: {ex.Message}", ex);
                    }
                    finally
                    {
                        ArrayPool<byte>.Shared.Return(buffer);
                    }
                }
                catch (GraphMailAttachmentException ex) when (
                    ex.Message.Contains("SESSION_INVALID_404") ||
                    ex.Message.Contains("ErrorItemNotFound"))
                {
                    // This is a known Graph API backend issue - intermittent failures
                    // where upload sessions become invalid immediately or aren't properly
                    // initialized in Exchange backend

                    if (sessionAttempt == maxSessionRetries - 1)
                    {
                        // Last attempt failed - throw with context
                        throw new GraphMailAttachmentException(
                            $"Failed to upload '{fileName}' after {maxSessionRetries} session attempts. " +
                            $"This appears to be a Graph API backend issue (ErrorItemNotFound). " +
                            $"Draft message: {messageId}", ex);
                    }

                    // Calculate  backoff 
                    var delaySeconds = jitterDelays[sessionAttempt];                   
                    _logger?.LogSessionExpired(fileName, sessionAttempt+1, maxSessionRetries, delaySeconds.TotalSeconds, ex);

                    await Task.Delay(delaySeconds, ct).ConfigureAwait(false);

                    // Loop continues - will create new session and retry
                }
            }

            // Should never reach here due to return or throw in loop
            throw new GraphMailAttachmentException(
                $"Failed to upload '{fileName}' - unexpected exit from retry loop");
        }

        private static string GetMimeType(string fileName)
        {
            return _provider.TryGetContentType(fileName, out var contentType)
                ? contentType
                : "application/octet-stream";
        }

        private static async Task<string> GetErrorDetailsAsync(HttpResponseMessage? response)
        {
            try
            {
                if (response == null)
                    return "Error: Response is null";

                if (response.Content == null)
                    return $"Status: {response.StatusCode}, Body: (no content)";

                var errorBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                // Try to parse and format Graph API error for better readability
                if (!string.IsNullOrWhiteSpace(errorBody) && errorBody.TrimStart().StartsWith("{", StringComparison.Ordinal))
                {
                    try
                    {
                        var errorJson = JObject.Parse(errorBody);
                        var errorCode = errorJson["error"]?["code"]?.ToString();
                        var errorMessage = errorJson["error"]?["message"]?.ToString();

                        if (!string.IsNullOrEmpty(errorCode))
                        {
                            return $"Status: {response.StatusCode}, Code: {errorCode}, Message: {errorMessage ?? "N/A"}";
                        }
                    }
                    catch
                    {
                        // Fall through to return raw body
                    }
                }

                return $"Status: {response.StatusCode}, Body: {errorBody}";
            }
            catch (Exception ex)
            {
                return $"Status: {response?.StatusCode ?? 0}, Error reading response: {ex.Message}";
            }
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) == 1)
                return;

            // Only dispose HttpClient if we own it (created it, not injected)
            if (_ownsHttpClient)
            {
                _httpClient?.Dispose();
            }
           
            GC.SuppressFinalize(this);
        }
    }
}

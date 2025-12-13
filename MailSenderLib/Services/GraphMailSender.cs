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
using Microsoft.Extensions.Http;
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
using System.Net.Mime;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MailSenderLib.Services
{
    public class GraphMailSender : IDisposable, IGraphMailSender
    {
        private const long LargeAttachmentThreshold = 3 * 1024 * 1024; // 3MB
        private const int ChunkSize = 5 * 1024 * 1024; // 5MB
        private readonly GraphMailOptionsAuth _optionsAuth;
        private readonly ClientSecretCredential _credential;
        private readonly ILogger<GraphMailSender>? _logger;
        private readonly HttpClient _httpClient;
        private readonly bool _ownsHttpClient;
        private static readonly string[] scopes = { "https://graph.microsoft.com/.default" };
        private static readonly FileExtensionContentTypeProvider _provider = new FileExtensionContentTypeProvider();
        private const string HttpClientName = "GraphMailSender";

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
            ILogger<GraphMailSender>? logger = null)
        {
            _optionsAuth = optionsAuth ?? throw new ArgumentNullException(nameof(optionsAuth));
            _credential = new ClientSecretCredential(_optionsAuth.TenantId, _optionsAuth.ClientId, _optionsAuth.ClientSecret);
            _logger = logger;
            _httpClient = (httpClientFactory ?? throw new ArgumentNullException(nameof(httpClientFactory))).CreateClient(HttpClientName);
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
            ILogger<GraphMailSender>? logger = null)
        {
            _optionsAuth = optionsAuth ?? throw new ArgumentNullException(nameof(optionsAuth));
            _credential = new ClientSecretCredential(_optionsAuth.TenantId, _optionsAuth.ClientId, _optionsAuth.ClientSecret);
            _logger = logger;
            _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
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
            ILogger<GraphMailSender>? logger = null)
        {
            _optionsAuth = optionsAuth ?? throw new ArgumentNullException(nameof(optionsAuth));
            _credential = new ClientSecretCredential(_optionsAuth.TenantId, _optionsAuth.ClientId, _optionsAuth.ClientSecret);
            _logger = logger;
            _httpClient = new HttpClient();
            _ownsHttpClient = true; // We own this one
            _retryPolicy = CreateRetryPolicy();
        }

        /// <summary> 
        /// Get access token using client credentials flow with proper caching and expiration handling
        /// </summary>
        private async Task<string> GetAccessTokenAsync(CancellationToken ct)
        {
            // Simplified: rely on ClientSecretCredential built-in caching/refresh
            var token = await _credential.GetTokenAsync(new TokenRequestContext(scopes), ct).ConfigureAwait(false);
            _logger?.LogTokenAcquired(token.ExpiresOn);
            return token.Token;
        }

        private AsyncPolicy<HttpResponseMessage> CreateRetryPolicy()
        {            
            // Exponential backoff with jitter (decorrelated)
            var delays = Backoff.DecorrelatedJitterBackoffV2(
                medianFirstRetryDelay: TimeSpan.FromSeconds(1),
                retryCount: 5);

            return Policy<HttpResponseMessage>
                .Handle<HttpRequestException>() // network errors
                .Or<TaskCanceledException>()     // timeouts in .NET Standard 2.0 can surface as TaskCanceledException
                .OrResult(r =>
                    r.StatusCode == HttpStatusCode.RequestTimeout ||           // 408
                    r.StatusCode == (HttpStatusCode)429 ||                     //HttpStatusCode.TooManyRequests 
                    (int)r.StatusCode >= 500)                                  // 5xx
                .WaitAndRetryAsync(
                    delays,
                    async (outcome, timespan, retryAttempt, context) =>
                    {
                        // Honor Retry-After header if present
                        var result = outcome.Result;
                        if (result != null && result.Headers?.RetryAfter?.Delta != null)
                        {
                            timespan = result.Headers.RetryAfter.Delta.Value;
                        }
                        
                        Exception? contentReadException = null;
                        string? reason = null;
                        try
                        {
                            contentReadException = null;
                            if (result != null && result.Content != null)
                            {
                                reason = await result.Content.ReadAsStringAsync().ConfigureAwait(false);
                                if (!string.IsNullOrEmpty(reason) && reason.Length > 500) reason = reason.Substring(0, 500) + "...";
                            }
                        }
                        catch(Exception ex)
                        {
                            contentReadException = ex;
                        }

                        _logger?.LogRetrying(retryAttempt, timespan, result?.StatusCode ?? 0, reason ?? "Unknown", contentReadException);
                        
                    });
        }

        private Task<HttpResponseMessage> SendWithRetryAsync(Func<Task<HttpResponseMessage>> sendFunc)
        {
            return _retryPolicy.ExecuteAsync(sendFunc);
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
            CancellationToken ct = default)
        {
            bool draftCreated = false;
            string messageId = string.Empty;
            Exception? originalException = null;
            string token = string.Empty;
            string userEncoded = string.Empty;
            var sw = Stopwatch.StartNew();
            
            try
            {
                if (toRecipients == null || !(toRecipients.Count > 0))
                    throw new ArgumentException("At least one recipient is required", nameof(toRecipients));

                fromEmail ??= _optionsAuth.MailboxAddress;
                userEncoded = Uri.EscapeDataString(fromEmail);

                _logger?.LogSendingEmail(fromEmail, toRecipients.Count);

                body = EmailSanitizer.SanitizeBody(body);
                subject = EmailSanitizer.SanitizeSubject(subject);

                // Fetch token on demand
                token = await GetAccessTokenAsync(ct).ConfigureAwait(false);
                
                _logger?.LogExecutionStep("Step 1: Create draft message", sw.ElapsedMilliseconds);

                var messageUrl = $"https://graph.microsoft.com/v1.0/users/{userEncoded}/messages";

                var message = new MessagePayload
                {
                    Subject = subject,
                    Body = new BodyPayload
                    {
                        ContentType = isHtml ? "HTML" : "Text",
                        Content = body
                    },
                    ToRecipients = toRecipients.Select(email => new RecipientPayload
                    {
                        EmailAddress = new EmailAddressPayload { Address = email }
                    }).ToList()
                };

                if (ccRecipients != null && ccRecipients.Count > 0)
                {
                    message.CcRecipients = ccRecipients.Select(email => new RecipientPayload
                    {
                        EmailAddress = new EmailAddressPayload { Address = email }
                    }).ToList();
                }

                if (bccRecipients != null && bccRecipients.Count > 0)
                {
                    message.BccRecipients = bccRecipients.Select(email => new RecipientPayload
                    {
                        EmailAddress = new EmailAddressPayload { Address = email }
                    }).ToList();
                }

                var messageJson = JsonConvert.SerializeObject(message, new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Ignore,
                    ContractResolver = new Newtonsoft.Json.Serialization.CamelCasePropertyNamesContractResolver()
                });

                var messageResponse = await SendWithRetryAsync(() =>
                    _httpClient.SendJsonWithTokenAsync(HttpMethod.Post, messageUrl, messageJson, token, ct)).ConfigureAwait(false);

                if (!messageResponse.IsSuccessStatusCode)
                {
                    var error = await GetErrorDetailsAsync(messageResponse).ConfigureAwait(false);
                    _logger?.LogFailedToCreateMessage(error);
                    throw new GraphMailFailedCreateMessageException($"Failed to create message: {error}");
                }

                var messageResponseBody = await messageResponse.Content.ReadAsStringAsync().ConfigureAwait(false);
                var createdMessage = JObject.Parse(messageResponseBody);
                messageId = createdMessage["id"]?.ToString()
                    ?? throw new GraphMailFailedCreateMessageException($"Message ID not found in response: {messageResponseBody}");

                _logger?.LogDraftCreated(messageId);
                draftCreated = true;

                _logger?.LogExecutionStep("Step 2: Attach files", sw.ElapsedMilliseconds);
                if (attachments != null && attachments.Count > 0)
                {
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
                        var fileSize = fileInfo.Length;
                        var contentType = GetMimeType(attachment.FileName);

                        _logger?.LogAttachingFile(attachment.FileName, fileSize);
                        if (fileSize > LargeAttachmentThreshold) // > 3MB
                        {
                            await UploadLargeAttachmentStreamAsync(userEncoded, messageId, attachment.FileName, attachment.FilePath, fileSize, contentType, ct).ConfigureAwait(false);
                        }
                        else
                        {
                            await AddSmallAttachmentAsync(userEncoded, messageId, attachment.FileName, attachment.FilePath, contentType, ct).ConfigureAwait(false);
                        }
                    }
                }
                
                _logger?.LogExecutionStep("Step 3: Get the complete message with attachments", sw.ElapsedMilliseconds);
                var getMessageUrl = $"https://graph.microsoft.com/v1.0/users/{userEncoded}/messages/{messageId}?$expand=attachments";
                // Fetch token on demand (again)
                token = await GetAccessTokenAsync(ct).ConfigureAwait(false);
                var getResponse = await SendWithRetryAsync(() =>
                    _httpClient.SendJsonWithTokenAsync(HttpMethod.Get, getMessageUrl, null, token, ct)).ConfigureAwait(false);
                getResponse.EnsureSuccessStatusCode();
                
                var completeMessageBody = await getResponse.Content.ReadAsStringAsync().ConfigureAwait(false);
                var completeMessage = JObject.Parse(completeMessageBody);

                // Remove metadata and read-only fields
                var cleanMessage = CleanMessageForSending(completeMessage);

                _logger?.LogExecutionStep("Step 4: Send using sendMail endpoint", sw.ElapsedMilliseconds);
                var sendUrl = $"https://graph.microsoft.com/v1.0/users/{userEncoded}/sendMail";

                var sendPayload = new
                {
                    message = cleanMessage,
                    saveToSentItems = false
                };
                var sendJson = JsonConvert.SerializeObject(sendPayload);

                token = await GetAccessTokenAsync(ct).ConfigureAwait(false);
                var sendResponse = await SendWithRetryAsync(() =>
                    _httpClient.SendJsonWithTokenAsync(HttpMethod.Post, sendUrl, sendJson, token, ct)).ConfigureAwait(false);

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
                _logger?.LogExecutionStep("Step 5: Delete the draft message if it was created", sw.ElapsedMilliseconds);
                if (draftCreated && !string.IsNullOrEmpty(messageId))
                {
                    try
                    {
                        var deleteUrl = $"https://graph.microsoft.com/v1.0/users/{userEncoded}/messages/{messageId}";
                        var tokenForDelete = await GetAccessTokenAsync(ct).ConfigureAwait(false);
                        var deleteResponse = await SendWithRetryAsync(() =>
                            _httpClient.SendJsonWithTokenAsync(HttpMethod.Delete, deleteUrl, null, tokenForDelete, ct)).ConfigureAwait(false);

                        if (!deleteResponse.IsSuccessStatusCode)
                        {
                            var error = await GetErrorDetailsAsync(deleteResponse).ConfigureAwait(false);
                            _logger?.LogFailedToDeleteDraft(messageId, error);

                            var cleanupEx = new GraphMailFailedDeleteDraftMessageException($"Failed to delete draft message {messageId}, Error {error}");

                            if (originalException != null)
                            {
                                // Combine both exceptions
                                originalException = new AggregateException(
                                    "Email operation failed with multiple errors",
                                    originalException,
                                    cleanupEx);
                            }
                            else
                            {
                                // This is the only error
                                originalException = cleanupEx;
                            }
                        }
                    }
                    catch (Exception cleanupEx)
                    {
                        if (originalException != null)
                        {
                            // Combine both exceptions
                            originalException = new AggregateException(
                                "Email operation failed with multiple errors",
                                originalException,
                                cleanupEx);
                        }
                        else
                        {
                            // This is the only error
                            originalException = cleanupEx;
                        }
                    }
                }

                // Now throw if there was any exception
                if (originalException != null)
                {
                    throw originalException;
                }
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
                            cleanAttachments.Add(cleanAtt);
                        }
                    }
                    cleanMessage["attachments"] = cleanAttachments;
                }
            }

            return cleanMessage;
        }

        private async Task AddSmallAttachmentAsync(string fromEmail, string messageId,
            string fileName, string filePath, string contentType, CancellationToken ct)
        {
            // Do NOT sanitize filePath anymore
            fileName = fileName.SanitizeFilename();

            var attachUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages/{messageId}/attachments";

            // Read file and encode to base64
            var fileBytes = await filePath.ReadAllBytesAsync(ct).ConfigureAwait(false);
            var base64Content = Convert.ToBase64String(fileBytes);

            var attachment = new
            {
                odataType = "#microsoft.graph.fileAttachment",
                name = fileName,
                contentType,
                contentBytes = base64Content
            };

            var json = JsonConvert.SerializeObject(attachment, new JsonSerializerSettings
            {
                ContractResolver = new ODataContractResolver()
            });

            var token = await GetAccessTokenAsync(ct).ConfigureAwait(false);
            var response = await SendWithRetryAsync(() =>
                _httpClient.SendJsonWithTokenAsync(HttpMethod.Post, attachUrl, json, token, ct)).ConfigureAwait(false);
            response.EnsureSuccessStatusCode();

            _logger?.LogSmallAttachmentAdded(fileName);
        }

        private async Task UploadLargeAttachmentStreamAsync(string fromEmail, string messageId,
            string fileName, string filePath, long fileSize, string contentType, CancellationToken ct)
        {
            // Do NOT sanitize filePath anymore
            fileName = fileName.SanitizeFilename();

            // Create upload session
            var sessionUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages/{messageId}/attachments/createUploadSession";

            var sessionData = new
            {
                AttachmentItem = new
                {
                    attachmentType = "file",
                    name = fileName,
                    size = fileSize
                }
            };

            var sessionJson = JsonConvert.SerializeObject(sessionData);

            var token = await GetAccessTokenAsync(ct).ConfigureAwait(false);
            var sessionResponse = await SendWithRetryAsync(() =>
                _httpClient.SendJsonWithTokenAsync(HttpMethod.Post, sessionUrl, sessionJson, token, ct)).ConfigureAwait(false);

            if (!sessionResponse.IsSuccessStatusCode)
            {
                var errorBody = await sessionResponse.Content.ReadAsStringAsync().ConfigureAwait(false);
                _logger?.LogChunkFailed((int)sessionResponse.StatusCode, sessionResponse.ReasonPhrase ?? "", errorBody);
                throw new GraphMailAttachmentException($"Failed to create upload session: {(int)sessionResponse.StatusCode} {sessionResponse.ReasonPhrase} - {errorBody}");
            }

            var sessionResponseBody = await sessionResponse.Content.ReadAsStringAsync().ConfigureAwait(false);
            var sessionInfo = JObject.Parse(sessionResponseBody);
            var uploadUrl = sessionInfo["uploadUrl"]?.ToString()
                ?? throw new GraphMailAttachmentException("uploadUrl not found in response");

            _logger?.LogUploadSessionUrl(uploadUrl, fileName);

            // Upload in chunks using streaming (5MB chunks)
            var buffer = ArrayPool<byte>.Shared.Rent(ChunkSize);
            long offset = 0;

            try
            {
                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 4096, useAsync: true))
                {
                    while (offset < fileSize)
                    {
                        // Check cancellation before each chunk
                        ct.ThrowIfCancellationRequested();

                        int bytesRead = await fileStream.ReadAsync(buffer, 0, Math.Min(buffer.Length, (int)(fileSize - offset)), ct).ConfigureAwait(false);

                        // FIXED: Validate that we read the expected number of bytes
                        if (bytesRead <= 0)
                        {
                            throw new GraphMailAttachmentException(
                                $"Unexpected end of file while uploading '{fileName}'. " +
                                $"Expected to read from offset {offset} but file stream returned {bytesRead} bytes. " +
                                $"File size: {fileSize}, bytes uploaded: {offset}");
                        }

                        long end = offset + bytesRead - 1;

                        using (var request = new HttpRequestMessage(HttpMethod.Put, uploadUrl))
                        using (var content = new ByteArrayContent(buffer, 0, bytesRead))
                        {
                            request.Content = content;
                            request.Content.Headers.ContentType = new MediaTypeHeaderValue(contentType);
                            request.Content.Headers.ContentLength = bytesRead;
                            request.Content.Headers.ContentRange = new ContentRangeHeaderValue(offset, end, fileSize);

                            // Note: uploadUrl from Graph API is pre-authenticated, so we don't need to set Authorization header
                            var response = await SendWithRetryAsync(() => _httpClient.SendAsync(request, ct)).ConfigureAwait(false);

                            if (!response.IsSuccessStatusCode)
                            {
                                var errorBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                                _logger?.LogChunkFailed((int)response.StatusCode, response.ReasonPhrase ?? "", errorBody);
                                throw new GraphMailAttachmentException(
                                    $"Chunk upload failed for '{fileName}' at offset {offset}: " +
                                    $"{(int)response.StatusCode} {response.ReasonPhrase} - {errorBody}");
                            }

                            var responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                            // Update offset
                            offset = end + 1;
                            _logger?.LogChunkStatus(offset, fileSize, fileName, (int)response.StatusCode);

                            // Check if there are more chunks to upload by looking at nextExpectedRanges
                            if (!string.IsNullOrWhiteSpace(responseBody))
                            {
                                var responseJson = JObject.Parse(responseBody);
                                if (responseJson["nextExpectedRanges"] is JArray nextRanges && nextRanges.Count > 0)
                                {
                                    // There are more chunks expected, continue uploading
                                    _logger?.LogResponseBodyTrace($"Next expected ranges: {string.Join(", ", nextRanges)}");
                                    continue;
                                }
                            }

                            // No nextExpectedRanges means upload is complete
                            break;
                        }
                    }

                    // FIXED: Final validation that we uploaded the complete file
                    if (offset != fileSize)
                    {
                        throw new GraphMailAttachmentException(
                            $"Incomplete file upload for '{fileName}'. " +
                            $"Expected to upload {fileSize} bytes but only uploaded {offset} bytes. " +
                            $"The file may have been truncated or modified during upload.");
                    }

                    _logger?.LogUploadComplete(fileName);
                }
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
        private static string GetMimeType(string fileName)
        {
            return _provider.TryGetContentType(fileName, out var contentType)
                ? contentType
                : "application/octet-stream";
        }

        private static async Task<string> GetErrorDetailsAsync(HttpResponseMessage response)
        {
            try
            {
                var errorBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                return $"Status: {response.StatusCode}, Body: {errorBody}";
            }
            catch
            {
                return $"Status: {response.StatusCode}";
            }
        }

        public void Dispose()
        {
            // Only dispose HttpClient if we own it (created it, not injected)
            if (_ownsHttpClient)
            {
                _httpClient?.Dispose();
            }
            (_credential as IDisposable)?.Dispose();
            GC.SuppressFinalize(this);
        }
    }
}

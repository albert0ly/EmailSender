using Azure.Core;
using Azure.Identity;
using MailSenderLib.Exceptions;
using MailSenderLib.Interfaces;
using MailSenderLib.Logging;
using MailSenderLib.Models;
using MailSenderLib.Options;
using MailSenderLib.Extensions;
using MailSenderLib.Utils;
using Microsoft.Extensions.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Codeuctivity;
using System;
using System.Buffers;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
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
        private const string HttpClientName = "GraphMailSender";


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

        /// <summary>
        /// Send email with large attachments without saving to Sent Items
        /// </summary>
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

            try
            {
                if (toRecipients == null || !(toRecipients.Count > 0))
                    throw new ArgumentException("At least one recipient is required", nameof(toRecipients));

                fromEmail ??= _optionsAuth.MailboxAddress;

                _logger?.LogSendingEmail(fromEmail, toRecipients.Count);

                body = EmailSanitizer.SanitizeBody(body);
                subject = EmailSanitizer.SanitizeSubject(subject);

                // Fetch token on demand
                token = await GetAccessTokenAsync(ct);

                // Step 1: Create draft message using Newtonsoft.Json serialization
                var messageUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages";

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

                var messageResponse = await _httpClient.SendJsonWithTokenAsync(HttpMethod.Post, messageUrl, messageJson, token, ct);

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

                // Step 2: Attach files (stream large files, direct upload small files)
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

                        _logger?.LogAttachingFile(attachment.FileName, fileSize);
                        if (fileSize > LargeAttachmentThreshold) // > 3MB
                        {
                            await UploadLargeAttachmentStreamAsync(fromEmail, messageId, attachment.FileName, attachment.FilePath, fileSize, ct).ConfigureAwait(false);
                        }
                        else
                        {
                            await AddSmallAttachmentAsync(fromEmail, messageId, attachment.FileName, attachment.FilePath, ct).ConfigureAwait(false);
                        }
                    }
                }

                // Step 3: Get the complete message with attachments
                var getMessageUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages/{messageId}?$expand=attachments";
                // Fetch token on demand (again)
                token = await GetAccessTokenAsync(ct);
                var getResponse = await _httpClient.SendJsonWithTokenAsync(HttpMethod.Get, getMessageUrl, null, token, ct);
                getResponse.EnsureSuccessStatusCode();

                var completeMessageBody = await getResponse.Content.ReadAsStringAsync().ConfigureAwait(false);
                var completeMessage = JObject.Parse(completeMessageBody);

                // Remove metadata and read-only fields
                var cleanMessage = CleanMessageForSending(completeMessage);

                // Step 4: Send using sendMail endpoint with SaveToSentItems = false
                var sendUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/sendMail";

                var sendPayload = new
                {
                    message = cleanMessage,
                    saveToSentItems = false
                };

                var sendJson = JsonConvert.SerializeObject(sendPayload);

                token = await GetAccessTokenAsync(ct);
                var sendResponse = await _httpClient.SendJsonWithTokenAsync(HttpMethod.Post, sendUrl, sendJson, token, ct);

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
                // Step 5: Delete the draft message if it was created
                if (draftCreated && !string.IsNullOrEmpty(messageId))
                {
                    try
                    {
                        var deleteUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages/{messageId}";
                        var tokenForDelete = await GetAccessTokenAsync(ct);
                        var deleteResponse = await _httpClient.SendJsonWithTokenAsync(HttpMethod.Delete, deleteUrl, null, tokenForDelete, ct);

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
            string fileName, string filePath, CancellationToken ct)
        {
            // Do NOT sanitize filePath anymore
            fileName = fileName.SanitizeFilename();

            var attachUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages/{messageId}/attachments";

            // Read file and encode to base64
            var fileBytes = await filePath.ReadAllBytesAsync();
            var base64Content = Convert.ToBase64String(fileBytes);

            var attachment = new
            {
                odataType = "#microsoft.graph.fileAttachment",
                name = fileName,
                contentBytes = base64Content
            };

            var json = JsonConvert.SerializeObject(attachment, new JsonSerializerSettings
            {
                ContractResolver = new ODataContractResolver()
            });

            var token = await GetAccessTokenAsync(ct);
            var response = await _httpClient.SendJsonWithTokenAsync(HttpMethod.Post, attachUrl, json, token, ct);
            response.EnsureSuccessStatusCode();

            _logger?.LogSmallAttachmentAdded(fileName);
        }

        private async Task UploadLargeAttachmentStreamAsync(string fromEmail, string messageId,
            string fileName, string filePath, long fileSize, CancellationToken ct)
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

            var token = await GetAccessTokenAsync(ct);
            var sessionResponse = await _httpClient.SendJsonWithTokenAsync(HttpMethod.Post, sessionUrl, sessionJson, token, ct);

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

                        var request = new HttpRequestMessage(HttpMethod.Put, uploadUrl)
                        {
                            Content = new ByteArrayContent(buffer, 0, bytesRead)
                        };

                        request.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                        request.Content.Headers.ContentLength = bytesRead;
                        request.Content.Headers.ContentRange = new ContentRangeHeaderValue(offset, end, fileSize);

                        // Note: uploadUrl from Graph API is pre-authenticated, so we don't need to set Authorization header
                        var response = await _httpClient.SendAsync(request, ct).ConfigureAwait(false);

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

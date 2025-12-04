using Azure.Core;
using Azure.Identity;
using MailSenderLib.Options;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MailSenderLib.Services
{
    public class GraphMailService : IDisposable
    {
        private readonly GraphMailOptionsAuth _optionsAuth;
        //private readonly string _tenantId;
        //private readonly string _clientId;
        //private readonly string _clientSecret;
        private readonly ClientSecretCredential _credential;
        private readonly ILogger<GraphMailService>? _logger;
        private readonly HttpClient _httpClient;
        private static readonly string[] scopes = { "https://graph.microsoft.com/.default" };

        // Cached token and lock for refresh
        private AccessToken _cachedToken;

        // Token caching with expiration
        //private string? _accessToken;
        //private DateTimeOffset _tokenExpiration;

        private readonly SemaphoreSlim _tokenLock = new SemaphoreSlim(1, 1);
        private static readonly TimeSpan TokenExpiryBuffer = TimeSpan.FromSeconds(30);

        // LoggerMessage delegates (avoid allocation-heavy LoggerExtensions calls)
        private static readonly Action<ILogger, Exception?> _logFailedToAcquireToken =
            LoggerMessage.Define(LogLevel.Error, new EventId(1000, nameof(_logFailedToAcquireToken)), "Failed to acquire access token for GraphMailSender");
        private static readonly Action<ILogger, Exception?> _logRefreshingToken =
            LoggerMessage.Define(LogLevel.Debug, new EventId(1001, nameof(_logRefreshingToken)), "Refreshing access token for GraphMailSender");
        private static readonly Action<ILogger, DateTimeOffset, Exception?> _logTokenAcquired =
            LoggerMessage.Define<DateTimeOffset>(LogLevel.Debug, new EventId(1002, nameof(_logTokenAcquired)), "Access token acquired, expires on {ExpiresOn}");
        private static readonly Action<ILogger, string,int, Exception?> _logSendingEmail =
                LoggerMessage.Define<string, int>(LogLevel.Debug, new EventId(1015, nameof(_logSendingEmail)), "Sending email from {From} to {ToCount} recipients");        
        private static readonly Action<ILogger, string, Exception?> _logFailedToCreateMessage =
          LoggerMessage.Define<string>(LogLevel.Error, new EventId(1016, nameof(_logFailedToCreateMessage)), "Failed to create message: {Error}");
        private static readonly Action<ILogger, string,  Exception?> _logDraftCreated =
          LoggerMessage.Define<string>(LogLevel.Debug, new EventId(1017, nameof(_logDraftCreated)), "Draft created {MessageId}");
        private static readonly Action<ILogger, string, long, Exception?> _logAttachingFile =
          LoggerMessage.Define<string, long>(LogLevel.Debug, new EventId(1018, nameof(_logAttachingFile)), "Attaching file {FileName} size {FileSize} recipients");        
        private static readonly Action<ILogger, string, Exception?> _logFailedToSendMessage =
          LoggerMessage.Define<string>(LogLevel.Error, new EventId(1019, nameof(_logFailedToSendMessage)), "Failed to send message: {Error}");        
        private static readonly Action<ILogger, string, string, Exception?> _logFailedToDeleteDraft =
          LoggerMessage.Define<string, string>(LogLevel.Error, new EventId(1020, nameof(_logFailedToDeleteDraft)), "Failed to delete draft message {MessageId}, Error {Error}");
        private static readonly Action<ILogger, string, Exception?> _logMessageSent =
          LoggerMessage.Define<string>(LogLevel.Debug, new EventId(1021, nameof(_logMessageSent)), "Message sent successfully without saving to Sent Items {MessageId}");
        private static readonly Action<ILogger, string, string, Exception?> _logUploadSessionUrl =
            LoggerMessage.Define<string, string>(LogLevel.Debug, new EventId(1013, nameof(_logUploadSessionUrl)), "Upload session URL: {Url} for file: {FileName}");
        private static readonly Action<ILogger, long, long, string, int, Exception?> _logChunkStatus =
            LoggerMessage.Define<long, long, string, int>(LogLevel.Debug, new EventId(1010, nameof(_logChunkStatus)), "Uploaded {Current}/{Total} bytes of {FileName}, Status {Status}");
        private static readonly Action<ILogger, string, Exception?> _logSmallAttachmentAdded =
            LoggerMessage.Define<string>(LogLevel.Debug, new EventId(1022, nameof(_logSmallAttachmentAdded)), "Small attachment added: {FileName}");
        private static readonly Action<ILogger, string, Exception?> _logUploadComplete =
            LoggerMessage.Define<string>(LogLevel.Debug, new EventId(1012, nameof(_logUploadComplete)), "Upload complete for {FileName}");
        private static readonly Action<ILogger, int, string, string, Exception?> _logChunkFailed =
            LoggerMessage.Define<int, string, string>(LogLevel.Error, new EventId(1014, nameof(_logChunkFailed)), "Chunk upload failed {Status} {Reason} - {Body}");


        private static readonly Action<ILogger, string, Exception?> _logResponseBodyTrace =
            LoggerMessage.Define<string>(LogLevel.Trace, new EventId(1011, nameof(_logResponseBodyTrace)), "{Body}");

        public GraphMailService(
            GraphMailOptionsAuth optionsAuth,
            ILogger<GraphMailService>? logger = null)
        {
            _optionsAuth = optionsAuth ?? throw new ArgumentNullException(nameof(optionsAuth));

            _credential = new ClientSecretCredential(_optionsAuth.TenantId, _optionsAuth.ClientId, _optionsAuth.ClientSecret);
            _logger = logger;
            _httpClient = new HttpClient();
        }

        /// <summary>
        /// Get access token using client credentials flow with proper caching and expiration handling
        /// </summary>
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

        /// <summary>
        /// Send email with large attachments without saving to Sent Items
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2201:Do not raise reserved exception types", Justification = "<Pending>")]
        public async Task SendMailWithLargeAttachmentsAsync(            
            List<string> toEmails,
            string subject,
            string body,
            List<EmailAttachment> attachments,
            string? fromEmail = null,
            List<string>? ccEmails = null,
            List<string>? bccEmails = null,
            bool isHtml = true,
            CancellationToken ct = default)
        {
            try
            {
                var token = await GetAccessTokenAsync(ct).ConfigureAwait(false);
                _httpClient.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", token.Token);
                fromEmail ??= _optionsAuth.MailboxAddress;
                if (_logger != null) _logSendingEmail(_logger, fromEmail, toEmails.Count, null);

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
                    ToRecipients = toEmails.Select(email => new RecipientPayload
                    {
                        EmailAddress = new EmailAddressPayload { Address = email }
                    }).ToList()
                };

                if (ccEmails != null && ccEmails.Count > 0)
                {
                    message.CcRecipients = ccEmails.Select(email => new RecipientPayload
                    {
                        EmailAddress = new EmailAddressPayload { Address = email }
                    }).ToList();
                }

                if (bccEmails != null && bccEmails.Count > 0)
                {
                    message.BccRecipients = bccEmails.Select(email => new RecipientPayload
                    {
                        EmailAddress = new EmailAddressPayload { Address = email }
                    }).ToList();
                }

                var messageJson = JsonConvert.SerializeObject(message, new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Ignore,
                    ContractResolver = new Newtonsoft.Json.Serialization.CamelCasePropertyNamesContractResolver()
                });
                var messageContent = new StringContent(messageJson, Encoding.UTF8, "application/json");

                var messageResponse = await _httpClient.PostAsync(messageUrl, messageContent,ct).ConfigureAwait(false);
                if (!messageResponse.IsSuccessStatusCode)
                {
                    var error = await GetErrorDetailsAsync(messageResponse).ConfigureAwait(false);
      
                    if (_logger != null) _logFailedToCreateMessage(_logger, error, null);
                    throw new Exception($"Failed to create message: {error}");
                }

                var messageResponseBody = await messageResponse.Content.ReadAsStringAsync().ConfigureAwait(false);
                var createdMessage = JObject.Parse(messageResponseBody);
                var messageId = createdMessage["id"]?.ToString() ?? throw new InvalidOperationException("Message ID not found in response");

                if (_logger != null) _logDraftCreated(_logger, messageId, null);              

                // Step 2: Attach files (stream large files, direct upload small files)
                if (attachments != null && attachments.Count > 0)
                {
                    foreach (var attachment in attachments)
                    {
                        var fileInfo = new FileInfo(attachment.FilePath);
                        var fileSize = fileInfo.Length;

                        if (_logger != null) _logAttachingFile(_logger, attachment.FileName, fileSize, null);

                        if (fileSize > 3 * 1024 * 1024) // > 3MB
                        {
                            await UploadLargeAttachmentStreamAsync(fromEmail, messageId,
                                attachment.FileName, attachment.FilePath, fileSize).ConfigureAwait(false);
                        }
                        else
                        {
                            await AddSmallAttachmentAsync(fromEmail, messageId,
                                attachment.FileName, attachment.FilePath).ConfigureAwait(false);
                        }
                    }
                }

                // Step 3: Get the complete message with attachments
                var getMessageUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages/{messageId}?$expand=attachments";
                var getResponse = await _httpClient.GetAsync(getMessageUrl,ct).ConfigureAwait(false);
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
                var sendContent = new StringContent(sendJson, Encoding.UTF8, "application/json");
                var sendResponse = await _httpClient.PostAsync(sendUrl, sendContent, ct).ConfigureAwait(false);

                if (!sendResponse.IsSuccessStatusCode)
                {
                    var error = await GetErrorDetailsAsync(sendResponse).ConfigureAwait(false);
                    if (_logger != null) _logFailedToSendMessage(_logger, error, null);                    
                    throw new Exception($"Failed to send message: {error}");
                }

                // Step 5: Delete the draft message
                var deleteUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages/{messageId}";
                var deleteResponse = await _httpClient.DeleteAsync(deleteUrl,ct).ConfigureAwait(false);

                if (!deleteResponse.IsSuccessStatusCode)
                {
                    var error = await GetErrorDetailsAsync(deleteResponse).ConfigureAwait(false);
                    if (_logger != null) _logFailedToDeleteDraft(_logger, messageId, error, null);
                    throw new Exception($"Failed to delete draft message {messageId}, Error {error}");
                }

                if (_logger != null) _logMessageSent(_logger, messageId, null);                
            }
            catch (Exception ex)
            {
                if (_logger != null) _logFailedToSendMessage(_logger, "", ex);
                throw;
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
            string fileName, string filePath)
        {
            var attachUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages/{messageId}/attachments";

            // Read file and encode to base64
            var fileBytes = File.ReadAllBytes(filePath);
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

            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var response = await _httpClient.PostAsync(attachUrl, content).ConfigureAwait(false);
            response.EnsureSuccessStatusCode();

            if(_logger != null) _logSmallAttachmentAdded(_logger, fileName, null);                 
        }

        private async Task UploadLargeAttachmentStreamAsync(string fromEmail, string messageId,
            string fileName, string filePath, long fileSize)
        {
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
            var sessionContent = new StringContent(sessionJson, Encoding.UTF8, "application/json");

            var sessionResponse = await _httpClient.PostAsync(sessionUrl, sessionContent).ConfigureAwait(false);
            sessionResponse.EnsureSuccessStatusCode();

            var sessionResponseBody = await sessionResponse.Content.ReadAsStringAsync().ConfigureAwait(false);
            var sessionInfo = JObject.Parse(sessionResponseBody);
            var uploadUrl = sessionInfo["uploadUrl"]?.ToString() ?? throw new InvalidOperationException("uploadUrl not found in response");

            if (_logger != null) _logUploadSessionUrl(_logger, uploadUrl, fileName, null);

            // Upload in chunks using streaming (5MB chunks)
            int chunkSize = 5 * 1024 * 1024; // 5MB
            byte[] buffer = new byte[chunkSize];
            long offset = 0;

            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 4096, useAsync: true))
            using (var uploadClient = new HttpClient())
            {
                while (offset < fileSize)
                {
                    int bytesRead = await fileStream.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false);
                    if (bytesRead <= 0) break;

                    long end = offset + bytesRead - 1;

                    var request = new HttpRequestMessage(HttpMethod.Put, uploadUrl)
                    {
                        Content = new ByteArrayContent(buffer, 0, bytesRead)
                    };

                    request.Content.Headers.ContentType =
                        new MediaTypeHeaderValue("application/octet-stream");
                    request.Content.Headers.ContentLength = bytesRead;
                    request.Content.Headers.ContentRange =
                        new ContentRangeHeaderValue(offset, end, fileSize);

                    var response = await uploadClient.SendAsync(request).ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                    {
                        var errorBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        if (_logger != null) _logChunkFailed(_logger, (int)response.StatusCode, response.ReasonPhrase ?? "", errorBody, null);
                        throw new InvalidOperationException($"Chunk upload failed: {(int)response.StatusCode} {response.ReasonPhrase} - {errorBody}");
                    }

                    var responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    if (_logger != null) _logChunkStatus(_logger, end + 1, fileSize, fileName, (int)response.StatusCode, null);

                    // Update offset
                    offset = end + 1;

                    // Check if there are more chunks to upload by looking at nextExpectedRanges
                    if (!string.IsNullOrWhiteSpace(responseBody))
                    {
                        var responseJson = JObject.Parse(responseBody);
                        if (responseJson["nextExpectedRanges"] is JArray nextRanges && nextRanges.Count > 0)
                        {
                            // There are more chunks expected, continue uploading
                            if (_logger != null) _logResponseBodyTrace(_logger, $"Next expected ranges: {string.Join(", ", nextRanges)}", null);
                            continue;
                        }
                    }

                    // No nextExpectedRanges means upload is complete
                    if (_logger != null) _logUploadComplete(_logger, fileName, null);
                    return;
                }
            }

            if (_logger != null) _logUploadComplete(_logger, fileName, null);
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
            _tokenLock?.Dispose();
            _httpClient?.Dispose();
            GC.SuppressFinalize(this);
        }
    }

    // Custom contract resolver for @odata.type
    internal class ODataContractResolver : Newtonsoft.Json.Serialization.DefaultContractResolver
    {
        protected override string ResolvePropertyName(string propertyName)
        {
            if (propertyName == "odataType")
                return "@odata.type";
            return base.ResolvePropertyName(propertyName);
        }
    }

    // Strongly-typed payload classes for better performance and type safety
    internal class MessagePayload
    {
        public string? Subject { get; set; }
        public BodyPayload? Body { get; set; }
        public List<RecipientPayload>? ToRecipients { get; set; }
        public List<RecipientPayload>? CcRecipients { get; set; }
        public List<RecipientPayload>? BccRecipients { get; set; }
    }

    internal class BodyPayload
    {
        public string ContentType { get; set; } = string.Empty;
        public string Content { get; set; } = string.Empty;
    }

    internal class RecipientPayload
    {
        public EmailAddressPayload? EmailAddress { get; set; }       
    }

    internal class EmailAddressPayload
    {
        public string Address { get; set; } = string.Empty;
    }

    public class EmailAttachment
    {
        public string FileName { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
    }
}

/*
Compatible with .NET Standard 2.0

Required NuGet Packages:
- Newtonsoft.Json (>= 12.0.3)
- Microsoft.Extensions.Logging.Abstractions (>= 2.1.0)

Improvements:
1. Streaming for large files - memory efficient
2. Proper token caching with expiration handling
3. ILogger integration for dependency injection
4. Strongly-typed JSON classes for better performance
5. Thread-safe token refresh with SemaphoreSlim
6. Comprehensive logging throughout

Usage Example:
var logger = LoggerFactory.Create(builder => builder.AddConsole()).CreateLogger<GraphMailService>();
var mailService = new GraphMailService(tenantId, clientId, clientSecret, logger);

var attachments = new List<EmailAttachment>
{
    new EmailAttachment 
    { 
        FileName = "large_file.pdf", 
        FilePath = @"C:\path\to\large_file.pdf" 
    }
};

await mailService.SendMailWithLargeAttachmentsAsync(
    fromEmail: "sender@yourdomain.com",
    toEmails: new List<string> { "recipient@example.com" },
    subject: "Test Email with Large Attachments",
    body: "<h1>Hello</h1><p>This email has large attachments.</p>",
    attachments: attachments,
    ccEmails: new List<string> { "cc@example.com" },
    isHtml: true
);
*/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.Extensions.Logging;

namespace MailSenderLib.Services
{
    public class GraphMailService : IDisposable
    {
        private readonly string _tenantId;
        private readonly string _clientId;
        private readonly string _clientSecret;
        private readonly ILogger<GraphMailService>? _logger;
        private readonly HttpClient _httpClient;

        // Token caching with expiration
        private string? _accessToken;
        private DateTimeOffset _tokenExpiration;
        private readonly SemaphoreSlim _tokenLock = new SemaphoreSlim(1, 1);
        private static readonly TimeSpan TokenExpiryBuffer = TimeSpan.FromSeconds(30);

        public GraphMailService(
            string tenantId,
            string clientId,
            string clientSecret,
            ILogger<GraphMailService>? logger = null)
        {
            _tenantId = tenantId ?? throw new ArgumentNullException(nameof(tenantId));
            _clientId = clientId ?? throw new ArgumentNullException(nameof(clientId));
            _clientSecret = clientSecret ?? throw new ArgumentNullException(nameof(clientSecret));
            _logger = logger;
            _httpClient = new HttpClient();
        }

        /// <summary>
        /// Get access token using client credentials flow with proper caching and expiration handling
        /// </summary>
        private async Task<string> GetAccessTokenAsync()
        {
            // Fast path: check if cached token is still valid
            if (!string.IsNullOrEmpty(_accessToken) &&
                _tokenExpiration > DateTimeOffset.UtcNow.Add(TokenExpiryBuffer))
            {
                return _accessToken;
            }

            // Acquire lock for token refresh
            await _tokenLock.WaitAsync();
            try
            {
                // Double-check after acquiring lock
                if (!string.IsNullOrEmpty(_accessToken) &&
                    _tokenExpiration > DateTimeOffset.UtcNow.Add(TokenExpiryBuffer))
                {
                    return _accessToken;
                }

                _logger?.LogDebug("Refreshing access token");

                var tokenUrl = $"https://login.microsoftonline.com/{_tenantId}/oauth2/v2.0/token";

                var content = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("client_id", _clientId),
                    new KeyValuePair<string, string>("client_secret", _clientSecret),
                    new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default"),
                    new KeyValuePair<string, string>("grant_type", "client_credentials")
                });

                var response = await _httpClient.PostAsync(tokenUrl, content);
                response.EnsureSuccessStatusCode();

                var responseBody = await response.Content.ReadAsStringAsync();
                var tokenResponse = JObject.Parse(responseBody);

               
                var tokenObj = tokenResponse["access_token"];
                if (tokenObj == null)
                {
                    // Log the error and stop
                    throw new InvalidOperationException("Access token not found in response.");
                }
                _accessToken = tokenObj.ToString();


                var expiresObj = tokenResponse["expires_in"];
                if (expiresObj == null)
                {
                    throw new InvalidOperationException("expires_in is missing in tokenResponse.");
                }
                var expiresIn = expiresObj.Value<int>();
                _tokenExpiration = DateTimeOffset.UtcNow.AddSeconds(expiresIn);

                _logger?.LogDebug("Access token acquired, expires at {ExpiresAt}", _tokenExpiration);

                return _accessToken;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to acquire access token");
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
        public async Task SendMailWithLargeAttachmentsAsync(
            string fromEmail,
            List<string> toEmails,
            string subject,
            string body,
            List<EmailAttachment> attachments,
            List<string> ccEmails = null,
            bool isHtml = true)
        {
            try
            {
                var token = await GetAccessTokenAsync();
                _httpClient.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", token);

                _logger?.LogInformation("Sending email from {From} to {ToCount} recipients",
                    fromEmail, toEmails.Count);

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

                var messageJson = JsonConvert.SerializeObject(message, new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Ignore,
                    ContractResolver = new Newtonsoft.Json.Serialization.CamelCasePropertyNamesContractResolver()
                });
                var messageContent = new StringContent(messageJson, Encoding.UTF8, "application/json");

                var messageResponse = await _httpClient.PostAsync(messageUrl, messageContent);
                if (!messageResponse.IsSuccessStatusCode)
                {
                    var error = await GetErrorDetailsAsync(messageResponse);
                    _logger?.LogError("Failed to create message: {Error}", error);
                    throw new Exception($"Failed to create message: {error}");
                }

                var messageResponseBody = await messageResponse.Content.ReadAsStringAsync();
                var createdMessage = JObject.Parse(messageResponseBody);
                var messageId = createdMessage["id"].ToString();

                _logger?.LogDebug("Draft message created: {MessageId}", messageId);

                // Step 2: Attach files (stream large files, direct upload small files)
                if (attachments != null && attachments.Count > 0)
                {
                    foreach (var attachment in attachments)
                    {
                        var fileInfo = new FileInfo(attachment.FilePath);
                        var fileSize = fileInfo.Length;

                        _logger?.LogDebug("Attaching file {FileName} ({Size} bytes)",
                            attachment.FileName, fileSize);

                        if (fileSize > 3 * 1024 * 1024) // > 3MB
                        {
                            await UploadLargeAttachmentStreamAsync(fromEmail, messageId,
                                attachment.FileName, attachment.FilePath, fileSize);
                        }
                        else
                        {
                            await AddSmallAttachmentAsync(fromEmail, messageId,
                                attachment.FileName, attachment.FilePath);
                        }
                    }
                }

                // Step 3: Get the complete message with attachments
                var getMessageUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages/{messageId}?$expand=attachments";
                var getResponse = await _httpClient.GetAsync(getMessageUrl);
                getResponse.EnsureSuccessStatusCode();

                var completeMessageBody = await getResponse.Content.ReadAsStringAsync();
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
                var sendResponse = await _httpClient.PostAsync(sendUrl, sendContent);

                if (!sendResponse.IsSuccessStatusCode)
                {
                    var error = await GetErrorDetailsAsync(sendResponse);
                    _logger?.LogError("Failed to send message: {Error}", error);
                    throw new Exception($"Failed to send message: {error}");
                }

                // Step 5: Delete the draft message
                var deleteUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages/{messageId}";
                var deleteResponse = await _httpClient.DeleteAsync(deleteUrl);

                if (!deleteResponse.IsSuccessStatusCode)
                {
                    _logger?.LogWarning("Failed to delete draft message {MessageId}", messageId);
                }

                _logger?.LogInformation("Message sent successfully without saving to Sent Items");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Error sending email");
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
                var attachmentsArray = completeMessage["attachments"] as JArray;
                if (attachmentsArray != null && attachmentsArray.Count > 0)
                {
                    var cleanAttachments = new JArray();
                    foreach (JObject att in attachmentsArray)
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
            var response = await _httpClient.PostAsync(attachUrl, content);
            response.EnsureSuccessStatusCode();

            _logger?.LogDebug("Small attachment added: {FileName}", fileName);
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

            var sessionResponse = await _httpClient.PostAsync(sessionUrl, sessionContent);
            sessionResponse.EnsureSuccessStatusCode();

            var sessionResponseBody = await sessionResponse.Content.ReadAsStringAsync();
            var sessionInfo = JObject.Parse(sessionResponseBody);
            var uploadUrl = sessionInfo["uploadUrl"].ToString();

            _logger?.LogDebug("Upload session created for {FileName}", fileName);

            // Upload in chunks using streaming (5MB chunks)
            int chunkSize = 5 * 1024 * 1024; // 5MB
            byte[] buffer = new byte[chunkSize];
            long offset = 0;

            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 4096, useAsync: true))
            using (var uploadClient = new HttpClient())
            {
                while (offset < fileSize)
                {
                    int bytesRead = await fileStream.ReadAsync(buffer, 0, buffer.Length);
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

                    var response = await uploadClient.SendAsync(request);

                    var responseBody = await response.Content.ReadAsStringAsync();

                    _logger?.LogDebug("Uploaded {Current}/{Total} bytes of {FileName}",
                        end + 1, fileSize, fileName);

                    // 200/201 = complete, 202 = continue
                    if ((int)response.StatusCode == 200 || (int)response.StatusCode == 201)
                    {
                        _logger?.LogDebug("Upload complete: {FileName}", fileName);
                        return;
                    }
                    else if ((int)response.StatusCode != 202)
                    {
                        _logger?.LogError("Chunk upload failed: {Status} {Reason} - {Body}",
                            (int)response.StatusCode, response.ReasonPhrase, responseBody);
                        throw new InvalidOperationException(
                            $"Chunk upload failed: {(int)response.StatusCode} {response.ReasonPhrase} - {responseBody}");
                    }

                    offset = end + 1;
                }
            }

            _logger?.LogDebug("Upload complete: {FileName}", fileName);
        }

        private static async Task<string> GetErrorDetailsAsync(HttpResponseMessage response)
        {
            try
            {
                var errorBody = await response.Content.ReadAsStringAsync();
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
        public string Subject { get; set; }
        public BodyPayload Body { get; set; }
        public List<RecipientPayload> ToRecipients { get; set; }
        public List<RecipientPayload> CcRecipients { get; set; }
    }

    internal class BodyPayload
    {
        public string ContentType { get; set; }
        public string Content { get; set; }
    }

    internal class RecipientPayload
    {
        public EmailAddressPayload EmailAddress { get; set; }
    }

    internal class EmailAddressPayload
    {
        public string Address { get; set; }
    }

    public class EmailAttachment
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
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
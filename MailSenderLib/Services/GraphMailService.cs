using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MailSenderLib.Services
{
    public class GraphMailService
    {
        private readonly string _tenantId;
        private readonly string _clientId;
        private readonly string _clientSecret;
        private string _accessToken;
        private readonly HttpClient _httpClient;

        public GraphMailService(string tenantId, string clientId, string clientSecret)
        {
            _tenantId = tenantId;
            _clientId = clientId;
            _clientSecret = clientSecret;
            _httpClient = new HttpClient();
        }

        /// <summary>
        /// Get access token using client credentials flow
        /// </summary>
        private async Task<string> GetAccessTokenAsync()
        {
            if (!string.IsNullOrEmpty(_accessToken))
                return _accessToken;

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
            _accessToken = tokenResponse["access_token"].ToString();

            return _accessToken;
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

                // Step 1: Create draft message
                var messageUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages";

                var message = new Dictionary<string, object>
                {
                    ["subject"] = subject,
                    ["body"] = new Dictionary<string, string>
                    {
                        ["contentType"] = isHtml ? "HTML" : "Text",
                        ["content"] = body
                    },
                    ["toRecipients"] = toEmails.Select(email => new Dictionary<string, object>
                    {
                        ["emailAddress"] = new Dictionary<string, string>
                        {
                            ["address"] = email
                        }
                    }).ToList()
                };

                if (ccEmails != null && ccEmails.Any())
                {
                    message["ccRecipients"] = ccEmails.Select(email => new Dictionary<string, object>
                    {
                        ["emailAddress"] = new Dictionary<string, string>
                        {
                            ["address"] = email
                        }
                    }).ToList();
                }

                var messageJson = JsonConvert.SerializeObject(message, new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Ignore
                });
                var messageContent = new StringContent(messageJson, Encoding.UTF8, "application/json");

                var messageResponse = await _httpClient.PostAsync(messageUrl, messageContent);
                if (!messageResponse.IsSuccessStatusCode)
                {
                    var error = await GetErrorDetailsAsync(messageResponse);
                    throw new Exception($"Failed to create message: {error}");
                }

                var messageResponseBody = await messageResponse.Content.ReadAsStringAsync();
                var createdMessage = JObject.Parse(messageResponseBody);
                var messageId = createdMessage["id"].ToString();

                Console.WriteLine($"Draft message created: {messageId}");

                // Step 2: Attach files
                foreach (var attachment in attachments)
                {
                    var fileBytes = File.ReadAllBytes(attachment.FilePath);
                    var fileSize = fileBytes.Length;

                    if (fileSize > 3 * 1024 * 1024) // > 3MB
                    {
                        await UploadLargeAttachmentAsync(fromEmail, messageId,
                            attachment.FileName, fileBytes);
                    }
                    else
                    {
                        await AddSmallAttachmentAsync(fromEmail, messageId,
                            attachment.FileName, fileBytes);
                    }
                }

                // Step 3: Get the complete message with attachments
                var getMessageUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages/{messageId}?$expand=attachments";
                var getResponse = await _httpClient.GetAsync(getMessageUrl);
                getResponse.EnsureSuccessStatusCode();

                var completeMessageBody = await getResponse.Content.ReadAsStringAsync();
                var completeMessage = JObject.Parse(completeMessageBody);

                // Remove metadata and read-only fields that Graph API doesn't accept in sendMail
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

                            // Copy only necessary fields
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

                // Step 4: Send using sendMail endpoint with SaveToSentItems = false
                var sendUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/sendMail";

                var sendPayload = new Dictionary<string, object>
                {
                    ["message"] = cleanMessage,
                    ["saveToSentItems"] = false
                };

                var sendJson = JsonConvert.SerializeObject(sendPayload);
                var sendContent = new StringContent(sendJson, Encoding.UTF8, "application/json");
                var sendResponse = await _httpClient.PostAsync(sendUrl, sendContent);

                if (!sendResponse.IsSuccessStatusCode)
                {
                    var error = await GetErrorDetailsAsync(sendResponse);
                    throw new Exception($"Failed to send message: {error}");
                }

                // Step 5: Delete the draft message since we don't need it anymore
                var deleteUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages/{messageId}";
                await _httpClient.DeleteAsync(deleteUrl);

                Console.WriteLine("Message sent successfully without saving to Sent Items");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error sending email: {ex.Message}");
                throw;
            }
        }

        private async Task AddSmallAttachmentAsync(string fromEmail, string messageId,
            string fileName, byte[] fileBytes)
        {
            var attachUrl = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/messages/{messageId}/attachments";

            var attachment = new Dictionary<string, object>
            {
                ["@odata.type"] = "#microsoft.graph.fileAttachment",
                ["name"] = fileName,
                ["contentBytes"] = Convert.ToBase64String(fileBytes)
            };

            var json = JsonConvert.SerializeObject(attachment);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var response = await _httpClient.PostAsync(attachUrl, content);
            response.EnsureSuccessStatusCode();

            Console.WriteLine($"Small attachment added: {fileName}");
        }

        private async Task UploadLargeAttachmentAsync(string fromEmail, string messageId,
            string fileName, byte[] fileBytes)
        {
            var fileSize = fileBytes.Length;

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

            // Upload in chunks (4MB chunks recommended)
            int chunkSize = 4 * 1024 * 1024; // 4MB
            int offset = 0;

            using (var uploadClient = new HttpClient())
            {
                while (offset < fileSize)
                {
                    int end = Math.Min(offset + chunkSize, fileSize);
                    int chunkLength = end - offset;
                    byte[] chunk = new byte[chunkLength];
                    Array.Copy(fileBytes, offset, chunk, 0, chunkLength);

                    var request = new HttpRequestMessage(HttpMethod.Put, uploadUrl)
                    {
                        Content = new ByteArrayContent(chunk)
                    };

                    request.Content.Headers.ContentType =
                        new MediaTypeHeaderValue("application/octet-stream");
                    request.Content.Headers.ContentLength = chunkLength;
                    request.Content.Headers.ContentRange =
                        new ContentRangeHeaderValue(offset, end - 1, fileSize);

                    var response = await uploadClient.SendAsync(request);
                    response.EnsureSuccessStatusCode();

                    Console.WriteLine($"Uploaded {end}/{fileSize} bytes of {fileName}");
                    offset = end;
                }
            }

            Console.WriteLine($"Large attachment uploaded: {fileName}");
        }

        private async Task<string> GetErrorDetailsAsync(HttpResponseMessage response)
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
            _httpClient?.Dispose();
        }
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

Azure AD App Setup:
1. Create App Registration in Azure AD
2. Add Application Permission: Mail.Send
3. Grant admin consent
4. Create client secret

Usage Example:
var mailService = new GraphMailService(tenantId, clientId, clientSecret);

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
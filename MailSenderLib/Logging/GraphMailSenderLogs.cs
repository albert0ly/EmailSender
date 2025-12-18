using Microsoft.Extensions.Logging;
using System;
using System.Net;

namespace MailSenderLib.Logging
{
    internal static partial class GraphMailSenderLogs
    {
        internal static readonly Action<ILogger, Exception?> FailedToAcquireToken =
            LoggerMessage.Define(LogLevel.Error, new EventId(1000, nameof(FailedToAcquireToken)),
                "Failed to acquire access token for GraphMailSender");

        internal static readonly Action<ILogger, Exception?> RefreshingToken =
            LoggerMessage.Define(LogLevel.Debug, new EventId(1001, nameof(RefreshingToken)),
                "Refreshing access token for GraphMailSender");

        internal static readonly Action<ILogger, DateTimeOffset, Exception?> TokenAcquired =
            LoggerMessage.Define<DateTimeOffset>(LogLevel.Debug, new EventId(1002, nameof(TokenAcquired)),
                "Access token acquired, expires on {ExpiresOn}");

        internal static readonly Action<ILogger, string, int, Exception?> SendingEmail =
            LoggerMessage.Define<string, int>(LogLevel.Debug, new EventId(1015, nameof(SendingEmail)),
                "Sending email from {From} to {ToCount} recipients");

        internal static readonly Action<ILogger, string, Exception?> FailedToCreateMessage =
            LoggerMessage.Define<string>(LogLevel.Error, new EventId(1016, nameof(FailedToCreateMessage)),
                "Failed to create message: {Error}");

        internal static readonly Action<ILogger, string, Exception?> DraftCreated =
            LoggerMessage.Define<string>(LogLevel.Debug, new EventId(1017, nameof(DraftCreated)),
               "Draft created {MessageId}");
        internal static readonly Action<ILogger, string, long, Exception?> AttachingFile =
            LoggerMessage.Define<string, long>(LogLevel.Debug, new EventId(1018, nameof(AttachingFile)),
              "Attaching file {FileName} size {FileSize}");

        internal static readonly Action<ILogger, string, Exception?> FailedToSendMessage =
            LoggerMessage.Define<string>(LogLevel.Error, new EventId(1019, nameof(FailedToSendMessage)),
                "Failed to send message: {Error}");

        internal static readonly Action<ILogger, string, string, Exception?> FailedToDeleteDraft =
            LoggerMessage.Define<string, string>(LogLevel.Error, new EventId(1020, nameof(FailedToDeleteDraft)),
                "Failed to delete draft message {MessageId}, Error {Error}");

        internal static readonly Action<ILogger, string, Exception?> MessageSent =
            LoggerMessage.Define<string>(LogLevel.Debug, new EventId(1021, nameof(MessageSent)),
                "Message sent successfully without saving to Sent Items {MessageId}");

        internal static readonly Action<ILogger, string, string, int, int, string, Exception?> UploadSessionUrl =
            LoggerMessage.Define<string, string, int, int, string>(LogLevel.Debug, new EventId(1013, nameof(UploadSessionUrl)),
                "Upload session URL: {Url} for file: {FileName}. Attempt {SessionAttempt}/{MaxSessionRetries}. Draft: {MessageId}");        

        internal static readonly Action<ILogger, long, long, string, int, Exception?> ChunkStatus =
            LoggerMessage.Define<long, long, string, int>(LogLevel.Debug, new EventId(1010, nameof(ChunkStatus)),
                "Uploaded {Current}/{Total} bytes of {FileName}, Status {Status}");

        internal static readonly Action<ILogger, string, Exception?> SmallAttachmentAdded =
            LoggerMessage.Define<string>(LogLevel.Debug, new EventId(1022, nameof(SmallAttachmentAdded)),
                "Small attachment added: {FileName}");

        internal static readonly Action<ILogger, string, Exception?> UploadComplete =
            LoggerMessage.Define<string>(LogLevel.Debug, new EventId(1012, nameof(UploadComplete)),
                "Upload complete for {FileName}");

        internal static readonly Action<ILogger, int, string, string, Exception?> ChunkFailed =
            LoggerMessage.Define<int, string, string>(LogLevel.Error, new EventId(1014, nameof(ChunkFailed)),
                "Chunk upload failed {Status} {Reason} - {Body}");

        internal static readonly Action<ILogger, string, Exception?> ResponseBodyTrace =
            LoggerMessage.Define<string>(LogLevel.Trace, new EventId(1011, nameof(ResponseBodyTrace)), "{Body}");

        internal static readonly Action<ILogger, string, long, long, Exception?> UploadCancelled =
            LoggerMessage.Define<string, long, long>(LogLevel.Error, new EventId(1023, nameof(UploadCancelled)),
                "Upload of '{FileName}' was cancelled at offset {Offset}/{FileSize}");

        internal static readonly Action<ILogger, int, double, HttpStatusCode, string, Exception?> Retrying =
            LoggerMessage.Define<int, double, HttpStatusCode, string>(LogLevel.Warning, new EventId(1024, nameof(Retrying)),
                "Retrying Graph API call. Attempt {RetryAttempt}, waiting {DelaySeconds:F1}s. Status: {StatusCode}. Reason: {Reason}");

        internal static readonly Action<ILogger, long, string, Exception?> ExecutionStep =
            LoggerMessage.Define<long, string>(LogLevel.Information, new EventId(1025, nameof(ExecutionStep)),
                "ExecutionStep: Elapsed: {ElapsedTime} ms. {Description}");

        internal static readonly Action<ILogger, string, int, int,  double, Exception?> SessionExpired =
            LoggerMessage.Define<string, int, int, double>(LogLevel.Warning, new EventId(1026, nameof(SessionExpired)),
                "Upload session failed for '{FileName}' (attempt {SessionAttempt}/{MaxSessionRetries}). Recreating session in {DelaySeconds:F1}s." +
                " This is a known Graph API backend issue.");
    }

    internal static class GraphMailSenderLoggerExtensions
    {
        public static void LogFailedToAcquireToken(this ILogger logger, Exception? ex) =>
            GraphMailSenderLogs.FailedToAcquireToken(logger, ex);

        public static void LogRefreshingToken(this ILogger logger) =>
            GraphMailSenderLogs.RefreshingToken(logger, null);

        public static void LogTokenAcquired(this ILogger logger, DateTimeOffset expiresOn) =>
            GraphMailSenderLogs.TokenAcquired(logger, expiresOn, null);

        public static void LogSendingEmail(this ILogger logger, string from, int toCount) =>
            GraphMailSenderLogs.SendingEmail(logger, from, toCount, null);

        public static void LogFailedToCreateMessage(this ILogger logger, string error) =>
            GraphMailSenderLogs.FailedToCreateMessage(logger, error, null);

        public static void LogDraftCreated(this ILogger logger, string messageId) =>
            GraphMailSenderLogs.DraftCreated(logger, messageId, null);

        public static void LogAttachingFile(this ILogger logger, string fileName, long fileSize) =>
            GraphMailSenderLogs.AttachingFile(logger, fileName, fileSize, null);

        public static void LogFailedToSendMessage(this ILogger logger, string error, Exception? ex=null) =>
            GraphMailSenderLogs.FailedToSendMessage(logger, error, ex);        

        public static void LogFailedToDeleteDraft(this ILogger logger, string messageId, string error, Exception? ex=null) =>
            GraphMailSenderLogs.FailedToDeleteDraft(logger, messageId, error, ex);

        public static void LogMessageSent(this ILogger logger, string messageId) =>
            GraphMailSenderLogs.MessageSent(logger, messageId, null);

        public static void LogUploadSessionUrl(this ILogger logger, string url, string fileName, int sessionAttempt, int maxSessionRetries, string messageId) =>
            GraphMailSenderLogs.UploadSessionUrl(logger, url, fileName, sessionAttempt, maxSessionRetries, messageId, null);

        public static void LogChunkStatus(this ILogger logger, long current, long total, string fileName, int status) =>
            GraphMailSenderLogs.ChunkStatus(logger, current, total, fileName,  status, null);

        public static void LogSmallAttachmentAdded(this ILogger logger, string fileName) =>
            GraphMailSenderLogs.SmallAttachmentAdded(logger, fileName, null);

        public static void LogUploadComplete(this ILogger logger, string fileName) =>
            GraphMailSenderLogs.UploadComplete(logger, fileName, null);

        public static void LogChunkFailed(this ILogger logger, int status, string reason, string body, Exception? ex=null) =>
            GraphMailSenderLogs.ChunkFailed(logger, status, reason, body, ex);

        public static void LogResponseBodyTrace(this ILogger logger, string body) =>
            GraphMailSenderLogs.ResponseBodyTrace(logger, body, null);

        public static void LogUploadCancelled(this ILogger logger, string fileName, long offset, long fileSize, Exception? ex=null) =>
            GraphMailSenderLogs.UploadCancelled(logger, fileName, offset, fileSize, ex);

        public static void LogRetrying(this ILogger logger, int retryAttempt, double timeSpan, HttpStatusCode statusCode, string reason, Exception? ex = null) =>
            GraphMailSenderLogs.Retrying(logger, retryAttempt, timeSpan, statusCode, reason, ex);

        public static void LogExecutionStep(this ILogger logger, string description, long elapsedTime) =>
            GraphMailSenderLogs.ExecutionStep(logger, elapsedTime, description, null);

        public static void LogSessionExpired(this ILogger logger, string fileName, int sessionAttempt, int maxSessionRetries, double delaySeconds, Exception? ex = null) =>
            GraphMailSenderLogs.SessionExpired(logger, fileName, sessionAttempt, maxSessionRetries, delaySeconds, ex);

    }
}

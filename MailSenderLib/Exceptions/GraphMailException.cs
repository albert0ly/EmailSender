// MailSenderLib/Exceptions/GraphMailException.cs
using System;

namespace MailSenderLib.Exceptions
{
    /// <summary>
    /// Base exception for all Graph Mail operations
    /// </summary>
    public class GraphMailException : Exception
    {
        public string? ErrorCode { get; }
        public int? StatusCode { get; }

        public GraphMailException() { }

        public GraphMailException(string message)
            : base(message) { }

        public GraphMailException(string message, Exception innerException)
            : base(message, innerException) { }

        public GraphMailException(string message, string? errorCode, int? statusCode = null)
            : base(message)
        {
            ErrorCode = errorCode;
            StatusCode = statusCode;
        }

        public GraphMailException(string message, string? errorCode, int? statusCode, Exception innerException)
            : base(message, innerException)
        {
            ErrorCode = errorCode;
            StatusCode = statusCode;
        }
    }


    public class GraphMailFailedCreateMessageException : GraphMailException
    {
        public GraphMailFailedCreateMessageException(string message) : base(message) { }
    }
    public class GraphMailFailedSendMessageException : GraphMailException
    {
        public GraphMailFailedSendMessageException(string message) : base(message) { }
    }
    
    public class GraphMailFailedDeleteDraftMessageException : GraphMailException
    {
        public GraphMailFailedDeleteDraftMessageException(string message) : base(message) { }
    }

    /// <summary>
    /// Thrown when attachment operations fail
    /// </summary>
    public class GraphMailAttachmentException : GraphMailException
    {
        public string? FileName { get; }
        public long? FileSize { get; }

        public GraphMailAttachmentException(string message, string? fileName = null, long? fileSize = null)
            : base(message)
        {
            FileName = fileName;
            FileSize = fileSize;
        }

        public GraphMailAttachmentException(string message, Exception innerException, string? fileName = null, long? fileSize = null)
            : base(message, innerException)
        {
            FileName = fileName;
            FileSize = fileSize;
        }
    }
}
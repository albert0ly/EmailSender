# Code Review: GraphMailSender Class

## Executive Summary

The `GraphMailSender` class is a well-structured implementation for sending emails via Microsoft Graph API. It demonstrates good practices in token caching, logging, and error handling. However, there are several areas for improvement related to HttpClient management, interface implementation, resource management, and code consistency.

---

## 🔴 Critical Issues

### 1. HttpClient Management Anti-Pattern
**Location:** Lines 76, 451

**Issue:** 
- The class creates its own `HttpClient` instance (line 76), which can lead to socket exhaustion
- In `UploadLargeAttachmentStreamAsync`, a new `HttpClient` is created for each upload (line 451), which is problematic

**Current Code:**
```csharp
_httpClient = new HttpClient();  // Line 76
// ...
using (var uploadClient = new HttpClient())  // Line 451
```

**Recommendation:**
- Accept `HttpClient` via dependency injection (similar to `GraphMailReceiver`)
- Use `IHttpClientFactory` for better lifecycle management
- Reuse the same `HttpClient` instance for uploads

**Suggested Fix:**
```csharp
public GraphMailSender(
    GraphMailOptionsAuth optionsAuth,
    HttpClient? httpClient = null,
    ILogger<GraphMailSender>? logger = null)
{
    _optionsAuth = optionsAuth ?? throw new ArgumentNullException(nameof(optionsAuth));
    _credential = new ClientSecretCredential(_optionsAuth.TenantId, _optionsAuth.ClientId, _optionsAuth.ClientSecret);
    _logger = logger;
    _httpClient = httpClient ?? new HttpClient();
}
```

### 2. Interface Implementation Mismatch
**Location:** Class declaration

**Issue:** 
- The class doesn't implement `IGraphMailSender` interface
- The interface signature uses `Stream` for attachments, but the class uses `FilePath` strings
- This creates an API inconsistency

**Recommendation:**
- Either implement the interface with proper signature, or
- Create a new interface that matches the current implementation
- Consider supporting both file paths and streams

---

## 🟡 Important Issues

### 3. Hardcoded Configuration Values
**Location:** Line 218

**Issue:** 
- The 3MB threshold for large attachments is hardcoded
- Chunk size (5MB) is also hardcoded (line 446)

**Recommendation:**
- Move these to `GraphMailOptions` or make them configurable
- Consider making them constants with clear names

**Suggested Fix:**
```csharp
private const long LargeAttachmentThreshold = 3 * 1024 * 1024; // 3MB
private const int ChunkSize = 5 * 1024 * 1024; // 5MB
```

### 4. Missing Input Validation
**Location:** `SendEmailAsync` method

**Issue:**
- No validation for email addresses
- No validation for file paths existence
- No validation for empty/null recipients lists
- No validation for subject/body

**Recommendation:**
- Add validation at the beginning of `SendEmailAsync`
- Validate email format (basic or regex)
- Check file existence before processing
- Validate required fields

**Suggested Fix:**
```csharp
if (toRecipients == null || !toRecipients.Any())
    throw new ArgumentException("At least one recipient is required", nameof(toRecipients));

if (string.IsNullOrWhiteSpace(subject))
    throw new ArgumentException("Subject cannot be empty", nameof(subject));

if (attachments != null)
{
    foreach (var attachment in attachments)
    {
        if (!File.Exists(attachment.FilePath))
            throw new FileNotFoundException($"Attachment file not found: {attachment.FilePath}");
    }
}
```

### 5. Inconsistent HttpClient Usage in Upload Method
**Location:** Line 451

**Issue:**
- Creates a new `HttpClient` for uploads instead of reusing the instance
- The upload URL might require different authentication, but this should be handled explicitly

**Recommendation:**
- Reuse `_httpClient` for uploads, or
- Document why a separate client is needed
- Consider if upload URLs require different headers

### 6. Token Expiry Buffer Inconsistency
**Location:** Line 33 vs GraphMailReceiver line 30

**Issue:**
- `GraphMailSender` uses 30 seconds buffer
- `GraphMailReceiver` uses 60 seconds buffer
- Inconsistent across the codebase

**Recommendation:**
- Standardize on one value (60 seconds is more conservative)
- Consider making it configurable

---

## 🟢 Code Quality Improvements

### 7. Simplify Logger Null Checks
**Location:** Throughout the file

**Issue:**
- Repeated `if (_logger != null)` checks

**Recommendation:**
- Use null-conditional operator where possible
- Or use a helper method

**Example:**
```csharp
_logger?.LogDebug("Message sent successfully without saving to Sent Items {MessageId}", messageId);
```

### 8. Improve Error Messages
**Location:** Various exception throws

**Issue:**
- Some error messages could be more descriptive
- Missing context information (e.g., file name in attachment errors)

**Recommendation:**
- Include more context in exception messages
- Use structured logging for better diagnostics

### 9. Magic Numbers
**Location:** Lines 218, 446

**Issue:**
- Magic numbers without clear meaning

**Recommendation:**
- Extract to named constants with comments

### 10. File Stream Buffer Size
**Location:** Line 450

**Issue:**
- Buffer size of 4096 is hardcoded

**Recommendation:**
- Use a constant or make it configurable
- Consider using `FileStream` default or a larger buffer for large files

### 11. Exception Handling in UploadLargeAttachmentStreamAsync
**Location:** Lines 412-504

**Issue:**
- Upload failures could leave partial uploads
- No retry logic for transient failures

**Recommendation:**
- Consider implementing retry logic for transient errors
- Document expected behavior on partial uploads

### 12. CleanMessageForSending Complexity
**Location:** Lines 331-382

**Issue:**
- Method is doing manual field copying which is error-prone
- Could use a more maintainable approach

**Recommendation:**
- Consider using a whitelist approach with reflection
- Or use a mapping library
- Document which fields are required vs optional

### 13. Missing XML Documentation
**Location:** Public methods

**Issue:**
- `SendEmailAsync` has good documentation
- Other public/internal methods lack XML docs

**Recommendation:**
- Add XML documentation for all public methods
- Document parameters, return values, and exceptions

### 14. EmailAttachment Class Location
**Location:** Line 565

**Issue:**
- `EmailAttachment` is defined at the bottom of the file
- Should be in a separate Models file for better organization

**Recommendation:**
- Move to `MailSenderLib/Models/EmailAttachment.cs`
- Keep it consistent with other model classes

### 15. ODataContractResolver Class
**Location:** Line 529

**Issue:**
- Internal class in the same file
- Could be in a separate file for better organization

**Recommendation:**
- Move to a separate file or namespace
- Consider if this is reusable elsewhere

---

## 🔵 Performance Considerations

### 16. Token Caching
**Status:** ✅ Well Implemented
- Good double-check locking pattern
- Proper thread safety with `SemaphoreSlim`

### 17. LoggerMessage Delegates
**Status:** ✅ Well Implemented
- Good use of `LoggerMessage.Define` for performance
- Avoids allocation-heavy logging

### 18. File Reading for Small Attachments
**Location:** Line 390

**Issue:**
- `File.ReadAllBytes` loads entire file into memory
- For files close to 3MB, this could be inefficient

**Recommendation:**
- Consider streaming even for smaller files if they're close to the threshold
- Or document the memory implications

### 19. JSON Serialization
**Location:** Multiple locations

**Issue:**
- Creating new `JsonSerializerSettings` instances
- Could cache settings for better performance

**Recommendation:**
- Create static readonly settings instances
- Reuse across serialization calls

---

## 🟣 Consistency Issues

### 20. Constructor Pattern
**Issue:**
- `GraphMailSender` doesn't accept optional `HttpClient` like `GraphMailReceiver` does
- Inconsistent API design

**Recommendation:**
- Align constructor signatures
- Make `HttpClient` optional parameter

### 21. Logger Parameter Type
**Issue:**
- `GraphMailReceiver` accepts `object? logger` and casts
- `GraphMailSender` accepts `ILogger<GraphMailSender>?` directly
- Inconsistent approach

**Recommendation:**
- Standardize on one approach (prefer strongly-typed `ILogger<T>`)

### 22. Error Handling Pattern
**Issue:**
- Different exception types and handling patterns
- Could be more consistent

**Recommendation:**
- Review exception hierarchy
- Ensure consistent error handling patterns

---

## 📋 Summary of Recommendations

### High Priority
1. ✅ Fix HttpClient management (accept via DI, reuse instance)
2. ✅ Implement or align with interface
3. ✅ Add input validation
4. ✅ Fix HttpClient creation in upload method

### Medium Priority
5. ✅ Extract hardcoded values to constants/config
6. ✅ Standardize token expiry buffer
7. ✅ Move EmailAttachment to Models folder
8. ✅ Add XML documentation
9. ✅ Improve error messages with context

### Low Priority
10. ✅ Simplify logger null checks
11. ✅ Cache JsonSerializerSettings
12. ✅ Consider retry logic for uploads
13. ✅ Improve CleanMessageForSending maintainability

---

## 🎯 Positive Aspects

1. ✅ Excellent token caching implementation with thread safety
2. ✅ Good use of LoggerMessage for performance
3. ✅ Proper disposal pattern
4. ✅ Good error handling with specific exception types
5. ✅ Clean separation of concerns (draft creation, attachment upload, sending)
6. ✅ Proper use of ConfigureAwait(false)
7. ✅ Good logging coverage

---

## 📝 Additional Notes

- The class handles large attachments well with chunked uploads
- The draft cleanup pattern is good (create → attach → send → delete)
- Consider adding unit tests for the token caching logic
- Consider adding integration tests for the full email sending flow
- The `saveToSentItems = false` behavior should be documented


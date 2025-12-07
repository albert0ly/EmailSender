# Using IHttpClientFactory with GraphMailSender

## Overview

`IHttpClientFactory` is the recommended way to manage `HttpClient` instances in .NET applications. It prevents socket exhaustion issues and properly manages the lifecycle of HTTP connections.

## Benefits of IHttpClientFactory

1. **Prevents Socket Exhaustion**: Reuses underlying HTTP connections efficiently
2. **Automatic Lifecycle Management**: Handles connection pooling and disposal
3. **Configuration Support**: Allows named clients with specific configurations
4. **Better Performance**: Reduces overhead from creating/disposing HttpClient instances

## Setup

### 1. Add Required Package

The `Microsoft.Extensions.Http` package is already included in `MailSenderLib.csproj`.

### 2. Register IHttpClientFactory in DI Container

#### ASP.NET Core (Program.cs or Startup.cs)

```csharp
using Microsoft.Extensions.DependencyInjection;

var builder = WebApplication.CreateBuilder(args);

// Register IHttpClientFactory (automatically available in ASP.NET Core)
// No explicit registration needed - it's included by default

// Optionally configure a named HttpClient for GraphMailSender
builder.Services.AddHttpClient("GraphMailSender", client =>
{
    client.Timeout = TimeSpan.FromMinutes(5); // Adjust as needed
    client.BaseAddress = new Uri("https://graph.microsoft.com/");
});

// Register GraphMailSender with DI
builder.Services.AddScoped<GraphMailSender>(serviceProvider =>
{
    var optionsAuth = serviceProvider.GetRequiredService<IOptions<GraphMailOptionsAuth>>().Value;
    var httpClientFactory = serviceProvider.GetRequiredService<IHttpClientFactory>();
    var logger = serviceProvider.GetService<ILogger<GraphMailSender>>();
    
    return new GraphMailSender(
        optionsAuth,
        httpClientFactory: httpClientFactory,
        logger: logger);
});
```

#### Console Application / Other Hosts

```csharp
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var services = new ServiceCollection();

// Register IHttpClientFactory
services.AddHttpClient();

// Optionally configure named client
services.AddHttpClient("GraphMailSender", client =>
{
    client.Timeout = TimeSpan.FromMinutes(5);
});

// Register GraphMailSender
services.AddScoped<GraphMailSender>(serviceProvider =>
{
    var optionsAuth = new GraphMailOptionsAuth
    {
        TenantId = "your-tenant-id",
        ClientId = "your-client-id",
        ClientSecret = "your-client-secret",
        MailboxAddress = "sender@domain.com"
    };
    
    var httpClientFactory = serviceProvider.GetRequiredService<IHttpClientFactory>();
    var logger = serviceProvider.GetService<ILogger<GraphMailSender>>();
    
    return new GraphMailSender(
        optionsAuth,
        httpClientFactory: httpClientFactory,
        logger: logger);
});

var serviceProvider = services.BuildServiceProvider();
var mailSender = serviceProvider.GetRequiredService<GraphMailSender>();
```

## Usage Examples

### Example 1: Direct Instantiation with IHttpClientFactory

```csharp
using Microsoft.Extensions.Http;

// Create IHttpClientFactory (in real apps, inject via DI)
var services = new ServiceCollection();
services.AddHttpClient();
var serviceProvider = services.BuildServiceProvider();
var httpClientFactory = serviceProvider.GetRequiredService<IHttpClientFactory>();

// Create GraphMailSender
var optionsAuth = new GraphMailOptionsAuth
{
    TenantId = "your-tenant-id",
    ClientId = "your-client-id",
    ClientSecret = "your-client-secret",
    MailboxAddress = "sender@domain.com"
};

var mailSender = new GraphMailSender(
    optionsAuth,
    httpClientFactory: httpClientFactory);

// Use it
await mailSender.SendEmailAsync(
    toRecipients: new List<string> { "recipient@domain.com" },
    ccRecipients: null,
    bccRecipients: null,
    subject: "Test Email",
    body: "This is a test",
    isHtml: false,
    attachments: null);
```

### Example 2: Dependency Injection in ASP.NET Core Controller

```csharp
using MailSenderLib.Services;
using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("api/[controller]")]
public class EmailController : ControllerBase
{
    private readonly GraphMailSender _mailSender;

    public EmailController(GraphMailSender mailSender)
    {
        _mailSender = mailSender;
    }

    [HttpPost("send")]
    public async Task<IActionResult> SendEmail([FromBody] EmailRequest request)
    {
        await _mailSender.SendEmailAsync(
            toRecipients: request.To,
            ccRecipients: request.Cc,
            bccRecipients: request.Bcc,
            subject: request.Subject,
            body: request.Body,
            isHtml: request.IsHtml,
            attachments: request.Attachments?.Select(a => new EmailAttachment
            {
                FileName = a.FileName,
                FilePath = a.FilePath
            }).ToList());

        return Ok();
    }
}
```

### Example 3: Fallback to Direct HttpClient (Not Recommended)

```csharp
// This still works but is NOT recommended for production
var httpClient = new HttpClient();
var mailSender = new GraphMailSender(
    optionsAuth,
    httpClient: httpClient); // Falls back to direct HttpClient

// Note: You must manage HttpClient lifecycle yourself
```

## Constructor Parameter Priority

The constructor accepts parameters in this priority order:

1. **IHttpClientFactory** (highest priority, recommended)
2. **HttpClient** (fallback, for testing or legacy code)
3. **new HttpClient()** (lowest priority, creates new instance)

```csharp
// Priority 1: IHttpClientFactory (best)
new GraphMailSender(optionsAuth, httpClientFactory: factory);

// Priority 2: HttpClient (fallback)
new GraphMailSender(optionsAuth, httpClient: client);

// Priority 3: new HttpClient() (not recommended)
new GraphMailSender(optionsAuth); // Creates new HttpClient internally
```

## Configuration Options

### Named HttpClient Configuration

You can configure a named HttpClient with specific settings:

```csharp
builder.Services.AddHttpClient("GraphMailSender", client =>
{
    client.Timeout = TimeSpan.FromMinutes(5);
    client.BaseAddress = new Uri("https://graph.microsoft.com/");
    
    // Add default headers if needed
    client.DefaultRequestHeaders.Add("User-Agent", "MailSenderLib/1.0");
});
```

### HttpClient Lifecycle

When using `IHttpClientFactory`:
- **You don't need to dispose** the HttpClient - the factory manages it
- The `Dispose()` method in `GraphMailSender` will **not** dispose HttpClient from factory
- Connections are pooled and reused automatically

## Migration Guide

### Before (Old Code)

```csharp
// ❌ Old way - creates new HttpClient
var mailSender = new GraphMailSender(optionsAuth, logger: logger);
```

### After (New Code)

```csharp
// ✅ New way - uses IHttpClientFactory
var mailSender = new GraphMailSender(
    optionsAuth,
    httpClientFactory: httpClientFactory,
    logger: logger);
```

## Important Notes

1. **Dispose Pattern**: When using `IHttpClientFactory`, the `GraphMailSender.Dispose()` method will **not** dispose the HttpClient because the factory manages its lifecycle.

2. **Thread Safety**: `IHttpClientFactory` creates thread-safe HttpClient instances that can be safely reused.

3. **Testing**: For unit tests, you can still inject a mock `HttpClient` directly:
   ```csharp
   var mockHttpClient = new Mock<HttpClient>();
   var mailSender = new GraphMailSender(optionsAuth, httpClient: mockHttpClient.Object);
   ```

4. **Backward Compatibility**: The old constructor signature still works, but it's recommended to migrate to `IHttpClientFactory` for production code.

## Troubleshooting

### Issue: Socket Exhaustion

**Symptom**: `SocketException` or "Unable to connect" errors after many requests.

**Solution**: Use `IHttpClientFactory` instead of creating new `HttpClient` instances.

### Issue: HttpClient Not Disposed

**Symptom**: Memory leaks or connection issues.

**Solution**: When using `IHttpClientFactory`, you don't need to dispose HttpClient manually. The factory handles it. If you're using direct `HttpClient`, ensure proper disposal.

### Issue: Timeout Issues

**Solution**: Configure timeout in named HttpClient:
```csharp
services.AddHttpClient("GraphMailSender", client =>
{
    client.Timeout = TimeSpan.FromMinutes(10); // For large file uploads
});
```


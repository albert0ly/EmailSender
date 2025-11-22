MailSenderLib

This project contains lightweight helpers to send/receive mail using Microsoft Graph APIs. It targets `netstandard2.0` so it can be consumed by multiple runtimes.

Structure
- `Models/` - DTOs used by the library
- `Options/` - configuration option types
- `Interfaces/` - public interfaces
- `Services/` - concrete service implementations

Notes
- Interface `IGraphMailSender` is implemented by `Services.GraphMailSender`.
- `GraphMailReceiver` is a consumer-side helper that uses `Options.GraphMailOptions` for configuration; it is not intended to implement `IGraphMailSender` because it performs message retrieval operations rather than sending.

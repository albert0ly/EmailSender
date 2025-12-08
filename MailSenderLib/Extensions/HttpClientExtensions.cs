using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MailSenderLib.Extensions
{
    internal static class HttpClientExtensions
    {
        public static async Task<HttpResponseMessage> SendJsonWithTokenAsync(
        this HttpClient httpClient,
        HttpMethod method,
        string requestUrl,
        string? jsonPayload = null,
        string? bearerToken = null,
        CancellationToken cancellationToken = default)
        {
            using var request = new HttpRequestMessage(method, requestUrl);

            // Only attach JSON payload if provided and method allows a body
            if (!string.IsNullOrWhiteSpace(jsonPayload) &&
                (method == HttpMethod.Post || method == HttpMethod.Put || method.Method == "PATCH"))
            {
                request.Content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
            }

            // Add Authorization header if token is provided
            if (!string.IsNullOrWhiteSpace(bearerToken))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
            }

            return await httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
        }
    }
}

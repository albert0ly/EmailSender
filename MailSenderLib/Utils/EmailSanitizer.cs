using Ganss.Xss;
using System.Text.RegularExpressions;

namespace MailSenderLib.Utils
{

    internal static class EmailSanitizer
    {

        public static string SanitizeSubject(string subject)
        {
            if (subject == null) return string.Empty;

            // Remove CR/LF to prevent header injection
            string sanitized = subject.Replace("\r", "").Replace("\n", "");

            // Remove control characters
            sanitized = Regex.Replace(sanitized, @"[\x00-\x1F\x7F]", "");

            // Trim and enforce max length
            if (sanitized.Length > 255)
                sanitized = sanitized.Substring(0, 255);

            return sanitized.Trim();
        }

        public static HtmlSanitizer CreateEmailSanitizer()
        {
            var sanitizer = new HtmlSanitizer();

            // Allow basic formatting
            sanitizer.AllowedTags.Add("p");
            sanitizer.AllowedTags.Add("br");
            sanitizer.AllowedTags.Add("div");
            sanitizer.AllowedTags.Add("span");
            sanitizer.AllowedTags.Add("b");
            sanitizer.AllowedTags.Add("i");
            sanitizer.AllowedTags.Add("u");
            sanitizer.AllowedTags.Add("strong");
            sanitizer.AllowedTags.Add("em");

            // Allow lists
            sanitizer.AllowedTags.Add("ul");
            sanitizer.AllowedTags.Add("ol");
            sanitizer.AllowedTags.Add("li");

            // Allow tables (Outlook often uses them)
            sanitizer.AllowedTags.Add("table");
            sanitizer.AllowedTags.Add("thead");
            sanitizer.AllowedTags.Add("tbody");
            sanitizer.AllowedTags.Add("tr");
            sanitizer.AllowedTags.Add("td");
            sanitizer.AllowedTags.Add("th");

            // Allow images
            sanitizer.AllowedTags.Add("img");
            sanitizer.AllowedAttributes.Add("src");
            sanitizer.AllowedAttributes.Add("alt");
            sanitizer.AllowedAttributes.Add("title");
            sanitizer.AllowedAttributes.Add("width");
            sanitizer.AllowedAttributes.Add("height");

            // Allow safe attributes
            sanitizer.AllowedAttributes.Add("style");
            sanitizer.AllowedAttributes.Add("class");
            sanitizer.AllowedAttributes.Add("align");

            // Allow safe CSS properties
            sanitizer.AllowedCssProperties.Add("color");
            sanitizer.AllowedCssProperties.Add("background-color");
            sanitizer.AllowedCssProperties.Add("font-size");
            sanitizer.AllowedCssProperties.Add("font-family");
            sanitizer.AllowedCssProperties.Add("text-align");
            sanitizer.AllowedCssProperties.Add("margin");
            sanitizer.AllowedCssProperties.Add("padding");
            sanitizer.AllowedCssProperties.Add("border");

            // Allow safe schemes
            sanitizer.AllowedSchemes.Add("http");
            sanitizer.AllowedSchemes.Add("https");
            sanitizer.AllowedSchemes.Add("data"); // inline base64 images
            sanitizer.AllowedSchemes.Add("cid"); // REQUIRED for inline email images
            return sanitizer;
        }

        public static string SanitizeBody(string htmlBody)
        {
            var sanitizer = CreateEmailSanitizer();
            return sanitizer.Sanitize(htmlBody);
        }
    }
}

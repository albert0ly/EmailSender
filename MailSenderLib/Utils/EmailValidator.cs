using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;

namespace MailSenderLib.Utils
{
    internal static class EmailValidator
    {
        private static readonly Regex EmailRegex = new Regex(
            @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static bool IsValidEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return false;

            email = email.Trim();

            if (email.Length > 254)
                return false;

            try
            {
                var mailAddress = new MailAddress(email);
                return mailAddress.Address == email && EmailRegex.IsMatch(email);
            }
            catch
            {
                return false;
            }
        }

        public static bool IsValidEmailList(IEnumerable<string> emails)
        {
            if (emails == null)
                return false;

            bool hasAny = false;
            foreach (var email in emails)
            {
                hasAny = true;
                if (!IsValidEmail(email))
                    return false;
            }

            return hasAny;
        }
    }
}

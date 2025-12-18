using System;
using System.Collections.Generic;
using System.Text;

namespace MailSenderLib.Extensions
{
    internal static  class StringExtensions
    {
        /// <summary>
        /// Returns the substring before the first occurrence of the given delimiter.
        /// If the delimiter is not found, returns the original string.
        /// </summary>
        public static string StripAfter(this string input, char delimiter)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            int idx = input.IndexOf(delimiter);
            return idx == -1 ? input : input.Substring(0, idx);
        }
    }
}

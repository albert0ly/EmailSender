using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MailSenderLib.Extensions
{
    internal static class FileExtensions
    {
        /// <summary>
        /// Asynchronously reads all bytes from a file using FileStream.
        /// Works in .NET Standard 2.0 where File.ReadAllBytesAsync is unavailable.
        /// </summary>
        /// <param name="filePath">Path to the file.</param>
        /// <returns>Byte array containing file contents.</returns>
        public static async Task<byte[]> ReadAllBytesAsync(this string filePath, CancellationToken? ct=null)
        {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));

            using (var stream = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                bufferSize: 4096,
                useAsync: true))
            {
                var bytes = new byte[stream.Length];
                int read = 0;
                while (read < bytes.Length)
                {
                    ct?.ThrowIfCancellationRequested();
                    int chunk = await stream.ReadAsync(bytes, read, bytes.Length - read).ConfigureAwait(false);
                    if (chunk == 0) break; // End of stream
                    read += chunk;
                }
                return bytes;
            }
        }

        /// <summary>
        /// Asynchronously reads all bytes from a FileInfo.
        /// </summary>
        public static Task<byte[]> ReadAllBytesAsync(this FileInfo fileInfo)
        {
            if (fileInfo == null) throw new ArgumentNullException(nameof(fileInfo));
            return fileInfo.FullName.ReadAllBytesAsync();
        }
    }
}

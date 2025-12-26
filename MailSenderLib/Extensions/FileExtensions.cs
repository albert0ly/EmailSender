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
        public static int BufferSize { get; set; } = 81920; // 8 KB chunks

        public static async Task<string> StreamFileAsBase64Async(this string filePath, CancellationToken ct = default)
        {
            var buffer = new byte[BufferSize];
            var sb = new StringBuilder();

            using var fs = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                BufferSize,
                useAsync: true);
            {
                int bytesRead;
                while ((bytesRead = await fs.ReadAsync(buffer, 0, buffer.Length, ct).ConfigureAwait(false)) > 0)
                {
                    ct.ThrowIfCancellationRequested();
                    // Convert only the chunk we read
                    string chunkBase64 = Convert.ToBase64String(buffer, 0, bytesRead);
                    sb.Append(chunkBase64);
                }
            }

            return sb.ToString();
        }


        /// <summary>
        /// Asynchronously reads all bytes from a file using FileStream.
        /// Works in .NET Standard 2.0 where File.ReadAllBytesAsync is unavailable.
        /// </summary>
        /// <param name="filePath">Path to the file.</param>
        /// <returns>Byte array containing file contents.</returns>
        public static async Task<byte[]> ReadAllBytesAsync(this string filePath, CancellationToken? ct = null)
        {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));

            using var stream = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                bufferSize: BufferSize,
                useAsync: true);
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

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using StatementFile.Domain.Interfaces.Services;

namespace StatementFile.Infrastructure.Services
{
    /// <summary>
    /// Packages one or more output files into a ZIP archive and generates
    /// a sidecar MD5 checksum file.
    ///
    /// Uses System.IO.Compression.ZipArchive (built-in to .NET Framework 4.5+)
    /// in place of the legacy SharpZip calls from clsBasFile.
    /// MD5 behaviour matches the legacy generateFileMD5() method:
    ///   – hash is written as a HEX string to {filePath}.MD5
    /// </summary>
    public sealed class FilePackagingService : IFilePackagingService
    {
        public string PackageFiles(IEnumerable<string> inputPaths, string zipOutputPath)
        {
            if (inputPaths == null)    throw new ArgumentNullException(nameof(inputPaths));
            if (string.IsNullOrWhiteSpace(zipOutputPath))
                throw new ArgumentException("zipOutputPath must not be empty.", nameof(zipOutputPath));

            string dir = Path.GetDirectoryName(zipOutputPath);
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using (var zip = ZipFile.Open(zipOutputPath, ZipArchiveMode.Create))
            {
                foreach (string input in inputPaths)
                {
                    if (File.Exists(input))
                        zip.CreateEntryFromFile(input, Path.GetFileName(input),
                            CompressionLevel.Optimal);
                }
            }

            GenerateMd5(zipOutputPath);
            return zipOutputPath;
        }

        public string GenerateMd5(string filePath)
        {
            using (var md5    = MD5.Create())
            using (var stream = File.OpenRead(filePath))
            {
                byte[] hash    = md5.ComputeHash(stream);
                string hex     = BitConverter.ToString(hash).Replace("-", "").ToUpperInvariant();
                string md5Path = filePath + ".MD5";
                File.WriteAllText(md5Path, hex, Encoding.ASCII);
                return hex;
            }
        }
    }
}

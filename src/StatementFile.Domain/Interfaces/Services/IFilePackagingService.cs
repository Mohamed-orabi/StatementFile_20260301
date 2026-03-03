using System.Collections.Generic;

namespace StatementFile.Domain.Interfaces.Services
{
    /// <summary>
    /// Packages one or more output files into a ZIP archive and generates
    /// a side-car MD5 checksum file.
    ///
    /// Used by all RawData, XML, and some HTML runs.
    /// Wraps the legacy SharpZip + generateFileMD5() calls from clsBasFile.
    /// </summary>
    public interface IFilePackagingService
    {
        /// <summary>
        /// Compresses <paramref name="inputPaths"/> into <paramref name="zipOutputPath"/>
        /// and writes a matching <c>.MD5</c> file next to the ZIP.
        /// Returns the path of the generated ZIP file.
        /// </summary>
        string PackageFiles(IEnumerable<string> inputPaths, string zipOutputPath);

        /// <summary>
        /// Computes the MD5 hash of a single file and writes it to a <c>.MD5</c>
        /// sidecar file next to the original.
        /// Returns the MD5 hex string.
        /// </summary>
        string GenerateMd5(string filePath);
    }
}

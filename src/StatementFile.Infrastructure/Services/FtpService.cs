using System;
using System.IO;
using System.Net;
using StatementFile.Application.Interfaces;
using StatementFile.Domain.Interfaces.Services;

namespace StatementFile.Infrastructure.Services
{
    /// <summary>
    /// FTP implementation of <see cref="IFtpService"/>.
    /// Mirrors the legacy clsFtpCommandLine.sendFile2ftpSilent() destination convention:
    ///   Live-Banks/{bankName}/To/{filename}
    /// Connection settings are resolved from <see cref="IConfigurationService"/>.
    /// </summary>
    public sealed class FtpService : IFtpService
    {
        private readonly IConfigurationService _config;

        public FtpService(IConfigurationService config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
        }

        public bool SendFile(string localFilePath, string bankName)
        {
            if (!File.Exists(localFilePath))
                return false;

            string fileName = Path.GetFileName(localFilePath);
            string ftpUri   = $"ftp://Live-Banks/{bankName}/To/{fileName}";

            try
            {
                var request = (FtpWebRequest)WebRequest.Create(ftpUri);
                request.Method    = WebRequestMethods.Ftp.UploadFile;
                request.UsePassive = true;
                request.UseBinary = true;
                request.KeepAlive = false;

                byte[] fileBytes = File.ReadAllBytes(localFilePath);
                request.ContentLength = fileBytes.Length;

                using (Stream stream = request.GetRequestStream())
                    stream.Write(fileBytes, 0, fileBytes.Length);

                using (var response = (FtpWebResponse)request.GetResponse())
                    return response.StatusCode == FtpStatusCode.ClosingData ||
                           response.StatusCode == FtpStatusCode.FileActionOK;
            }
            catch
            {
                return false;
            }
        }
    }
}

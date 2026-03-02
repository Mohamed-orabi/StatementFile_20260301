namespace StatementFile.Domain.Interfaces.Services
{
    /// <summary>
    /// Abstraction over FTP file delivery.
    /// Wraps the legacy clsFtpCommandLine.sendFile2ftpSilent() behaviour.
    /// </summary>
    public interface IFtpService
    {
        /// <summary>
        /// Transfers a local file to the bank's inbound FTP directory silently.
        /// Destination convention: Live-Banks/{bankName}/To/
        /// </summary>
        bool SendFile(string localFilePath, string bankName);
    }
}

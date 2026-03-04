using System.Collections.Generic;

namespace StatementFile.Domain.Interfaces.Services
{
    /// <summary>
    /// Abstraction over SMTP email delivery.
    /// Kept free of any infrastructure specifics so it can be mocked in tests.
    /// </summary>
    public interface IEmailService
    {
        /// <summary>
        /// Sends an HTML email with optional file attachments.
        /// </summary>
        void SendHtml(
            string fromAddress,
            string fromDisplayName,
            IEnumerable<string> toAddresses,
            IEnumerable<string> ccAddresses,
            IEnumerable<string> bccAddresses,
            string subject,
            string htmlBody,
            IEnumerable<string> attachmentPaths = null);
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using StatementFile.Application.Interfaces;
using StatementFile.Domain.Interfaces.Services;

namespace StatementFile.Infrastructure.Services
{
    /// <summary>
    /// SMTP implementation of <see cref="IEmailService"/>.
    /// Uses System.Net.Mail (framework-native); SMTP server comes from configuration.
    /// </summary>
    public sealed class EmailService : IEmailService
    {
        private readonly IConfigurationService _config;

        public EmailService(IConfigurationService config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
        }

        public void SendHtml(
            string              fromAddress,
            string              fromDisplayName,
            IEnumerable<string> toAddresses,
            IEnumerable<string> ccAddresses,
            IEnumerable<string> bccAddresses,
            string              subject,
            string              htmlBody,
            IEnumerable<string> attachmentPaths = null)
        {
            if (toAddresses == null) throw new ArgumentNullException(nameof(toAddresses));

            string smtpServer = _config.GetSmtpServer();

            using (var client  = new SmtpClient(smtpServer))
            using (var message = new MailMessage())
            {
                message.From            = new MailAddress(fromAddress, fromDisplayName, Encoding.UTF8);
                message.Subject         = subject;
                message.Body            = htmlBody;
                message.IsBodyHtml      = true;
                message.BodyEncoding    = Encoding.UTF8;
                message.SubjectEncoding = Encoding.UTF8;

                foreach (string to in toAddresses.Where(e => !string.IsNullOrWhiteSpace(e)))
                    message.To.Add(to);

                if (ccAddresses != null)
                    foreach (string cc in ccAddresses.Where(e => !string.IsNullOrWhiteSpace(e)))
                        message.CC.Add(cc);

                if (bccAddresses != null)
                    foreach (string bcc in bccAddresses.Where(e => !string.IsNullOrWhiteSpace(e)))
                        message.Bcc.Add(bcc);

                if (attachmentPaths != null)
                    foreach (string path in attachmentPaths.Where(p => !string.IsNullOrWhiteSpace(p)))
                        message.Attachments.Add(new Attachment(path));

                client.Send(message);
            }
        }
    }
}

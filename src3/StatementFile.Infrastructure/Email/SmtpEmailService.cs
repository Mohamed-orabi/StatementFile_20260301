using System;
using System.Net;
using System.Net.Mail;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Infrastructure.Email
{
    /// <summary>
    /// Sends statement emails via SMTP.
    /// Replaces the hardcoded SmtpClient calls in frmStatementFile.SendBankMail().
    /// Configuration is injected via <see cref="SmtpSettings"/> and read from
    /// appsettings.json / environment variables.
    /// </summary>
    public sealed class SmtpEmailService : IEmailService
    {
        private readonly SmtpSettings _settings;

        public SmtpEmailService(SmtpSettings settings) =>
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));

        public void Send(EmailMessage message)
        {
            if (message == null) throw new ArgumentNullException(nameof(message));

            using var client = new SmtpClient(_settings.Host, _settings.Port)
            {
                EnableSsl   = _settings.UseSsl,
                Credentials = string.IsNullOrWhiteSpace(_settings.Username)
                    ? null
                    : new NetworkCredential(_settings.Username, _settings.Password),
            };

            using var mail = new MailMessage
            {
                From       = new MailAddress(message.FromAddress, message.FromName),
                Subject    = message.Subject,
                Body       = message.Body,
                IsBodyHtml = message.IsBodyHtml,
            };

            if (message.ToAddresses != null)
                foreach (var to in message.ToAddresses)
                    if (!string.IsNullOrWhiteSpace(to))
                        mail.To.Add(to);

            if (message.Attachments != null)
                foreach (var path in message.Attachments)
                    if (!string.IsNullOrWhiteSpace(path) && System.IO.File.Exists(path))
                        mail.Attachments.Add(new Attachment(path));

            client.Send(mail);
        }
    }

    public sealed class SmtpSettings
    {
        public string Host     { get; set; }
        public int    Port     { get; set; } = 25;
        public string Username { get; set; }
        public string Password { get; set; }
        public bool   UseSsl   { get; set; }
    }
}

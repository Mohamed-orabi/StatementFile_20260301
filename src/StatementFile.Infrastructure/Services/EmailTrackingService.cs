using System;
using System.IO;
using StatementFile.Domain.Interfaces.Services;

namespace StatementFile.Infrastructure.Services
{
    /// <summary>
    /// Writes the two per-run email-tracking text files.
    ///
    /// Files produced:
    ///   {prefix}_Emails.txt       – pipe-delimited rows for accounts with a valid email.
    ///   {prefix}_WithoutEmails.txt – pipe-delimited rows for accounts without an email.
    ///
    /// Header (preserved from legacy clsStatHtml output):
    ///   AccountNumber | ClientID | Email | MobilePhone | DateTime
    /// </summary>
    public sealed class EmailTrackingService : IEmailTrackingService
    {
        private string       _emailsFilePath;
        private string       _noEmailsFilePath;
        private StreamWriter _withEmailWriter;
        private StreamWriter _withoutEmailWriter;
        private int          _withEmailCount;
        private int          _withoutEmailCount;
        private bool         _initialised;

        public void Initialise(string outputDirectory, string filePrefix)
        {
            if (_initialised)
                Finalise();

            _emailsFilePath   = Path.Combine(outputDirectory, filePrefix + "_Emails.txt");
            _noEmailsFilePath = Path.Combine(outputDirectory, filePrefix + "_WithoutEmails.txt");

            const string header = "AccountNumber|ClientID|Email|MobilePhone|DateTime";

            _withEmailWriter    = new StreamWriter(_emailsFilePath,   append: false);
            _withoutEmailWriter = new StreamWriter(_noEmailsFilePath, append: false);
            _withEmailWriter.WriteLine(header);
            _withoutEmailWriter.WriteLine(header);

            _withEmailCount    = 0;
            _withoutEmailCount = 0;
            _initialised       = true;
        }

        public void RecordWithEmail(string accountNumber, string clientId,
                                    string email, string mobilePhone)
        {
            EnsureInitialised();
            _withEmailWriter.WriteLine(
                $"{accountNumber}|{clientId}|{email}|{mobilePhone}|{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            _withEmailCount++;
        }

        public void RecordWithoutEmail(string accountNumber, string clientId, string mobilePhone)
        {
            EnsureInitialised();
            _withoutEmailWriter.WriteLine(
                $"{accountNumber}|{clientId}||{mobilePhone}|{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            _withoutEmailCount++;
        }

        public void Finalise()
        {
            _withEmailWriter?.Flush();
            _withEmailWriter?.Close();
            _withEmailWriter = null;

            _withoutEmailWriter?.Flush();
            _withoutEmailWriter?.Close();
            _withoutEmailWriter = null;

            _initialised = false;
        }

        public EmailTrackingSummary GetSummary() => new EmailTrackingSummary
        {
            WithEmailCount    = _withEmailCount,
            WithoutEmailCount = _withoutEmailCount,
            EmailsFilePath    = _emailsFilePath,
            NoEmailsFilePath  = _noEmailsFilePath,
        };

        private void EnsureInitialised()
        {
            if (!_initialised)
                throw new InvalidOperationException(
                    "Call Initialise() before recording email-tracking entries.");
        }
    }
}

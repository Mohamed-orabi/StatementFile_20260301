using System.Collections.Generic;

namespace StatementFile.Domain.Interfaces.Services
{
    /// <summary>
    /// Writes the two tracking text files produced after every HTML statement run:
    ///   1. {prefix}_Emails.txt       – accounts that have a valid email address.
    ///   2. {prefix}_WithoutEmails.txt – accounts that have no valid email address.
    ///
    /// Header format (pipe-delimited) matches the legacy clsStatHtml output:
    ///   AccountNumber | ClientID | Email | MobilePhone | DateTime
    ///
    /// These files are used by operations teams to reconcile e-statement delivery
    /// vs. physical mail delivery.
    /// </summary>
    public interface IEmailTrackingService
    {
        /// <summary>
        /// Initialises (creates/truncates) the two tracking files for a new run.
        /// Must be called once before any <see cref="Record"/> calls.
        /// </summary>
        /// <param name="outputDirectory">Directory where the files are written.</param>
        /// <param name="filePrefix">Prefix string, e.g. "UBA_CR_202601".</param>
        void Initialise(string outputDirectory, string filePrefix);

        /// <summary>Appends one row to the "with emails" tracking file.</summary>
        void RecordWithEmail(string accountNumber, string clientId, string email,
                             string mobilePhone);

        /// <summary>Appends one row to the "without emails" tracking file.</summary>
        void RecordWithoutEmail(string accountNumber, string clientId, string mobilePhone);

        /// <summary>Flushes and closes both tracking files.</summary>
        void Finalise();

        /// <summary>Returns summary statistics for the finished run.</summary>
        EmailTrackingSummary GetSummary();
    }

    public sealed class EmailTrackingSummary
    {
        public int WithEmailCount    { get; set; }
        public int WithoutEmailCount { get; set; }
        public int BadEmailCount     { get; set; }
        public string EmailsFilePath    { get; set; }
        public string NoEmailsFilePath  { get; set; }
    }
}

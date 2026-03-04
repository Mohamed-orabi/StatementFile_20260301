using System.Collections.Generic;

namespace StatementFile.Domain.Interfaces
{
    public interface IEmailService
    {
        /// <summary>
        /// Sends a statement email with optional file attachments.
        /// Mirrors <c>SendBankMail()</c> in frmStatementFile.
        /// </summary>
        void Send(EmailMessage message);
    }

    public sealed class EmailMessage
    {
        public string              FromAddress   { get; init; }
        public string              FromName      { get; init; }
        public IReadOnlyList<string> ToAddresses { get; init; }
        public string              Subject       { get; init; }
        public string              Body          { get; init; }
        public bool                IsBodyHtml    { get; init; }
        public IReadOnlyList<string> Attachments { get; init; }
    }
}

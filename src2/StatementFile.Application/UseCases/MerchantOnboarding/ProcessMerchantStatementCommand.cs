using System;

namespace StatementFile.Application.UseCases.MerchantOnboarding
{
    /// <summary>
    /// Command (input) for the Merchant Statement processing use case.
    /// Carries every piece of information required to fully process one
    /// merchant XML source file end-to-end.
    /// </summary>
    public sealed class ProcessMerchantStatementCommand
    {
        /// <summary>Absolute path to the incoming merchant XML file.</summary>
        public string   XmlSourceFilePath { get; }

        /// <summary>Human-readable bank name (e.g. "Network International").</summary>
        public string   BankFullName      { get; }

        /// <summary>Short bank code used in directory/file naming (e.g. "CPA", "BDK").</summary>
        public string   BankName          { get; }

        /// <summary>Numeric branch code stored in Oracle (e.g. 4).</summary>
        public string   BankCode          { get; }

        /// <summary>Processing date; drives directory naming and statement date fields.</summary>
        public DateTime ProcessingDate    { get; }

        public ProcessMerchantStatementCommand(
            string xmlSourceFilePath,
            string bankFullName,
            string bankName,
            string bankCode,
            DateTime processingDate)
        {
            if (string.IsNullOrWhiteSpace(xmlSourceFilePath))
                throw new ArgumentException("XML source file path is required.", nameof(xmlSourceFilePath));
            if (string.IsNullOrWhiteSpace(bankName))
                throw new ArgumentException("Bank name is required.", nameof(bankName));

            XmlSourceFilePath = xmlSourceFilePath;
            BankFullName      = bankFullName;
            BankName          = bankName;
            BankCode          = bankCode;
            ProcessingDate    = processingDate;
        }
    }
}

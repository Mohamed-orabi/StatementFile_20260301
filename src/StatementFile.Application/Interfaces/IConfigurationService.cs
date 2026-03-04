using System.Collections.Generic;
using StatementFile.Application.DTOs;

namespace StatementFile.Application.Interfaces
{
    /// <summary>
    /// Provides access to all runtime configuration: mail routing, output paths,
    /// SMTP settings, and SQL Server connection details.
    /// Implementations read from appsettings / JSON / App.config so the domain
    /// stays free of configuration coupling.
    /// </summary>
    public interface IConfigurationService
    {
        /// <summary>Returns the mail configuration for the given bank name.</summary>
        BankMailConfigDto GetBankMailConfig(string bankName, bool isProduction);

        /// <summary>Returns the root output path for generated statement files.</summary>
        string GetStatementOutputPath();

        /// <summary>Returns the SMTP server address.</summary>
        string GetSmtpServer();

        /// <summary>Returns the SQL Server connection string.</summary>
        string GetSqlConnectionString();

        /// <summary>Returns the default main schema prefix.</summary>
        string GetMainSchema();

        /// <summary>Returns the merchant statement MDB template path.</summary>
        string GetMerchantMdbTemplatePath();
    }

    public sealed class BankMailConfigDto
    {
        public string BankName  { get; set; }
        public string Message   { get; set; }
        public string OutputPath { get; set; }
        public List<EmailRecipientDto> ToRecipients  { get; set; } = new List<EmailRecipientDto>();
        public List<EmailRecipientDto> CcRecipients  { get; set; } = new List<EmailRecipientDto>();
    }
}

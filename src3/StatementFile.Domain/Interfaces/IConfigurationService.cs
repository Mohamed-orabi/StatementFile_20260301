namespace StatementFile.Domain.Interfaces
{
    /// <summary>
    /// Provides application-level settings that were previously read from
    /// App.config / Windows registry in the WinForms application.
    /// </summary>
    public interface IConfigurationService
    {
        string GetSqlConnectionString();
        string GetStatementOutputPath();
        string GetReportTemplatePath();
        string GetSmtpHost();
        int    GetSmtpPort();
        string GetSmtpUsername();
        string GetSmtpPassword();
        bool   GetSmtpUseSsl();
    }
}

using Microsoft.Extensions.Configuration;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Infrastructure.Configuration
{
    /// <summary>
    /// Reads application settings from appsettings.json / environment variables.
    /// Replaces direct App.config access from the WinForms layer.
    /// </summary>
    public sealed class AppConfigurationService : IConfigurationService
    {
        private readonly IConfiguration _config;

        public AppConfigurationService(IConfiguration config) => _config = config;

        public string GetSqlConnectionString() =>
            _config.GetConnectionString("SqlServer")
            ?? _config["SqlServer:ConnectionString"];

        public string GetStatementOutputPath() =>
            _config["StatementSettings:OutputPath"] ?? @"C:\StatementOutput";

        public string GetReportTemplatePath() =>
            _config["StatementSettings:ReportTemplatePath"] ?? @"C:\StatementReports";

        public string GetSmtpHost()     => _config["Smtp:Host"]     ?? "localhost";
        public int    GetSmtpPort()     => int.TryParse(_config["Smtp:Port"], out var p) ? p : 25;
        public string GetSmtpUsername() => _config["Smtp:Username"];
        public string GetSmtpPassword() => _config["Smtp:Password"];
        public bool   GetSmtpUseSsl()   => bool.TryParse(_config["Smtp:UseSsl"], out var s) && s;
    }
}

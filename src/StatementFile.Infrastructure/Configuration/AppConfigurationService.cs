using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using StatementFile.Application.DTOs;
using StatementFile.Application.Interfaces;

namespace StatementFile.Infrastructure.Configuration
{
    /// <summary>
    /// Reads all application configuration from:
    ///   - App.config  (connection strings, SMTP, file paths)
    ///   - mailConfiguration.json  (per-bank email routing)
    ///   - pathConfiguration.json  (output root paths)
    ///
    /// Implements <see cref="IConfigurationService"/> so the Application and Domain
    /// layers never reference System.Configuration or file I/O directly.
    ///
    /// Banks with special email-validation logic (QNB_ALAHLI, ICBG, HBLN) are
    /// handled by the same filter logic from the legacy ConfigurationReader —
    /// now encapsulated and testable here.
    /// </summary>
    public sealed class AppConfigurationService : IConfigurationService
    {
        // ── Lazy-loaded config stores ──────────────────────────────────────────────
        private readonly Lazy<List<MailConfigEntry>>  _mailConfigs;
        private readonly Lazy<PathConfigEntry>        _pathConfig;

        // ── Banks that apply the valid-flag filter on To recipients ───────────────
        private static readonly HashSet<string> ValidatedEmailBanks =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "QNB_ALAHLI", "ICBG", "HBLN" };

        public AppConfigurationService()
        {
            _mailConfigs = new Lazy<List<MailConfigEntry>>(LoadMailConfigs);
            _pathConfig  = new Lazy<PathConfigEntry>(LoadPathConfig);
        }

        // ── IConfigurationService ─────────────────────────────────────────────────

        public BankMailConfigDto GetBankMailConfig(string bankName, bool isProduction)
        {
            var entry = _mailConfigs.Value
                .FirstOrDefault(e => string.Equals(e.name, bankName,
                                     StringComparison.OrdinalIgnoreCase));

            if (entry == null) return null;

            // Apply valid-flag filter for designated banks (database-driven extension point)
            List<MailRecipientEntry> toList = entry.to ?? new List<MailRecipientEntry>();
            if (ValidatedEmailBanks.Contains(bankName))
            {
                toList = toList
                    .Where(v => isProduction ? v.valid != false : v.valid != true)
                    .ToList();
            }

            var dto = new BankMailConfigDto
            {
                BankName   = bankName,
                Message    = entry.message,
                OutputPath = entry.path,
            };

            foreach (var r in toList)
                dto.ToRecipients.Add(new EmailRecipientDto
                    { Email = r.email, Name = r.name, Valid = r.valid, RecipientType = "To" });

            foreach (var r in entry.cc ?? new List<MailCcEntry>())
                dto.CcRecipients.Add(new EmailRecipientDto
                    { Email = r.email, Name = r.name, RecipientType = "CC" });

            return dto;
        }

        public string GetStatementOutputPath()
        {
            string configured = _pathConfig.Value?.stmtPath;
            return string.IsNullOrWhiteSpace(configured)
                ? AppDomain.CurrentDomain.BaseDirectory
                : configured;
        }

        public string GetSmtpServer()
            => ConfigurationManager.AppSettings["SmtpServer"] ?? "localhost";

        public string GetOracleConnectionString()
        {
            // Mirrors the legacy clsDbCon.sConOracle pattern
            string server = ConfigurationManager.AppSettings["ServerName"] ?? string.Empty;
            string schema = ConfigurationManager.AppSettings["MainSchema"]  ?? string.Empty;
            return $"Data Source={server};User Id={schema};Password=;";
        }

        public string GetMainSchema()
            => ConfigurationManager.AppSettings["MainSchema"] ?? string.Empty;

        public string GetMerchantMdbTemplatePath()
            => ConfigurationManager.AppSettings["MerchantMdbTemplatePath"]
               ?? @"D:\pC#\ProjData\Statement\_Data\MerchantStatementTemplate.mdb";

        // ── Private Helpers ────────────────────────────────────────────────────────

        private static List<MailConfigEntry> LoadMailConfigs()
        {
            string relative = ConfigurationManager.AppSettings["mailConfigFilePath"] ?? string.Empty;
            string fullPath = Path.Combine(Directory.GetCurrentDirectory(), relative.TrimStart('\\'));

            if (!File.Exists(fullPath))
                throw new FileNotFoundException(
                    $"Mail configuration file not found: {fullPath}");

            string json = File.ReadAllText(fullPath);
            return JsonConvert.DeserializeObject<List<MailConfigEntry>>(json)
                   ?? new List<MailConfigEntry>();
        }

        private static PathConfigEntry LoadPathConfig()
        {
            string relative = ConfigurationManager.AppSettings["pathConfiguration"] ?? string.Empty;
            string fullPath = Path.Combine(Directory.GetCurrentDirectory(), relative.TrimStart('\\'));

            if (!File.Exists(fullPath))
                return new PathConfigEntry();

            string json = File.ReadAllText(fullPath);
            return JsonConvert.DeserializeObject<PathConfigEntry>(json) ?? new PathConfigEntry();
        }

        // ── JSON schema POCOs (internal to Infrastructure) ────────────────────────

        private sealed class MailConfigEntry
        {
            public string                  name    { get; set; }
            public string                  message { get; set; }
            public List<MailRecipientEntry> to      { get; set; }
            public List<MailCcEntry>        cc      { get; set; }
            public string                  path    { get; set; }
        }

        private sealed class MailRecipientEntry
        {
            public string email { get; set; }
            public string name  { get; set; }
            public bool?  valid { get; set; }
        }

        private sealed class MailCcEntry
        {
            public string email { get; set; }
            public string name  { get; set; }
        }

        private sealed class PathConfigEntry
        {
            public string stmtPath { get; set; }
        }
    }
}

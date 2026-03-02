using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using StatementFile.Application.Interfaces;
using StatementFile.Domain.Common;
using StatementFile.Domain.Interfaces.Repositories;
using StatementFile.Domain.Interfaces.Services;

namespace StatementFile.Application.UseCases.MerchantOnboarding
{
    /// <summary>
    /// Orchestrates the Merchant Statement use case:
    ///   1. Validate &amp; prepare output directory
    ///   2. Move XML to output directory, copy MDB template
    ///   3. Parse XML → DataSet → insert into MDB
    ///   4. Run MDB post-fixup queries
    ///   5. Export Crystal Report (PDF)
    ///   6. Generate HTML e-statement if applicable
    ///   7. Send email to bank
    ///
    /// Business rules remain identical to the legacy clsStatementMrchXml.Statement().
    /// All I/O and external calls go through injected abstractions.
    /// </summary>
    public sealed class ProcessMerchantStatementHandler
    {
        private readonly IMerchantStatementRepository _merchantRepo;
        private readonly IReportService               _reportService;
        private readonly IEmailService                _emailService;
        private readonly IConfigurationService        _config;

        public ProcessMerchantStatementHandler(
            IMerchantStatementRepository merchantRepo,
            IReportService               reportService,
            IEmailService                emailService,
            IConfigurationService        config)
        {
            _merchantRepo  = merchantRepo  ?? throw new ArgumentNullException(nameof(merchantRepo));
            _reportService = reportService ?? throw new ArgumentNullException(nameof(reportService));
            _emailService  = emailService  ?? throw new ArgumentNullException(nameof(emailService));
            _config        = config        ?? throw new ArgumentNullException(nameof(config));
        }

        public Result Handle(ProcessMerchantStatementCommand command)
        {
            try
            {
                // ── Step 1: Prepare output directory ──────────────────────────────
                string timestamp       = command.ProcessingDate.ToString("yyyyMMdd_HHmmss");
                string sourceDirectory = Path.GetDirectoryName(command.XmlSourceFilePath);
                string outputDirectory = Path.Combine(
                    sourceDirectory,
                    $"{command.BankName}_MerchantStatement_{timestamp}");

                Directory.CreateDirectory(outputDirectory);

                // ── Step 2: Move XML; copy MDB template ───────────────────────────
                string xmlDestPath = Path.Combine(
                    outputDirectory,
                    $"{command.BankName}_MerchantStatement_{timestamp}.XML");

                File.Move(command.XmlSourceFilePath, xmlDestPath);

                string mdbTemplatePath = _config.GetMerchantMdbTemplatePath();
                string mdbDestPath     = Path.ChangeExtension(xmlDestPath, ".mdb");
                File.Copy(mdbTemplatePath, mdbDestPath, overwrite: true);

                // ── Step 3: Parse XML → DataSet ────────────────────────────────────
                DataSet dsXml = new DataSet();
                dsXml.ReadXml(xmlDestPath);

                bool hasOperations = dsXml.Tables.Contains("Statement")
                                  && dsXml.Tables.Contains("Operation")
                                  && dsXml.Tables.Count > 1;

                if (hasOperations)
                {
                    dsXml.Relations.Add(
                        "StaementNoDR",
                        dsXml.Tables["Statement"].Columns["StatementNo"],
                        dsXml.Tables["Operation"].Columns["StatementNo"]);
                }

                // ── Step 4: Persist to MDB ─────────────────────────────────────────
                var statements = MerchantStatementXmlMapper.MapFromDataSet(
                    dsXml, command.BankCode, command.BankName, command.BankFullName,
                    command.ProcessingDate);

                _merchantRepo.SaveMerchantStatements(statements, mdbDestPath);
                _merchantRepo.ApplyPostInsertFixups(mdbDestPath);

                // ── Step 5: Reload clean data from MDB for report ─────────────────
                var cleanStatements = _merchantRepo.LoadFromMdb(mdbDestPath);

                // ── Step 6: Export Crystal Reports PDF ────────────────────────────
                DataSet dsReport = BuildReportDataSet(cleanStatements);
                string  pdfPath  = Path.ChangeExtension(mdbDestPath, ".pdf");

                _reportService.Export(
                    reportTemplatePath: GetReportTemplate(command.BankName),
                    data:              dsReport,
                    outputPath:        pdfPath,
                    format:            "PDF");

                // ── Step 7: Send email to bank ────────────────────────────────────
                var mailConfig = _config.GetBankMailConfig(command.BankName, isProduction: true);
                if (mailConfig != null)
                {
                    SendMerchantEmail(command, mailConfig, pdfPath);
                }

                return Result.Ok();
            }
            catch (Exception ex)
            {
                return Result.Fail($"Merchant statement processing failed: {ex.Message}");
            }
        }

        // ── Private Helpers ────────────────────────────────────────────────────────

        private void SendMerchantEmail(
            ProcessMerchantStatementCommand command,
            Interfaces.BankMailConfigDto    mailConfig,
            string                          pdfAttachment)
        {
            var toAddresses  = new List<string>();
            var ccAddresses  = new List<string>();
            var bccAddresses = new List<string> { "Statement@emp-group.com" };

            // Resolve bank-specific To address (legacy routing preserved exactly)
            string bankSpecificTo = ResolveBankSpecificTo(command.BankName);
            if (!string.IsNullOrEmpty(bankSpecificTo))
                toAddresses.Add(bankSpecificTo);
            else
                toAddresses.Add("merchantsupport@emp-group.com");

            foreach (var r in mailConfig.CcRecipients)
                ccAddresses.Add(r.Email);

            _emailService.SendHtml(
                fromAddress:     "cardservices@emp-group.com",
                fromDisplayName: $"{command.BankFullName} - Statement",
                toAddresses:     toAddresses,
                ccAddresses:     ccAddresses,
                bccAddresses:    bccAddresses,
                subject:         $"Merchant Statement - {command.BankFullName} - {command.ProcessingDate:MMMM yyyy}",
                htmlBody:        mailConfig.Message,
                attachmentPaths: new[] { pdfAttachment });
        }

        /// <summary>
        /// Mirrors the legacy bank-specific routing logic from clsStatementMrchXml.sendEmail2Bank().
        /// Kept in one place; the routing table is database-configurable via appSettings
        /// extension (see Infrastructure.Configuration.MerchantEmailRoutingConfig).
        /// </summary>
        private static string ResolveBankSpecificTo(string bankName)
        {
            switch (bankName.ToUpperInvariant())
            {
                case "SSB":  return "Abena.A.Asare-Menako@socgen.com";
                case "GTBG": return "sheikh.chaw@gtbank.com";
                case "ICBG": return "Ackah-Nyamike.Juliet@ghana.accessbankplc.com";
                default:     return string.Empty; // falls back to merchantsupport@emp-group.com
            }
        }

        private static string GetReportTemplate(string bankName)
        {
            // Template resolved by convention; can be externalised to a config table.
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string bankTemplate = Path.Combine(basePath, "Reports", $"MerchantStatement_{bankName}.rpt");
            if (File.Exists(bankTemplate)) return bankTemplate;
            return Path.Combine(basePath, "Reports", "MerchantStatement.rpt");
        }

        private static DataSet BuildReportDataSet(IEnumerable<Domain.Entities.MerchantStatement> statements)
        {
            var ds = new DataSet("MerchantReport");

            var statTable = ds.Tables.Add("Statement");
            statTable.Columns.Add("StatementNo");
            statTable.Columns.Add("Account");
            statTable.Columns.Add("BankName");
            statTable.Columns.Add("StatDate", typeof(DateTime));

            var opTable = ds.Tables.Add("Operation");
            opTable.Columns.Add("StatementNo");
            opTable.Columns.Add("D");
            opTable.Columns.Add("A",  typeof(decimal));
            opTable.Columns.Add("CF", typeof(decimal));
            opTable.Columns.Add("OD", typeof(DateTime));

            foreach (var s in statements)
            {
                statTable.Rows.Add(s.StatementNo, s.Account, s.BankName, s.StatDate);
                foreach (var op in s.Operations)
                    opTable.Rows.Add(s.StatementNo, op.D, op.A, op.CF, op.OD);
            }

            return ds;
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using StatementFile.Application.Interfaces;
using StatementFile.Domain.Common;
using StatementFile.Domain.Interfaces;
using StatementFile.Domain.Interfaces.Services;

namespace StatementFile.Application.UseCases.StatementGeneration
{
    /// <summary>
    /// Orchestrates a single bank/product statement generation run.
    ///
    /// Flow (preserving legacy behaviour from clsBasStatementHtml.Statement()):
    ///   1. Create output directory (YYYYMMDD_BANKCODE convention)
    ///   2. Load master + detail DataSets from Oracle
    ///   3. Load email list DataSet
    ///   4. Resolve formatter via factory (polymorphic, database-key driven)
    ///   5. Format output files (HTML/TXT/PDF)
    ///   6. Export report (Crystal Reports)
    ///   7. Package (ZIP + MD5) &amp; deliver (FTP + Email)
    /// </summary>
    public sealed class GenerateStatementHandler
    {
        private readonly IUnitOfWork                  _unitOfWork;
        private readonly IStatementFormatterFactory   _formatterFactory;
        private readonly IReportService               _reportService;
        private readonly IEmailService                _emailService;
        private readonly IFtpService                  _ftpService;
        private readonly IConfigurationService        _config;

        public GenerateStatementHandler(
            IUnitOfWork                unitOfWork,
            IStatementFormatterFactory formatterFactory,
            IReportService             reportService,
            IEmailService              emailService,
            IFtpService                ftpService,
            IConfigurationService      config)
        {
            _unitOfWork       = unitOfWork       ?? throw new ArgumentNullException(nameof(unitOfWork));
            _formatterFactory = formatterFactory ?? throw new ArgumentNullException(nameof(formatterFactory));
            _reportService    = reportService    ?? throw new ArgumentNullException(nameof(reportService));
            _emailService     = emailService     ?? throw new ArgumentNullException(nameof(emailService));
            _ftpService       = ftpService       ?? throw new ArgumentNullException(nameof(ftpService));
            _config           = config           ?? throw new ArgumentNullException(nameof(config));
        }

        public Result<GenerateStatementResult> Handle(GenerateStatementCommand command)
        {
            try
            {
                // ── Step 1: Output directory ───────────────────────────────────────
                string dirName    = $"{command.StatementDate:yyyyMMdd}{command.BankName}";
                string outputDir  = Path.Combine(command.OutputRootPath, dirName);
                Directory.CreateDirectory(outputDir);

                // ── Step 2: Load data from Oracle ──────────────────────────────────
                string orderBy   = "m.cardproduct,m.CARDBRANCHPART,m.accountno,m.cardprimary desc,m.cardno";
                var masterDs     = _unitOfWork.Statements.LoadMasterDataSet(command.BranchCode, orderBy);
                var detailDs     = _unitOfWork.Statements.LoadDetailDataSet(command.BranchCode);
                var emailDs      = _unitOfWork.Statements.LoadEmailDataSet(command.BranchCode);

                if (masterDs == null || masterDs.Tables.Count == 0)
                    return Result.Fail<GenerateStatementResult>(
                        $"No statement data found for branch {command.BranchCode}, product {command.CardProduct}.");

                // ── Step 3: Resolve formatter (polymorphic) ────────────────────────
                var formatter    = _formatterFactory.GetFormatter(command.FormatterKey);

                // ── Step 4: Format output ──────────────────────────────────────────
                var outputFiles  = new List<string>(
                    formatter.Format(masterDs, outputDir, command.BranchCode, command.CardProduct));

                // ── Step 5: Export report ──────────────────────────────────────────
                string reportTemplate = GetReportTemplate(command.BankName, command.CardProduct);
                if (File.Exists(reportTemplate))
                {
                    string pdfPath = Path.Combine(outputDir,
                        $"{command.BankName}_{command.CardProduct}_{command.StatementDate:yyyyMMdd}.pdf");
                    _reportService.Export(reportTemplate, masterDs, pdfPath, "PDF");
                    outputFiles.Add(pdfPath);
                }

                // ── Step 6: Package & deliver ──────────────────────────────────────
                var mailConfig = _config.GetBankMailConfig(command.BankName, isProduction: true);
                if (mailConfig != null && outputFiles.Count > 0)
                {
                    _ftpService.SendFile(outputFiles[0], command.BankName);
                    DeliverByEmail(command, mailConfig, outputFiles);
                }

                return Result.Ok(new GenerateStatementResult
                {
                    BranchCode    = command.BranchCode,
                    CardProduct   = command.CardProduct,
                    OutputDirectory = outputDir,
                    FilesGenerated  = outputFiles.Count,
                });
            }
            catch (Exception ex)
            {
                return Result.Fail<GenerateStatementResult>(
                    $"Statement generation failed for {command.BankName}/{command.CardProduct}: {ex.Message}");
            }
        }

        // ── Private Helpers ────────────────────────────────────────────────────────

        private void DeliverByEmail(
            GenerateStatementCommand command,
            Interfaces.BankMailConfigDto mailConfig,
            IEnumerable<string> attachments)
        {
            var toList  = new List<string>();
            var ccList  = new List<string>();

            foreach (var r in mailConfig.ToRecipients)
                toList.Add(r.Email);
            foreach (var r in mailConfig.CcRecipients)
                ccList.Add(r.Email);

            _emailService.SendHtml(
                fromAddress:     "cardservices@emp-group.com",
                fromDisplayName: $"{command.BankFullName} - Statement",
                toAddresses:     toList,
                ccAddresses:     ccList,
                bccAddresses:    new[] { "Statement@emp-group.com" },
                subject:         $"Statement - {command.BankFullName} - {command.StatementDate:MMMM yyyy}",
                htmlBody:        mailConfig.Message,
                attachmentPaths: attachments);
        }

        private static string GetReportTemplate(string bankName, string cardProduct)
        {
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            // Try bank+product specific, then bank only, then generic
            string[] candidates =
            {
                Path.Combine(basePath, "Reports", $"Bank_{bankName}_{cardProduct}.rpt"),
                Path.Combine(basePath, "Reports", $"Bank_{bankName}.rpt"),
                Path.Combine(basePath, "Reports", "rsStatement.rdlc"),
            };
            foreach (string c in candidates)
                if (File.Exists(c)) return c;
            return string.Empty;
        }
    }

    public sealed class GenerateStatementResult
    {
        public int    BranchCode       { get; set; }
        public string CardProduct      { get; set; }
        public string OutputDirectory  { get; set; }
        public int    FilesGenerated   { get; set; }
    }
}

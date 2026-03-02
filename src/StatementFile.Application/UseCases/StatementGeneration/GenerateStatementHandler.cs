using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using StatementFile.Application.Interfaces;
using StatementFile.Domain.Common;
using StatementFile.Domain.Enums;
using StatementFile.Domain.Interfaces;
using StatementFile.Domain.Interfaces.Services;

namespace StatementFile.Application.UseCases.StatementGeneration
{
    /// <summary>
    /// Orchestrates a single bank/product statement generation run for ALL output types.
    ///
    /// Complete flow (preserving every behaviour found in the Banks/ and BasStat/ folders):
    ///
    ///   Step 1  – Prepare / clean output directory
    ///   Step 2  – Load the appropriate Oracle DataSet variant (standard / sorted /
    ///             with-overdue-days / exclude-VISA / VIP-only)
    ///   Step 3  – Load supplementary DataSets (installments, rewards, emails,
    ///             identity data, products)
    ///   Step 4  – Resolve formatter via registry key (HTML / RawData / TextLabel /
    ///             Text / XML / PDF)
    ///   Step 5  – Run formatter → produces output file list
    ///   Step 6  – For RawData and XML runs: package ZIP + MD5
    ///   Step 7  – For HTML/PDF runs: send per-card emails (tracked in Email/NoEmail files)
    ///   Step 8  – Deliver completed package to bank via FTP + notification email
    /// </summary>
    public sealed class GenerateStatementHandler
    {
        private readonly IUnitOfWork                _unitOfWork;
        private readonly IStatementFormatterFactory _formatterFactory;
        private readonly IStatementQueryService     _queryService;
        private readonly IReportService             _reportService;
        private readonly IEmailService              _emailService;
        private readonly IFtpService                _ftpService;
        private readonly IEmailTrackingService      _emailTracking;
        private readonly IFilePackagingService      _packagingService;
        private readonly IDataMaintenanceService    _maintenanceService;
        private readonly IConfigurationService      _config;

        public GenerateStatementHandler(
            IUnitOfWork                unitOfWork,
            IStatementFormatterFactory formatterFactory,
            IStatementQueryService     queryService,
            IReportService             reportService,
            IEmailService              emailService,
            IFtpService                ftpService,
            IEmailTrackingService      emailTracking,
            IFilePackagingService      packagingService,
            IDataMaintenanceService    maintenanceService,
            IConfigurationService      config)
        {
            _unitOfWork         = unitOfWork         ?? throw new ArgumentNullException(nameof(unitOfWork));
            _formatterFactory   = formatterFactory   ?? throw new ArgumentNullException(nameof(formatterFactory));
            _queryService       = queryService       ?? throw new ArgumentNullException(nameof(queryService));
            _reportService      = reportService      ?? throw new ArgumentNullException(nameof(reportService));
            _emailService       = emailService       ?? throw new ArgumentNullException(nameof(emailService));
            _ftpService         = ftpService         ?? throw new ArgumentNullException(nameof(ftpService));
            _emailTracking      = emailTracking      ?? throw new ArgumentNullException(nameof(emailTracking));
            _packagingService   = packagingService   ?? throw new ArgumentNullException(nameof(packagingService));
            _maintenanceService = maintenanceService ?? throw new ArgumentNullException(nameof(maintenanceService));
            _config             = config             ?? throw new ArgumentNullException(nameof(config));
        }

        public Result<GenerateStatementResult> Handle(GenerateStatementCommand command)
        {
            try
            {
                // ── Step 1: Prepare output directory ──────────────────────────────
                string outputDir = BuildOutputDirectory(command);
                if (!command.AppendMode)
                {
                    if (Directory.Exists(outputDir))
                        Directory.Delete(outputDir, recursive: true);
                }
                Directory.CreateDirectory(outputDir);

                // ── Step 2: Run pre-processing maintenance ─────────────────────────
                RunPreProcessing(command);

                // ── Step 3: Load the appropriate master DataSet ────────────────────
                DataSet masterDs = LoadMasterDataSet(command);
                if (masterDs == null || masterDs.Tables.Count == 0 ||
                    masterDs.Tables[0].Rows.Count == 0)
                {
                    return Result.Fail<GenerateStatementResult>(
                        $"No data found for branch {command.BranchCode} / product {command.CardProduct}.");
                }

                // ── Step 4: Load supplementary DataSets ────────────────────────────
                StatementDataContext ctx = LoadSupplementaryData(command, masterDs);

                // ── Step 5: Resolve formatter and generate output ──────────────────
                var formatter    = _formatterFactory.GetFormatter(command.FormatterKey);
                var outputFiles  = new List<string>(
                    formatter.Format(ctx, outputDir, command));

                // ── Step 6: Package ZIP + MD5 (RawData and XML runs) ──────────────
                if (command.OutputType == StatementOutputType.RawData ||
                    command.OutputType == StatementOutputType.Xml)
                {
                    string zipPath = Path.Combine(outputDir,
                        $"{command.BankName}_{command.StatementTypeSuffix}_{command.StatementDate:yyyyMM}.zip");
                    string zipped = _packagingService.PackageFiles(outputFiles, zipPath);
                    outputFiles.Add(zipped);
                }

                // ── Step 7: Initialise email tracking (HTML runs) ─────────────────
                if (command.OutputType == StatementOutputType.Html)
                {
                    string prefix = $"{command.BankName}_{command.StatementTypeSuffix}_{command.StatementDate:yyyyMM}";
                    _emailTracking.Initialise(outputDir, prefix);
                    // Tracking rows are written inside the formatter via IEmailTrackingService;
                    // the formatter receives it through the context passed by the factory.
                    _emailTracking.Finalise();
                }

                // ── Step 8: FTP + notification email delivery ──────────────────────
                var mailConfig = _config.GetBankMailConfig(command.BankName, isProduction: true);
                if (mailConfig != null && outputFiles.Count > 0)
                {
                    // FTP the primary output (first file or zip)
                    _ftpService.SendFile(outputFiles[outputFiles.Count - 1], command.BankName);

                    // Send notification email to bank operations team
                    SendNotificationEmail(command, mailConfig, outputFiles);
                }

                return Result.Ok(new GenerateStatementResult
                {
                    BranchCode      = command.BranchCode,
                    CardProduct     = command.CardProduct,
                    OutputDirectory = outputDir,
                    FilesGenerated  = outputFiles.Count,
                    EmailsSent      = _emailTracking.GetSummary().WithEmailCount,
                    NoEmailCount    = _emailTracking.GetSummary().WithoutEmailCount,
                });
            }
            catch (Exception ex)
            {
                return Result.Fail<GenerateStatementResult>(
                    $"Statement generation failed for {command.BankName}/{command.CardProduct}: {ex.Message}");
            }
        }

        // ── Pre-processing ────────────────────────────────────────────────────────

        private void RunPreProcessing(GenerateStatementCommand command)
        {
            // Delete on-hold rows (HOLSTMT = 'Y') if required
            if (command.HasMode(ProcessingMode.DeleteOnHold))
                _maintenanceService.CleanNullCards(command.BranchCode,
                    excludeReward:      command.HasMode(ProcessingMode.Reward),
                    excludeInstallment: command.HasMode(ProcessingMode.Installment),
                    installmentCondition: command.InstallmentContractCondition);

            // Fix Arabic address fields if required
            if (command.HasMode(ProcessingMode.FixArabicAddress))
                _maintenanceService.FixArabicAddress(command.BranchCode);
        }

        // ── Oracle DataSet loading ────────────────────────────────────────────────

        private DataSet LoadMasterDataSet(GenerateStatementCommand command)
        {
            if (command.HasMode(ProcessingMode.Vip))
                return _queryService.LoadVipOnly(command.BranchCode);

            if (command.HasMode(ProcessingMode.WithOverdueDays))
                return _queryService.LoadWithOverdueDays(command.BranchCode);

            if (command.HasMode(ProcessingMode.ExcludeVisa))
                return _queryService.LoadExcludingVisa(command.BranchCode);

            if (command.HasMode(ProcessingMode.SortCardPriority))
                return _queryService.LoadSortedByCardPriority(command.BranchCode);

            return _queryService.LoadStandard(command.BranchCode);
        }

        private StatementDataContext LoadSupplementaryData(
            GenerateStatementCommand command,
            DataSet                  masterDs)
        {
            var ctx = new StatementDataContext { MasterDataSet = masterDs };

            if (command.HasMode(ProcessingMode.Installment))
                ctx.InstallmentDataSet = _queryService.LoadInstallments(command.BranchCode);

            if (command.HasMode(ProcessingMode.Reward))
                ctx.RewardDataSet = _queryService.LoadRewards(
                    command.BranchCode, command.RewardContractCondition);

            ctx.EmailDataSet = _queryService.LoadClientEmails(command.BranchCode);

            if (command.HasMode(ProcessingMode.ExternalId))
                ctx.IdentityDataSet = _queryService.LoadClientIdentity(command.BranchCode);

            ctx.ProductDataSet = _queryService.LoadProducts(command.BranchCode);

            return ctx;
        }

        // ── Email notification ────────────────────────────────────────────────────

        private void SendNotificationEmail(
            GenerateStatementCommand command,
            BankMailConfigDto        mailConfig,
            IEnumerable<string>      attachments)
        {
            var toList  = new List<string>();
            var ccList  = new List<string>();
            var bccList = new List<string> { "Statement@emp-group.com" };

            foreach (var r in mailConfig.ToRecipients)
                toList.Add(r.Email);
            foreach (var r in mailConfig.CcRecipients)
                ccList.Add(r.Email);

            _emailService.SendHtml(
                fromAddress:     command.EmailFromAddress,
                fromDisplayName: $"{command.BankFullName} - Statement",
                toAddresses:     toList,
                ccAddresses:     ccList,
                bccAddresses:    bccList,
                subject:         $"Statement - {command.BankFullName} - {command.StatementDate:MMMM yyyy}",
                htmlBody:        mailConfig.Message,
                attachmentPaths: attachments);
        }

        // ── Output directory convention ───────────────────────────────────────────

        /// <summary>
        /// Builds the output directory path following the legacy convention:
        ///   {rootPath}{yyyyMM}{BankName}_{StatementTypeSuffix}
        /// e.g. D:\Statements\202601UBA_CR
        /// </summary>
        private static string BuildOutputDirectory(GenerateStatementCommand command)
        {
            string dirName = command.StatementDate.ToString("yyyyMM")
                           + command.BankName
                           + "_" + command.StatementTypeSuffix;
            return Path.Combine(command.OutputRootPath, dirName);
        }
    }

    // ── Result and context ─────────────────────────────────────────────────────────

    public sealed class GenerateStatementResult
    {
        public int    BranchCode      { get; set; }
        public string CardProduct     { get; set; }
        public string OutputDirectory { get; set; }
        public int    FilesGenerated  { get; set; }
        public int    EmailsSent      { get; set; }
        public int    NoEmailCount    { get; set; }
    }

    /// <summary>
    /// All DataSets loaded from Oracle for a single generation run.
    /// Passed to the formatter so it can access supplementary data
    /// without re-querying the database.
    /// </summary>
    public sealed class StatementDataContext
    {
        public DataSet MasterDataSet      { get; set; }
        public DataSet InstallmentDataSet { get; set; }
        public DataSet RewardDataSet      { get; set; }
        public DataSet EmailDataSet       { get; set; }
        public DataSet IdentityDataSet    { get; set; }
        public DataSet ProductDataSet     { get; set; }
    }
}

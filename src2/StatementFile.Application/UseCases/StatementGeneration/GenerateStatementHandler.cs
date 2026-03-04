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
    ///   Step 2  – Run pre-processing maintenance (on-hold delete, NULL-card delete,
    ///             card-branch alignment, Arabic address fix, address split, lang code)
    ///   Step 3  – Load the appropriate Oracle DataSet variant (standard / sorted /
    ///             with-overdue-days / exclude-VISA / VIP-only / markup-fee-removal)
    ///   Step 4  – Load supplementary DataSets (installments, rewards, emails,
    ///             identity data, passport+birthyear, products, branch parts, client bank)
    ///   Step 5  – Resolve formatter via registry key (HTML / RawData / TextLabel /
    ///             Text / XML / PDF)
    ///   Step 6  – Run formatter → produces output file list
    ///   Step 7  – For RawData and XML runs: package ZIP + MD5
    ///   Step 8  – For HTML/PDF runs: send per-card emails (tracked in Email/NoEmail files)
    ///   Step 9  – Insert statement summary record in a4m.MSCC_PROD_STAT_MASTER
    ///   Step 10 – Deliver completed package to bank via FTP + notification email
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
        private readonly IStatementSummaryService   _summaryService;
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
            IStatementSummaryService   summaryService,
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
            _summaryService     = summaryService     ?? throw new ArgumentNullException(nameof(summaryService));
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

                int masterRowCount = masterDs.Tables[0].Rows.Count;

                // ── Step 4: Load supplementary DataSets ────────────────────────────
                StatementDataContext ctx = LoadSupplementaryData(command, masterDs);

                // ── Step 5: Resolve formatter and generate output ──────────────────
                var formatter   = _formatterFactory.GetFormatter(command.FormatterKey);
                var outputFiles = new List<string>(formatter.Format(ctx, outputDir, command));

                int transactionCount = ctx.DetailDataSet != null && ctx.DetailDataSet.Tables.Count > 0
                    ? ctx.DetailDataSet.Tables[0].Rows.Count
                    : 0;

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
                    _emailTracking.Finalise();
                }

                // ── Step 8: Insert statement summary in a4m.MSCC_PROD_STAT_MASTER ─
                // Matches clsStatementSummary.InsertRecordDb() called after generation.
                try
                {
                    _summaryService.InsertRecord(
                        bankCode:         command.BranchCode,
                        productCode:      0,
                        productName:      command.CardProduct,
                        statementDate:    command.StatementDate,
                        noOfStatements:   masterRowCount,
                        noOfTransactions: transactionCount,
                        creationDate:     DateTime.Now);
                }
                catch
                {
                    // Summary insert failure must not abort the generation run
                }

                // ── Step 9: FTP + notification email delivery ──────────────────────
                var mailConfig = _config.GetBankMailConfig(command.BankName, isProduction: true);
                if (mailConfig != null && outputFiles.Count > 0)
                {
                    _ftpService.SendFile(outputFiles[outputFiles.Count - 1], command.BankName);
                    SendNotificationEmail(command, mailConfig, outputFiles);
                }

                return Result.Ok(new GenerateStatementResult
                {
                    BranchCode       = command.BranchCode,
                    CardProduct      = command.CardProduct,
                    OutputDirectory  = outputDir,
                    FilesGenerated   = outputFiles.Count,
                    EmailsSent       = _emailTracking.GetSummary().WithEmailCount,
                    NoEmailCount     = _emailTracking.GetSummary().WithoutEmailCount,
                    StatementsCount  = masterRowCount,
                    TransactionCount = transactionCount,
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
            // Step 1: Delete on-hold rows (POSTINGDATE IS NULL AND DOCNO IS NULL)
            if (command.HasMode(ProcessingMode.DeleteOnHold))
                _maintenanceService.DeleteOnHoldTransactions(
                    command.BranchCode,
                    isReward: command.HasMode(ProcessingMode.Reward));

            // Step 2: Merge Mark-Up Fee transactions by docno
            if (command.HasMode(ProcessingMode.MergeMarkUpFees))
                _maintenanceService.MergeMarkUpFees(command.BranchCode);

            // Step 3: NULL-card row delete
            _maintenanceService.CleanNullCards(command.BranchCode,
                excludeReward:      command.HasMode(ProcessingMode.Reward),
                excludeInstallment: command.HasMode(ProcessingMode.Installment),
                installmentCondition: command.InstallmentContractCondition);

            // Step 4: Card-branch-part alignment
            _maintenanceService.MatchCardBranchForAccount(command.BranchCode);

            // Step 5: Arabic address corruption fix
            if (command.HasMode(ProcessingMode.FixArabicAddress))
                _maintenanceService.FixArabicAddress(command.BranchCode);

            // Step 6: Long address split
            if (command.HasMode(ProcessingMode.FixAddress))
                _maintenanceService.FixAddress(command.BranchCode);

            // Step 7: Arabic language code assignment
            if (command.HasMode(ProcessingMode.FixArabicAddressLang))
                _maintenanceService.FixArabicAddressLang(command.BranchCode);
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

            if (command.HasMode(ProcessingMode.RemoveMarkupFee))
                return _queryService.LoadWithMarkupFeeRemoval(command.BranchCode);

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

            if (command.HasMode(ProcessingMode.PasNoAndBirthYear))
                ctx.PasNoAndBirthYearDataSet = _queryService.LoadClientPasNoAndBirthYear(command.BranchCode);

            ctx.ProductDataSet    = _queryService.LoadProducts(command.BranchCode);
            ctx.BranchPartDataSet = _queryService.LoadBranchPart(command.BranchCode);

            if (command.HasMode(ProcessingMode.ClientBank))
                ctx.ClientBankDataSet = _queryService.LoadClientBank(command.BranchCode);

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
        public int    BranchCode       { get; set; }
        public string CardProduct      { get; set; }
        public string OutputDirectory  { get; set; }
        public int    FilesGenerated   { get; set; }
        public int    EmailsSent       { get; set; }
        public int    NoEmailCount     { get; set; }
        public int    StatementsCount  { get; set; }
        public int    TransactionCount { get; set; }
    }

    /// <summary>
    /// All DataSets loaded from Oracle for a single generation run.
    /// Passed to the formatter so it can access supplementary data
    /// without re-querying the database.
    /// </summary>
    public sealed class StatementDataContext
    {
        public DataSet MasterDataSet            { get; set; }
        public DataSet DetailDataSet            { get; set; }
        public DataSet InstallmentDataSet       { get; set; }
        public DataSet RewardDataSet            { get; set; }
        public DataSet EmailDataSet             { get; set; }
        public DataSet IdentityDataSet          { get; set; }
        public DataSet PasNoAndBirthYearDataSet { get; set; }
        public DataSet ProductDataSet           { get; set; }
        public DataSet BranchPartDataSet        { get; set; }
        public DataSet ClientBankDataSet        { get; set; }
    }
}

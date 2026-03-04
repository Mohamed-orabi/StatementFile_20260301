using Microsoft.EntityFrameworkCore;
using StatementFile.Application.Interfaces;
using StatementFile.Application.UseCases.BulkProcessing;
using StatementFile.Application.UseCases.MerchantOnboarding;
using StatementFile.Application.UseCases.StatementGeneration;
using StatementFile.Domain.Interfaces;
using StatementFile.Domain.Interfaces.Repositories;
using StatementFile.Domain.Interfaces.Services;
using StatementFile.Infrastructure.Data;
using StatementFile.Infrastructure.Data.Repositories;
using StatementFile.Infrastructure.Formatters;
using StatementFile.Infrastructure.Formatters.Html;
using StatementFile.Infrastructure.Formatters.RawData;
using StatementFile.Infrastructure.Formatters.TextLabel;
using StatementFile.Infrastructure.Formatters.Xml;
using StatementFile.Infrastructure.Services;

namespace StatementFile.Infrastructure.Configuration
{
    /// <summary>
    /// Static composition root that wires the entire dependency graph.
    ///
    /// All formatters are native implementations — no legacy clsStatHtml*,
    /// clsStatRawData*, clsStatTxtLbl* classes are referenced anywhere in this
    /// project. Each formatter extends the appropriate native base class and
    /// implements IStatementFormatterService directly.
    ///
    /// Schema constants (preserved from clsSessionValues):
    ///   SessionContext.StatementDbSchema = "A4M."  → TSTATEMENTMASTERTABLE / DETAILTABLE
    ///   GetMainSchema() from config                → client/reference tables
    ///
    /// To add a new bank-specific formatter:
    ///   1. Create a class that extends the appropriate native base
    ///      (NativeHtmlFormatterBase / NativeRawDataFormatterBase /
    ///       NativeTextLabelFormatterBase) with a unique FormatterKey.
    ///   2. Register it in the formatters array below — no other code changes needed.
    /// </summary>
    public static class DependencyInjection
    {
        public static CompositionRoot Compose()
        {
            // ── Singletons ─────────────────────────────────────────────────────────
            var configService  = new AppConfigurationService();
            var sessionContext = new SessionContext();   // StatementDbSchema defaults to "A4M."
            var connFactory    = new SqlConnectionFactory(configService);

            // ── EF Core DbContext options (used by BankProductConfigRepository) ────
            var dbContextOptions = new DbContextOptionsBuilder<AppDbContext>()
                .UseSqlServer(configService.GetSqlConnectionString())
                .Options;

            // ── Services ───────────────────────────────────────────────────────────
            var emailService      = new EmailService(configService);
            var ftpService        = new FtpService(configService);
            var reportService     = new ReportService();
            var maintenanceSvc    = new DataMaintenanceService(connFactory, configService, sessionContext);
            var queryService      = new StatementQueryService(connFactory, configService, sessionContext);
            var emailTracking     = new EmailTrackingService();
            var packagingService  = new FilePackagingService();
            var summaryService    = new StatementSummaryService(connFactory);
            var pageValidationSvc = new PageSizeValidationService();
            var statUpdateSvc     = new StatUpdateService(connFactory, sessionContext);

            // ── Formatter registry ────────────────────────────────────────────────
            // All formatters are native — no legacy class calls anywhere.
            // Open/Closed: to add a new bank, add one entry here only.
            var genericFormatter = new GenericHtmlStatementFormatter();
            var formatters       = new IStatementFormatterService[]
            {
                genericFormatter,

                // ── Native HTML formatters ─────────────────────────────────────────
                new HtmlUbaFormatter(),
                new HtmlAbpFormatter(),
                new HtmlAbpSupFormatter(),
                new HtmlAibkFormatter(),
                new HtmlAibkValuFormatter(),
                new HtmlAlxbFormatter(),
                new HtmlAlxbCpFormatter(),
                new HtmlBaiCreditFormatter(),
                new HtmlBaiPrepaidFormatter(),
                new HtmlBdcaFormatter(),
                new HtmlBpcFormatter(),
                new HtmlBpcPrepaidFormatter(),
                new HtmlCmbFormatter(),
                new HtmlDbnFormatter(),
                new HtmlFbpgFormatter(),
                new HtmlGtbkFormatter(),
                new HtmlGtbkDebitFormatter(),
                new HtmlGtbnFormatter(),
                new HtmlGtbuPrepaidFormatter(),
                new HtmlFbnFormatter(),
                new HtmlFbnCorpFormatter(),
                new HtmlFbnDebitFormatter(),
                new HtmlFbnSupFormatter(),
                new HtmlHblnFormatter(),
                new HtmlHblnPrepaidFormatter(),
                new HtmlIcbgFormatter(),
                new HtmlImbPrepaidFormatter(),
                new HtmlNbsFormatter(),
                new HtmlRbghFormatter(),
                new HtmlSbnFormatter(),
                new HtmlSbnNewFormatter(),
                new HtmlSbnSignatureFormatter(),
                new HtmlSbpFormatter(),
                new HtmlSbpDebitFormatter(),
                new HtmlSbpPrepaidFormatter(),
                new HtmlSibnFormatter(),
                new HtmlUbaGPrepaidFormatter(),
                new HtmlUnbnFormatter(),
                new HtmlWemaFormatter(),
                new HtmlWemaDebitFormatter(),
                new HtmlAaibFormatter(),
                new HtmlGnrImbPrepaidFormatter(),

                // ── Native PDF formatters (rendered as structured HTML) ─────────────
                new PdfQnbFormatter(),
                new PdfBdcaFormatter(),

                // ── Native RawData formatters ──────────────────────────────────────
                new RawDataAibkFormatter(),
                new RawDataEgbFormatter(),
                new RawDataAaibFormatter(),
                new RawDataAibkAltFormatter(),
                new RawDataAlxbFormatter(),
                new RawDataAlxbCorpFormatter(),
                new RawDataBrkaFormatter(),
                new RawDataUnbFormatter(),
                new RawDataVcbkFormatter(),

                // ── Native Text-label (fixed-width printer) formatters ─────────────
                new TextLabelFcmbFormatter(),
                new TextLabelSuezFormatter(),
                new TextLabelDbFbnFormatter(),
                new TextLabelDbAibFormatter(),
                new TextLabelDbBcaFormatter(),
                new TextLabelDbIcbgFormatter(),
                new TextLabelEdbeFormatter(),
                new TextLabelFabgFormatter(),

                // ── Native Plain-text (control-character) formatters ───────────────
                new TextEdbeFormatter(),

                // ── Native XML formatters ──────────────────────────────────────────
                new XmlIdbeFormatter(),
            };

            var formatterFactory = new StatementFormatterFactory(formatters, genericFormatter);

            // ── Repositories ───────────────────────────────────────────────────────
            var merchantRepo       = new MerchantStatementRepository();
            var bankProductCfgRepo = new BankProductConfigRepository(dbContextOptions);

            // ── Use Case Handlers ──────────────────────────────────────────────────
            var merchantHandler = new ProcessMerchantStatementHandler(
                merchantRepo, reportService, emailService, configService);

            var bulkHandler = new RunBulkMaintenanceHandler(maintenanceSvc);

            return new CompositionRoot(
                configService:      configService,
                sessionContext:     sessionContext,
                connFactory:        connFactory,
                emailService:       emailService,
                ftpService:         ftpService,
                reportService:      reportService,
                maintenanceService: maintenanceSvc,
                queryService:       queryService,
                emailTracking:      emailTracking,
                packagingService:   packagingService,
                summaryService:     summaryService,
                pageValidationSvc:  pageValidationSvc,
                statUpdateSvc:      statUpdateSvc,
                formatterFactory:   formatterFactory,
                merchantRepo:       merchantRepo,
                bankProductCfgRepo: bankProductCfgRepo,
                merchantHandler:    merchantHandler,
                bulkHandler:        bulkHandler);
        }
    }

    /// <summary>
    /// Value object that carries all fully-wired dependencies to the Presentation layer.
    /// </summary>
    public sealed class CompositionRoot
    {
        public AppConfigurationService         ConfigService      { get; }
        public SessionContext                  Session            { get; }
        public SqlConnectionFactory            ConnFactory        { get; }
        public IEmailService                   EmailService       { get; }
        public IFtpService                     FtpService         { get; }
        public IReportService                  ReportService      { get; }
        public IDataMaintenanceService         MaintenanceService { get; }
        public IStatementQueryService          QueryService       { get; }
        public IEmailTrackingService           EmailTracking      { get; }
        public IFilePackagingService           PackagingService   { get; }
        public IStatementSummaryService        SummaryService     { get; }
        public IPageSizeValidationService      PageValidation     { get; }
        public IStatUpdateService              StatUpdate         { get; }
        public IStatementFormatterFactory      FormatterFactory   { get; }
        public IMerchantStatementRepository    MerchantRepo       { get; }
        public IBankProductConfigRepository    BankProductCfgRepo { get; }
        public ProcessMerchantStatementHandler MerchantHandler    { get; }
        public RunBulkMaintenanceHandler       BulkHandler        { get; }

        internal CompositionRoot(
            AppConfigurationService         configService,
            SessionContext                  sessionContext,
            SqlConnectionFactory            connFactory,
            IEmailService                   emailService,
            IFtpService                     ftpService,
            IReportService                  reportService,
            IDataMaintenanceService         maintenanceService,
            IStatementQueryService          queryService,
            IEmailTrackingService           emailTracking,
            IFilePackagingService           packagingService,
            IStatementSummaryService        summaryService,
            IPageSizeValidationService      pageValidationSvc,
            IStatUpdateService              statUpdateSvc,
            IStatementFormatterFactory      formatterFactory,
            IMerchantStatementRepository    merchantRepo,
            IBankProductConfigRepository    bankProductCfgRepo,
            ProcessMerchantStatementHandler merchantHandler,
            RunBulkMaintenanceHandler       bulkHandler)
        {
            ConfigService      = configService;
            Session            = sessionContext;
            ConnFactory        = connFactory;
            EmailService       = emailService;
            FtpService         = ftpService;
            ReportService      = reportService;
            MaintenanceService = maintenanceService;
            QueryService       = queryService;
            EmailTracking      = emailTracking;
            PackagingService   = packagingService;
            SummaryService     = summaryService;
            PageValidation     = pageValidationSvc;
            StatUpdate         = statUpdateSvc;
            FormatterFactory   = formatterFactory;
            MerchantRepo       = merchantRepo;
            BankProductCfgRepo = bankProductCfgRepo;
            MerchantHandler    = merchantHandler;
            BulkHandler        = bulkHandler;
        }

        /// <summary>
        /// Creates a scoped UnitOfWork (one SQL Server connection per statement run).
        /// Caller is responsible for disposing it.
        /// </summary>
        public IUnitOfWork CreateUnitOfWork()
            => new UnitOfWork(ConnFactory, ConfigService, Session);

        /// <summary>
        /// Creates a GenerateStatementHandler for one statement run.
        /// </summary>
        public GenerateStatementHandler CreateStatementHandler()
            => new GenerateStatementHandler(
                CreateUnitOfWork(),
                FormatterFactory,
                QueryService,
                ReportService,
                EmailService,
                FtpService,
                EmailTracking,
                PackagingService,
                MaintenanceService,
                SummaryService,
                ConfigService);
    }
}

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
using StatementFile.Infrastructure.Services;

namespace StatementFile.Infrastructure.Configuration
{
    /// <summary>
    /// Static composition root that wires the entire dependency graph.
    ///
    /// In a full DI-container scenario this class would be replaced by container
    /// registrations.  For the Blazor Server presentation layer, calling
    /// <see cref="Compose"/> returns a fully wired <see cref="CompositionRoot"/>
    /// that the scoped Blazor services resolve from.
    ///
    /// Schema constants (preserved from clsSessionValues):
    ///   SessionContext.StatementDbSchema = "A4M."  → TSTATEMENTMASTERTABLE / DETAILTABLE
    ///   GetMainSchema() from config                → client/reference tables
    ///
    /// To add a new bank-specific formatter:
    ///   1. Implement IStatementFormatterService with a unique FormatterKey.
    ///   2. Register it in the formatters array below — no other code changes needed.
    /// </summary>
    public static class DependencyInjection
    {
        public static CompositionRoot Compose()
        {
            // ── Singletons ─────────────────────────────────────────────────────────
            var configService  = new AppConfigurationService();
            var sessionContext = new SessionContext();   // StatementDbSchema defaults to "A4M."
            var connFactory    = new OracleConnectionFactory(configService);

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

            // ── Formatter registry (Open/Closed: add new formatters here only) ─────
            var genericFormatter  = new GenericHtmlStatementFormatter();
            var formatters        = new IStatementFormatterService[]
            {
                genericFormatter,

                // ── HTML formatters ────────────────────────────────────────────────
                new HtmlUbaAdapter(),
                new HtmlAbpAdapter(),
                new HtmlAbpSupAdapter(),
                new HtmlAibkAdapter(),
                new HtmlAibkValuAdapter(),
                new HtmlAlxbAdapter(),
                new HtmlAlxbCpAdapter(),
                new HtmlBaiCreditAdapter(),
                new HtmlBaiPrepaidAdapter(),
                new HtmlBdcaAdapter(),
                new HtmlBpcAdapter(),
                new HtmlBpcPrepaidAdapter(),
                new HtmlCmbAdapter(),
                new HtmlDbnAdapter(),
                new HtmlFbpgAdapter(),
                new HtmlGtbkAdapter(),
                new HtmlGtbkDebitAdapter(),
                new HtmlGtbnAdapter(),
                new HtmlGtbuPrepaidAdapter(),
                new HtmlFbnAdapter(),
                new HtmlFbnCorpAdapter(),
                new HtmlFbnDebitAdapter(),
                new HtmlFbnSupAdapter(),
                new HtmlHblnAdapter(),
                new HtmlHblnPrepaidAdapter(),
                new HtmlIcbgAdapter(),
                new HtmlImbPrepaidAdapter(),
                new HtmlNbsAdapter(),
                new HtmlRbghAdapter(),
                new HtmlSbnAdapter(),
                new HtmlSbnNewAdapter(),
                new HtmlSbnSignatureAdapter(),
                new HtmlSbpAdapter(),
                new HtmlSbpDebitAdapter(),
                new HtmlSbpPrepaidAdapter(),
                new HtmlSibnAdapter(),
                new HtmlUbaGPrepaidAdapter(),
                new HtmlUnbnAdapter(),
                new HtmlWemaAdapter(),
                new HtmlWemaDebitAdapter(),
                new HtmlAaibAdapter(),
                new HtmlGnrImbPrepaidAdapter(),

                // ── PDF formatters ─────────────────────────────────────────────────
                new PdfQnbAdapter(),
                new PdfBdcaAdapter(),

                // ── RawData formatters ─────────────────────────────────────────────
                new RawDataAibkAdapter(),
                new RawDataEgbAdapter(),
                new RawDataAaibAdapter(),
                new RawDataAibkAltAdapter(),
                new RawDataAlxbAdapter(),
                new RawDataAlxbCorpAdapter(),
                new RawDataBrkaAdapter(),
                new RawDataUnbAdapter(),
                new RawDataVcbkAdapter(),

                // ── Text-label (fixed-width printer) formatters ────────────────────
                new TextLabelFcmbAdapter(),
                new TextLabelSuezAdapter(),
                new TextLabelDbFbnAdapter(),
                new TextLabelDbAibAdapter(),
                new TextLabelDbBcaAdapter(),
                new TextLabelDbIcbgAdapter(),
                new TextLabelEdbeAdapter(),
                new TextLabelFabgAdapter(),

                // ── Plain-text (control-character) formatters ──────────────────────
                new TextEdbeAdapter(),

                // ── XML formatters ─────────────────────────────────────────────────
                new XmlIdbeAdapter(),
            };
            var formatterFactory  = new StatementFormatterFactory(formatters, genericFormatter);

            // ── Repositories ───────────────────────────────────────────────────────
            var merchantRepo      = new MerchantStatementRepository();

            // ── Use Case Handlers ──────────────────────────────────────────────────
            var merchantHandler   = new ProcessMerchantStatementHandler(
                merchantRepo, reportService, emailService, configService);

            var bulkHandler       = new RunBulkMaintenanceHandler(maintenanceSvc);

            return new CompositionRoot(
                configService:     configService,
                sessionContext:    sessionContext,
                connFactory:       connFactory,
                emailService:      emailService,
                ftpService:        ftpService,
                reportService:     reportService,
                maintenanceService: maintenanceSvc,
                queryService:      queryService,
                emailTracking:     emailTracking,
                packagingService:  packagingService,
                summaryService:    summaryService,
                pageValidationSvc: pageValidationSvc,
                statUpdateSvc:     statUpdateSvc,
                formatterFactory:  formatterFactory,
                merchantRepo:      merchantRepo,
                merchantHandler:   merchantHandler,
                bulkHandler:       bulkHandler);
        }
    }

    /// <summary>
    /// Value object that carries all fully-wired dependencies to the Presentation layer.
    /// </summary>
    public sealed class CompositionRoot
    {
        public AppConfigurationService         ConfigService      { get; }
        public SessionContext                  Session            { get; }
        public OracleConnectionFactory         ConnFactory        { get; }
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
        public ProcessMerchantStatementHandler MerchantHandler    { get; }
        public RunBulkMaintenanceHandler       BulkHandler        { get; }

        internal CompositionRoot(
            AppConfigurationService         configService,
            SessionContext                  sessionContext,
            OracleConnectionFactory         connFactory,
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
            MerchantHandler    = merchantHandler;
            BulkHandler        = bulkHandler;
        }

        /// <summary>
        /// Creates a scoped UnitOfWork (one Oracle connection per statement run).
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

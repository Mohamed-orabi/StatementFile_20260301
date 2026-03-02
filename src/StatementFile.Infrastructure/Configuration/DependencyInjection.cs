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
using StatementFile.Infrastructure.Services;

namespace StatementFile.Infrastructure.Configuration
{
    /// <summary>
    /// Static composition root that wires the entire dependency graph.
    ///
    /// In a full DI-container scenario (Unity, Autofac, MS.Extensions.DI) this
    /// class would be replaced by container registrations.  For .NET Framework
    /// Windows Forms apps without a container, calling <see cref="Compose"/>
    /// returns a fully wired <see cref="CompositionRoot"/> that the presentation
    /// layer uses to resolve handlers.
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
            var configService    = new AppConfigurationService();
            var sessionContext   = new SessionContext();
            var connFactory      = new OracleConnectionFactory(configService);

            // ── Services ───────────────────────────────────────────────────────────
            var emailService     = new EmailService(configService);
            var ftpService       = new FtpService(configService);
            var reportService    = new ReportService();
            var maintenanceSvc   = new DataMaintenanceService(connFactory, configService, sessionContext);

            // ── Formatter registry (Open/Closed: add new formatters here only) ─────
            var genericFormatter = new GenericHtmlStatementFormatter();
            var formatters       = new IStatementFormatterService[]
            {
                genericFormatter,
                // Add bank-specific formatters here, e.g.:
                // new BaiHtmlStatementFormatter(),
                // new UbaTxtStatementFormatter(),
            };
            var formatterFactory = new StatementFormatterFactory(formatters, genericFormatter);

            // ── Repositories & UoW factory ─────────────────────────────────────────
            var merchantRepo     = new MerchantStatementRepository();

            // ── Use Case Handlers ──────────────────────────────────────────────────
            var merchantHandler  = new ProcessMerchantStatementHandler(
                merchantRepo, reportService, emailService, configService);

            var bulkHandler      = new RunBulkMaintenanceHandler(maintenanceSvc);

            return new CompositionRoot(
                configService:       configService,
                sessionContext:      sessionContext,
                connFactory:         connFactory,
                emailService:        emailService,
                ftpService:          ftpService,
                reportService:       reportService,
                maintenanceService:  maintenanceSvc,
                formatterFactory:    formatterFactory,
                merchantRepo:        merchantRepo,
                merchantHandler:     merchantHandler,
                bulkHandler:         bulkHandler);
        }
    }

    /// <summary>
    /// Value object that carries all fully-wired dependencies to the Presentation layer.
    /// </summary>
    public sealed class CompositionRoot
    {
        public AppConfigurationService        ConfigService      { get; }
        public SessionContext                 Session            { get; }
        public OracleConnectionFactory        ConnFactory        { get; }
        public IEmailService                  EmailService       { get; }
        public IFtpService                    FtpService         { get; }
        public IReportService                 ReportService      { get; }
        public IDataMaintenanceService        MaintenanceService { get; }
        public IStatementFormatterFactory     FormatterFactory   { get; }
        public IMerchantStatementRepository   MerchantRepo       { get; }
        public ProcessMerchantStatementHandler MerchantHandler   { get; }
        public RunBulkMaintenanceHandler      BulkHandler        { get; }

        internal CompositionRoot(
            AppConfigurationService        configService,
            SessionContext                 sessionContext,
            OracleConnectionFactory        connFactory,
            IEmailService                  emailService,
            IFtpService                    ftpService,
            IReportService                 reportService,
            IDataMaintenanceService        maintenanceService,
            IStatementFormatterFactory     formatterFactory,
            IMerchantStatementRepository   merchantRepo,
            ProcessMerchantStatementHandler merchantHandler,
            RunBulkMaintenanceHandler      bulkHandler)
        {
            ConfigService      = configService;
            Session            = sessionContext;
            ConnFactory        = connFactory;
            EmailService       = emailService;
            FtpService         = ftpService;
            ReportService      = reportService;
            MaintenanceService = maintenanceService;
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
                ReportService,
                EmailService,
                FtpService,
                ConfigService);
    }
}

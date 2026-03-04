using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using StatementFile.Application.Commands;
using StatementFile.Domain.Interfaces;
using StatementFile.Infrastructure.Email;
using StatementFile.Infrastructure.Formatters;
using StatementFile.Infrastructure.Oracle;

namespace StatementFile.Infrastructure.Configuration
{
    /// <summary>
    /// Extension method that wires all Infrastructure services into the .NET 8 DI container.
    ///
    /// Call <c>services.AddInfrastructure(configuration)</c> from Program.cs.
    ///
    /// This replaces the manual <c>CompositionRoot</c> pattern used in the
    /// original WinForms and src / src2 solutions.
    /// </summary>
    public static class InfrastructureServiceExtensions
    {
        public static IServiceCollection AddInfrastructure(
            this IServiceCollection services,
            IConfiguration          configuration)
        {
            // ── Configuration ─────────────────────────────────────────────────────
            services.AddSingleton<IConfigurationService, AppConfigurationService>();

            // ── Oracle ────────────────────────────────────────────────────────────
            var connStr = configuration.GetConnectionString("Oracle")
                          ?? configuration["Oracle:ConnectionString"]
                          ?? string.Empty;

            services.AddSingleton<IDbConnectionFactory>(_ => new OracleConnectionFactory(connStr));
            services.AddSingleton<IBankProductConfigRepository, BankProductConfigRepository>();
            services.AddSingleton<IStatementRunRepository, StatementRunRepository>();
            services.AddSingleton<IBulkMaintenanceService, OracleBulkMaintenanceService>();

            // ── Email ─────────────────────────────────────────────────────────────
            var smtp = new SmtpSettings
            {
                Host     = configuration["Smtp:Host"]     ?? "localhost",
                Port     = int.TryParse(configuration["Smtp:Port"], out var p) ? p : 25,
                Username = configuration["Smtp:Username"],
                Password = configuration["Smtp:Password"],
                UseSsl   = bool.TryParse(configuration["Smtp:UseSsl"], out var s) && s,
            };
            services.AddSingleton(smtp);
            services.AddSingleton<IEmailService, SmtpEmailService>();

            // ── Formatter registry ────────────────────────────────────────────────
            // Register one formatter per FormatterKey value.
            // New banks are added here — no switch-case required anywhere.
            services.AddSingleton<FormatterRegistry>(sp =>
            {
                var email   = sp.GetRequiredService<IEmailService>();
                var htmlFmt = new HtmlStatementFormatter(email);
                var txtFmt  = new TextStatementFormatter(email);
                var rawFmt  = new RawDataExportFormatter();

                var registry = new FormatterRegistry();

                // ── HTML formatters (one entry per bank that uses HTML output) ───
                foreach (var key in new[]
                {
                    "HTML_UBA",   "HTML_BDCA",  "HTML_AIB",   "HTML_BAI",
                    "HTML_BIC",   "HTML_NSGB",  "HTML_CORP",  "HTML_PREPAID",
                    "HTML_RAS",   "HTML_REWARD", "HTML_DEFAULT"
                })
                    registry.Register(key, htmlFmt);

                // ── Text / export formatters ──────────────────────────────────────
                foreach (var key in new[]
                {
                    "TXT_NSGB", "TXT_EXPORT", "TXT_DEFAULT"
                })
                    registry.Register(key, txtFmt);

                // ── Raw-data / CSV formatters ─────────────────────────────────────
                foreach (var key in new[]
                {
                    "RAWDATA_EXPORT", "RAWDATA_REWARD", "RAWDATA_RAS",
                    "CSV_EXPORT",     "CSV_DEFAULT"
                })
                    registry.Register(key, rawFmt);

                return registry;
            });

            // Expose as the domain interface
            services.AddSingleton<IFormatterRegistry>(sp =>
                sp.GetRequiredService<FormatterRegistry>());

            // ── Application handlers ──────────────────────────────────────────────
            services.AddTransient<GenerateStatementHandler>();
            services.AddTransient<RunMaintenanceHandler>();

            return services;
        }
    }
}

using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using StatementFile.Application.Commands;
using StatementFile.Domain.Interfaces;
using StatementFile.Infrastructure.Data;
using StatementFile.Infrastructure.Email;
using StatementFile.Infrastructure.Formatters;

namespace StatementFile.Infrastructure.Configuration
{
    /// <summary>
    /// Wires all Infrastructure services into the .NET 8 DI container.
    /// Call <c>services.AddInfrastructure(configuration)</c> from Program.cs.
    /// </summary>
    public static class InfrastructureServiceExtensions
    {
        public static IServiceCollection AddInfrastructure(
            this IServiceCollection services,
            IConfiguration          configuration)
        {
            // ── Configuration service ─────────────────────────────────────────────
            services.AddSingleton<IConfigurationService, AppConfigurationService>();

            // ── EF Core – SQL Server ──────────────────────────────────────────────
            var connStr = configuration.GetConnectionString("SqlServer")
                          ?? configuration["SqlServer:ConnectionString"]
                          ?? string.Empty;

            services.AddDbContext<StatementFileDbContext>(opts =>
                opts.UseSqlServer(connStr,
                    sql => sql.CommandTimeout(300)));

            // Scoped repositories (lifetime matches DbContext)
            services.AddScoped<IBankProductConfigRepository, EfBankProductConfigRepository>();
            services.AddScoped<IStatementRunRepository,      EfStatementRunRepository>();

            // Bulk maintenance uses ADO.NET directly (stored-proc calls)
            services.AddSingleton<IBulkMaintenanceService, SqlBulkMaintenanceService>();

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
            services.AddSingleton<FormatterRegistry>(sp =>
            {
                var email   = sp.GetRequiredService<IEmailService>();
                var htmlFmt = new HtmlStatementFormatter(email);
                var txtFmt  = new TextStatementFormatter(email);
                var rawFmt  = new RawDataExportFormatter();

                var registry = new FormatterRegistry();

                foreach (var key in new[]
                {
                    "HTML_UBA", "HTML_BDCA", "HTML_AIB", "HTML_BAI",
                    "HTML_BIC", "HTML_NSGB", "HTML_CORP", "HTML_PREPAID",
                    "HTML_RAS", "HTML_REWARD", "HTML_DEFAULT"
                })
                    registry.Register(key, htmlFmt);

                foreach (var key in new[] { "TXT_NSGB", "TXT_EXPORT", "TXT_DEFAULT" })
                    registry.Register(key, txtFmt);

                foreach (var key in new[]
                {
                    "RAWDATA_EXPORT", "RAWDATA_REWARD", "RAWDATA_RAS",
                    "CSV_EXPORT",     "CSV_DEFAULT"
                })
                    registry.Register(key, rawFmt);

                return registry;
            });

            services.AddSingleton<IFormatterRegistry>(sp =>
                sp.GetRequiredService<FormatterRegistry>());

            // ── Application handlers ──────────────────────────────────────────────
            services.AddScoped<GenerateStatementHandler>();
            services.AddScoped<RunMaintenanceHandler>();

            return services;
        }
    }
}

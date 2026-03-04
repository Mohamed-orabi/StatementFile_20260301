using System;

namespace StatementFile.Application.DTOs
{
    /// <summary>
    /// Flat, fully-serialisable request DTO for the statement generation API endpoint.
    ///
    /// Mirrors every parameter of <see cref="StatementFile.Application.UseCases.StatementGeneration.GenerateStatementCommand"/>
    /// so the Presentation layer can build it without referencing the command class directly.
    ///
    /// Enum values are sent as their underlying integers; the API controller
    /// casts them back before constructing the command.
    /// </summary>
    public sealed class GenerateStatementApiRequest
    {
        // ── Identity ──────────────────────────────────────────────────────────
        public int    BranchCode                   { get; set; }
        public string BankName                     { get; set; } = string.Empty;
        public string BankFullName                 { get; set; } = string.Empty;
        public string CardProduct                  { get; set; } = string.Empty;

        // ── Output format (int = StatementOutputType; long = ProcessingMode) ──
        public int    OutputType                   { get; set; }
        public string FormatterKey                 { get; set; } = string.Empty;
        public long   ProcessingModes              { get; set; }

        // ── Card classification ───────────────────────────────────────────────
        public int    CardType                     { get; set; }
        public int    StatementType                { get; set; }

        // ── Date + path ───────────────────────────────────────────────────────
        public DateTime StatementDate              { get; set; } = DateTime.Today;

        /// <summary>
        /// Override the statement output root path.
        /// When null or empty the API uses the configured default
        /// (pathConfiguration.json / App.config "stmtPath").
        /// </summary>
        public string OutputRootPath               { get; set; }

        // ── Optional flags ────────────────────────────────────────────────────
        public bool   AppendMode                   { get; set; }
        public string RewardContractCondition      { get; set; } = "'New Reward Contract'";
        public string InstallmentContractCondition { get; set; }
        public bool   UseCorporateAccountNumber    { get; set; }
        public string StatementTypeSuffix          { get; set; } = "CR";
        public string ReportTemplateName           { get; set; }
        public string PdfPasswordScheme            { get; set; }
        public string EmailFromAddress             { get; set; } = "cardservices@emp-group.com";
        public string BankWebsiteUrl               { get; set; }
        public string FacebookUrl                  { get; set; }
        public string LinkedInUrl                  { get; set; }
        public string YouTubeUrl                   { get; set; }
        public int    MaxTransactionsPerPage       { get; set; } = 20;
        public int    MaxTransactionsLastPage      { get; set; } = 27;
        public int    EmailWaitPeriodMs            { get; set; } = 7000;
        public string FieldSeparator               { get; set; } = ",";
    }
}

using System;

namespace StatementFile.Application.DTOs
{
    /// <summary>
    /// Flat, fully-serialisable projection of <c>BankProductConfig</c>.
    ///
    /// Enum values are stored as their underlying integers so that
    /// System.Text.Json can round-trip them without requiring
    /// [JsonConverter] attributes or shared assembly references.
    ///
    /// Used by:
    ///   - StatementFile.Api      → maps entity → DTO → JSON response
    ///   - StatementFile.Presentation → deserialises API response → UI model
    /// </summary>
    public sealed class BankConfigDto
    {
        // ── Primary key ───────────────────────────────────────────────────────
        public int    Id                           { get; set; }

        // ── Identity ──────────────────────────────────────────────────────────
        public int    BranchCode                   { get; set; }
        public string BankName                     { get; set; } = string.Empty;
        public string BankFullName                 { get; set; } = string.Empty;
        public string CardProduct                  { get; set; } = string.Empty;

        // ── Output format (int = StatementOutputType enum) ────────────────────
        public int    OutputType                   { get; set; }
        public string FormatterKey                 { get; set; } = string.Empty;

        // ── Processing flags (long = ProcessingMode [Flags] enum) ────────────
        public long   ProcessingModes              { get; set; }

        // ── Card classification (ints = CardType / StatementType enums) ───────
        public int    CardType                     { get; set; }
        public int    StatementType                { get; set; }

        // ── Statement suffix ──────────────────────────────────────────────────
        public string StatementTypeSuffix          { get; set; } = "CR";

        // ── Optional Oracle conditions ────────────────────────────────────────
        public string RewardContractCondition      { get; set; } = "'New Reward Contract'";
        public string InstallmentContractCondition { get; set; }

        // ── Flags ─────────────────────────────────────────────────────────────
        public bool   AppendMode                   { get; set; }
        public bool   UseCorporateAccountNumber    { get; set; }

        // ── PDF ───────────────────────────────────────────────────────────────
        public string ReportTemplateName           { get; set; }
        public string PdfPasswordScheme            { get; set; }

        // ── Email / branding ──────────────────────────────────────────────────
        public string EmailFromAddress             { get; set; } = "cardservices@emp-group.com";
        public string BankWebsiteUrl               { get; set; }
        public string FacebookUrl                  { get; set; }
        public string LinkedInUrl                  { get; set; }
        public string YouTubeUrl                   { get; set; }

        // ── Pagination ────────────────────────────────────────────────────────
        public int    MaxTransactionsPerPage       { get; set; } = 20;
        public int    MaxTransactionsLastPage      { get; set; } = 27;

        // ── Email delivery ────────────────────────────────────────────────────
        public int    EmailWaitPeriodMs            { get; set; } = 7000;

        // ── Raw data ──────────────────────────────────────────────────────────
        public string FieldSeparator               { get; set; } = ",";

        // ── Scheduling ────────────────────────────────────────────────────────
        public int    ScheduledDay                 { get; set; } = 1;

        // ── Display ───────────────────────────────────────────────────────────
        public string DisplayName                  { get; set; }
        public int    SortOrder                    { get; set; }
        public bool   IsActive                     { get; set; } = true;

        // ── Audit ─────────────────────────────────────────────────────────────
        public DateTime CreatedAt                  { get; set; }
        public DateTime ModifiedAt                 { get; set; }
    }
}

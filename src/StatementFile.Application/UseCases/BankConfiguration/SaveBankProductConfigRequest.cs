namespace StatementFile.Application.UseCases.BankConfiguration
{
    /// <summary>
    /// Request body sent from Blazor to the bank-config API for both
    /// POST (create) and PUT (update) operations.
    ///
    /// Enum fields are stored as their underlying integer values so that
    /// the JSON contract remains stable across enum additions.
    /// </summary>
    public sealed class SaveBankProductConfigRequest
    {
        // ── Identity (required for POST; ignored on PUT — identity is immutable) ──
        public int    BranchCode { get; set; }
        public string BankName   { get; set; }

        // ── Mutable fields ───────────────────────────────────────────────────────
        public string BankFullName                { get; set; }
        public string CardProduct                 { get; set; }
        public int    OutputType                  { get; set; }   // StatementOutputType
        public string FormatterKey                { get; set; }
        public long   ProcessingModes             { get; set; }   // ProcessingMode bitmask
        public int    CardType                    { get; set; }   // CardType enum value
        public int    StatementType               { get; set; }   // StatementType enum value
        public string StatementTypeSuffix         { get; set; } = "CR";
        public string RewardContractCondition     { get; set; }
        public string InstallmentContractCondition{ get; set; }
        public bool   AppendMode                  { get; set; }
        public bool   UseCorporateAccountNumber   { get; set; }
        public string ReportTemplateName          { get; set; }
        public string PdfPasswordScheme           { get; set; }
        public string EmailFromAddress            { get; set; } = "cardservices@emp-group.com";
        public string BankWebsiteUrl              { get; set; }
        public string FacebookUrl                 { get; set; }
        public string LinkedInUrl                 { get; set; }
        public string YouTubeUrl                  { get; set; }
        public int    MaxTransactionsPerPage      { get; set; } = 20;
        public int    MaxTransactionsLastPage     { get; set; } = 27;
        public int    EmailWaitPeriodMs           { get; set; } = 7000;
        public string FieldSeparator              { get; set; } = ",";
        public int    ScheduledDay                { get; set; } = 1;
        public string DisplayName                 { get; set; }
        public int    SortOrder                   { get; set; }
        public bool   IsActive                    { get; set; } = true;
    }
}

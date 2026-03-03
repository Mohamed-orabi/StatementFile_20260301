using System;
using StatementFile.Domain.Enums;

namespace StatementFile.Domain.Entities
{
    /// <summary>
    /// Persistent configuration record for one bank/product statement run.
    ///
    /// Stored in Oracle table STAT_BANK_PRODUCT_CONFIG.
    /// Each row maps to all parameters needed by GenerateStatementCommand,
    /// so adding a new bank never requires a code change — only a new DB row.
    ///
    /// Replaces the hundreds of hardcoded chkLstProducts.Items.Add() calls
    /// in the legacy frmStatementFileExtn.cs.
    /// </summary>
    public sealed class BankProductConfig
    {
        // ── Primary key ───────────────────────────────────────────────────────────
        public int Id { get; private set; }

        // ── Identity ──────────────────────────────────────────────────────────────
        public int    BranchCode   { get; private set; }
        public string BankName     { get; private set; }   // short code, e.g. "UBA"
        public string BankFullName { get; private set; }   // display name
        public string CardProduct  { get; private set; }   // Oracle product code

        // ── Output format ─────────────────────────────────────────────────────────
        public StatementOutputType OutputType    { get; private set; }
        public string              FormatterKey  { get; private set; }

        // ── Processing flags ──────────────────────────────────────────────────────
        public ProcessingMode ProcessingModes { get; private set; }

        // ── Card classification ───────────────────────────────────────────────────
        public CardType      CardType      { get; private set; }
        public StatementType StatementType { get; private set; }

        // ── Statement identity suffix ─────────────────────────────────────────────
        public string StatementTypeSuffix { get; private set; }  // "CR","DB","CP","VIP"…

        // ── Optional Oracle conditions ────────────────────────────────────────────
        public string RewardContractCondition      { get; private set; }
        public string InstallmentContractCondition { get; private set; }

        // ── Flags ─────────────────────────────────────────────────────────────────
        public bool AppendMode               { get; private set; }
        public bool UseCorporateAccountNumber { get; private set; }

        // ── PDF ───────────────────────────────────────────────────────────────────
        public string ReportTemplateName { get; private set; }
        public string PdfPasswordScheme  { get; private set; }

        // ── Email / branding ──────────────────────────────────────────────────────
        public string EmailFromAddress { get; private set; }
        public string BankWebsiteUrl   { get; private set; }
        public string FacebookUrl      { get; private set; }
        public string LinkedInUrl      { get; private set; }
        public string YouTubeUrl       { get; private set; }

        // ── Text-label pagination ─────────────────────────────────────────────────
        public int MaxTransactionsPerPage  { get; private set; }
        public int MaxTransactionsLastPage { get; private set; }

        // ── Email delivery ────────────────────────────────────────────────────────
        public int EmailWaitPeriodMs { get; private set; }

        // ── Raw data ──────────────────────────────────────────────────────────────
        public string FieldSeparator { get; private set; }

        // ── Scheduling ────────────────────────────────────────────────────────────
        /// <summary>Day-of-month on which this product is typically scheduled (1–31).</summary>
        public int ScheduledDay { get; private set; }

        // ── Display ───────────────────────────────────────────────────────────────
        /// <summary>
        /// Friendly label shown in the UI product list.
        /// Equivalent to the legacy ListItem text, e.g.
        /// "29) UBA United Bank for Africa Plc Nigeria >> Credit Emails 1/m".
        /// </summary>
        public string DisplayName { get; private set; }

        /// <summary>Sort order in the UI product list.</summary>
        public int SortOrder { get; private set; }

        /// <summary>When false the product does not appear in the UI.</summary>
        public bool IsActive { get; private set; }

        // ── Audit ─────────────────────────────────────────────────────────────────
        public DateTime CreatedAt  { get; private set; }
        public DateTime ModifiedAt { get; private set; }

        private BankProductConfig() { }

        public static BankProductConfig Create(
            int                branchCode,
            string             bankName,
            string             bankFullName,
            string             cardProduct,
            StatementOutputType outputType,
            string             formatterKey,
            ProcessingMode     processingModes          = ProcessingMode.Standard,
            CardType           cardType                 = CardType.Credit,
            StatementType      statementType            = StatementType.Normal,
            string             statementTypeSuffix      = "CR",
            string             rewardContractCondition  = "'New Reward Contract'",
            string             installmentContractCondition = null,
            bool               appendMode               = false,
            bool               useCorporateAccountNumber = false,
            string             reportTemplateName       = null,
            string             pdfPasswordScheme        = null,
            string             emailFromAddress         = "cardservices@emp-group.com",
            string             bankWebsiteUrl           = "www.emp-group.com",
            string             facebookUrl              = null,
            string             linkedInUrl              = null,
            string             youTubeUrl               = null,
            int                maxTransactionsPerPage   = 20,
            int                maxTransactionsLastPage  = 27,
            int                emailWaitPeriodMs        = 7000,
            string             fieldSeparator           = ",",
            int                scheduledDay             = 1,
            string             displayName              = null,
            int                sortOrder                = 0,
            bool               isActive                 = true)
        {
            if (string.IsNullOrWhiteSpace(bankName))
                throw new ArgumentException("BankName is required.", nameof(bankName));
            if (string.IsNullOrWhiteSpace(formatterKey))
                throw new ArgumentException("FormatterKey is required.", nameof(formatterKey));

            var now = DateTime.UtcNow;
            return new BankProductConfig
            {
                BranchCode                   = branchCode,
                BankName                     = bankName,
                BankFullName                 = bankFullName ?? bankName,
                CardProduct                  = cardProduct ?? string.Empty,
                OutputType                   = outputType,
                FormatterKey                 = formatterKey,
                ProcessingModes              = processingModes,
                CardType                     = cardType,
                StatementType                = statementType,
                StatementTypeSuffix          = statementTypeSuffix ?? "CR",
                RewardContractCondition      = rewardContractCondition ?? "'New Reward Contract'",
                InstallmentContractCondition = installmentContractCondition,
                AppendMode                   = appendMode,
                UseCorporateAccountNumber    = useCorporateAccountNumber,
                ReportTemplateName           = reportTemplateName,
                PdfPasswordScheme            = pdfPasswordScheme,
                EmailFromAddress             = emailFromAddress ?? "cardservices@emp-group.com",
                BankWebsiteUrl               = bankWebsiteUrl  ?? "www.emp-group.com",
                FacebookUrl                  = facebookUrl,
                LinkedInUrl                  = linkedInUrl,
                YouTubeUrl                   = youTubeUrl,
                MaxTransactionsPerPage       = maxTransactionsPerPage > 0 ? maxTransactionsPerPage : 20,
                MaxTransactionsLastPage      = maxTransactionsLastPage > 0 ? maxTransactionsLastPage : 27,
                EmailWaitPeriodMs            = emailWaitPeriodMs > 0 ? emailWaitPeriodMs : 7000,
                FieldSeparator               = fieldSeparator ?? ",",
                ScheduledDay                 = scheduledDay,
                DisplayName                  = displayName ?? $"{branchCode}) {bankName} {bankFullName} >> {outputType} {scheduledDay}/m",
                SortOrder                    = sortOrder,
                IsActive                     = isActive,
                CreatedAt                    = now,
                ModifiedAt                   = now,
            };
        }

        /// <summary>Used by the repository to reconstitute a persisted entity (sets Id).</summary>
        internal static BankProductConfig Reconstitute(int id, BankProductConfig config)
        {
            config.Id = id;
            return config;
        }

        /// <summary>Apply an edit and update the ModifiedAt timestamp.</summary>
        public void Update(
            string             bankFullName,
            string             cardProduct,
            StatementOutputType outputType,
            string             formatterKey,
            ProcessingMode     processingModes,
            CardType           cardType,
            StatementType      statementType,
            string             statementTypeSuffix,
            string             rewardContractCondition,
            string             installmentContractCondition,
            bool               appendMode,
            bool               useCorporateAccountNumber,
            string             reportTemplateName,
            string             pdfPasswordScheme,
            string             emailFromAddress,
            string             bankWebsiteUrl,
            string             facebookUrl,
            string             linkedInUrl,
            string             youTubeUrl,
            int                maxTransactionsPerPage,
            int                maxTransactionsLastPage,
            int                emailWaitPeriodMs,
            string             fieldSeparator,
            int                scheduledDay,
            string             displayName,
            int                sortOrder,
            bool               isActive)
        {
            BankFullName                 = bankFullName ?? BankFullName;
            CardProduct                  = cardProduct  ?? CardProduct;
            OutputType                   = outputType;
            FormatterKey                 = formatterKey ?? FormatterKey;
            ProcessingModes              = processingModes;
            CardType                     = cardType;
            StatementType                = statementType;
            StatementTypeSuffix          = statementTypeSuffix      ?? StatementTypeSuffix;
            RewardContractCondition      = rewardContractCondition   ?? RewardContractCondition;
            InstallmentContractCondition = installmentContractCondition;
            AppendMode                   = appendMode;
            UseCorporateAccountNumber    = useCorporateAccountNumber;
            ReportTemplateName           = reportTemplateName;
            PdfPasswordScheme            = pdfPasswordScheme;
            EmailFromAddress             = emailFromAddress ?? EmailFromAddress;
            BankWebsiteUrl               = bankWebsiteUrl  ?? BankWebsiteUrl;
            FacebookUrl                  = facebookUrl;
            LinkedInUrl                  = linkedInUrl;
            YouTubeUrl                   = youTubeUrl;
            MaxTransactionsPerPage       = maxTransactionsPerPage  > 0 ? maxTransactionsPerPage  : MaxTransactionsPerPage;
            MaxTransactionsLastPage      = maxTransactionsLastPage > 0 ? maxTransactionsLastPage : MaxTransactionsLastPage;
            EmailWaitPeriodMs            = emailWaitPeriodMs       > 0 ? emailWaitPeriodMs       : EmailWaitPeriodMs;
            FieldSeparator               = fieldSeparator ?? FieldSeparator;
            ScheduledDay                 = scheduledDay;
            DisplayName                  = displayName ?? DisplayName;
            SortOrder                    = sortOrder;
            IsActive                     = isActive;
            ModifiedAt                   = DateTime.UtcNow;
        }
    }
}

using System;
using StatementFile.Domain.Enums;

namespace StatementFile.Application.UseCases.StatementGeneration
{
    /// <summary>
    /// Fully parameterised command for a single bank/product statement generation run.
    ///
    /// Every configurable axis found across all 30+ bank formatters in the Banks/ folder
    /// is represented here so the Application layer never needs to inspect the bank
    /// name or formatter key to make business decisions.
    ///
    /// The UI (frmGenerateStatement) populates this command from the product-configuration
    /// table; the handler dispatches purely on command properties.
    /// </summary>
    public sealed class GenerateStatementCommand
    {
        // ── Identity ──────────────────────────────────────────────────────────────

        /// <summary>Oracle branch code, e.g. 4 (AIBK), 6 (AAIB), 16 (ABP), 25 (AUB).</summary>
        public int BranchCode { get; }

        /// <summary>Short bank code used in directory/file naming, e.g. "UBA", "GTBK".</summary>
        public string BankName { get; }

        /// <summary>Human-readable bank name for email subjects / statement headers.</summary>
        public string BankFullName { get; }

        /// <summary>
        /// Oracle product code from TPRODUCTTABLE.CODE, e.g. "UBACR", "GTBKCR_15".
        /// Drives the Oracle WHERE clause product filter.
        /// </summary>
        public string CardProduct { get; }

        // ── Output Type and Format ────────────────────────────────────────────────

        /// <summary>Physical format of the files produced by this run.</summary>
        public StatementOutputType OutputType { get; }

        /// <summary>
        /// Formatter registry key used to resolve IStatementFormatterService.
        /// Convention: "{OutputType}_{BankCode}" e.g. "HTML_UBA", "RAWDATA_AIBK",
        /// "TEXTLABEL_FCMB", "XML_IDBE", "PDF_QNB", "HTML_ABP_SUP".
        /// </summary>
        public string FormatterKey { get; }

        // ── Processing Modes (flags) ──────────────────────────────────────────────

        /// <summary>Combined flags controlling optional processing branches.</summary>
        public ProcessingMode ProcessingModes { get; }

        // ── Card Classification ───────────────────────────────────────────────────

        public CardType      CardType      { get; }
        public StatementType StatementType { get; }

        // ── Date and Path ─────────────────────────────────────────────────────────

        /// <summary>
        /// Processing date. Drives output directory naming (yyyyMM prefix)
        /// and the statement period.
        /// </summary>
        public DateTime StatementDate { get; }

        /// <summary>Root output path from pathConfiguration.json, e.g. "D:\Statements\".</summary>
        public string OutputRootPath { get; }

        // ── Append Mode ───────────────────────────────────────────────────────────

        /// <summary>
        /// When true the output directory is NOT deleted before generating
        /// (legacy pAppendData flag). New files are appended alongside existing ones.
        /// </summary>
        public bool AppendMode { get; }

        // ── Reward Programme ──────────────────────────────────────────────────────

        /// <summary>
        /// Oracle IN-condition for reward contract types,
        /// e.g. "'New Reward Contract'" or "'Reward Program (Airmile)'".
        /// Used only when ProcessingModes includes Reward.
        /// </summary>
        public string RewardContractCondition { get; }

        // ── Installment Programme ─────────────────────────────────────────────────

        /// <summary>
        /// Oracle IN-condition for installment contract types,
        /// e.g. "'Purchase Installment With Interest Rate'".
        /// Used only when ProcessingModes includes Installment.
        /// </summary>
        public string InstallmentContractCondition { get; }

        // ── Corporate Mode ────────────────────────────────────────────────────────

        /// <summary>
        /// When ProcessingModes includes Corporate, controls whether the account number
        /// is read from cardaccountno (true) or accountno (false).
        /// Legacy: accountNoName field in clsBasStatementHtml.
        /// </summary>
        public bool UseCorporateAccountNumber { get; }

        // ── Statement Type Suffix ─────────────────────────────────────────────────

        /// <summary>
        /// Short suffix appended to directory and file names:
        /// "CR" (Credit), "DB" (Debit), "CP" (Prepaid), "RWD" (Reward),
        /// "CORP" (Corporate), "VIP" (VIP).
        /// Legacy: pStrFile / pStmntType parameter in Statement() overloads.
        /// </summary>
        public string StatementTypeSuffix { get; }

        // ── PDF-specific ──────────────────────────────────────────────────────────

        /// <summary>
        /// Crystal Reports .rpt template file name (relative to Reports/ folder).
        /// Required when OutputType == Pdf.
        /// </summary>
        public string ReportTemplateName { get; }

        /// <summary>
        /// Password derivation scheme for PasswordProtectedPdf mode.
        /// "DOB"   = customer date of birth (DDMMYYYY).
        /// "CARD4" = last 4 digits of card number.
        /// null    = no password.
        /// Legacy: PdfSharp.Pdf.Security in clsStatHtmlUBA/ABP.
        /// </summary>
        public string PdfPasswordScheme { get; }

        // ── Email / Branding ──────────────────────────────────────────────────────

        /// <summary>From-address for all outbound emails in this run.</summary>
        public string EmailFromAddress { get; }

        /// <summary>Bank website URL embedded in HTML statement footers.</summary>
        public string BankWebsiteUrl { get; }

        /// <summary>
        /// Social media URLs embedded in HTML statement footers (ALXB-specific).
        /// Other banks leave these null.
        /// </summary>
        public string FacebookUrl    { get; }
        public string LinkedInUrl    { get; }
        public string YouTubeUrl     { get; }

        // ── Print Formatting ─────────────────────────────────────────────────────

        /// <summary>
        /// Max transaction rows per page for text/label formatters.
        /// Default: 20 (matching MaxDetailInPage in all legacy classes).
        /// </summary>
        public int MaxTransactionsPerPage { get; }

        /// <summary>
        /// Max transaction rows on the final page for text/label formatters.
        /// Default: 27 (matching MaxDetailInLastPage in all legacy classes).
        /// </summary>
        public int MaxTransactionsLastPage { get; }

        // ── Email Delivery Timing ────────────────────────────────────────────────

        /// <summary>
        /// Wait period in ms between consecutive email sends.
        /// Default: 7000 ms matching the legacy waitPeriodVal.
        /// </summary>
        public int EmailWaitPeriodMs { get; }

        // ── Raw Data Field Separator ─────────────────────────────────────────────

        /// <summary>
        /// Field separator for RawData formatters.
        /// "," for AIBK/EGB/ALXB, "|" for AAIB/BRKA/UNB/VCBK.
        /// </summary>
        public string FieldSeparator { get; }

        // ── Constructor ───────────────────────────────────────────────────────────

        public GenerateStatementCommand(
            int                 branchCode,
            string              bankName,
            string              bankFullName,
            string              cardProduct,
            StatementOutputType outputType,
            string              formatterKey,
            ProcessingMode      processingModes,
            CardType            cardType,
            StatementType       statementType,
            DateTime            statementDate,
            string              outputRootPath,
            bool                appendMode                   = false,
            string              rewardContractCondition      = "'New Reward Contract'",
            string              installmentContractCondition = null,
            bool                useCorporateAccountNumber    = false,
            string              statementTypeSuffix          = "CR",
            string              reportTemplateName           = null,
            string              pdfPasswordScheme            = null,
            string              emailFromAddress             = "cardservices@emp-group.com",
            string              bankWebsiteUrl               = "www.emp-group.com",
            string              facebookUrl                  = null,
            string              linkedInUrl                  = null,
            string              youTubeUrl                   = null,
            int                 maxTransactionsPerPage       = 20,
            int                 maxTransactionsLastPage      = 27,
            int                 emailWaitPeriodMs            = 7000,
            string              fieldSeparator               = ",")
        {
            if (string.IsNullOrWhiteSpace(bankName))
                throw new ArgumentException("BankName is required.", nameof(bankName));
            if (string.IsNullOrWhiteSpace(formatterKey))
                throw new ArgumentException("FormatterKey is required.", nameof(formatterKey));

            BranchCode                   = branchCode;
            BankName                     = bankName;
            BankFullName                 = bankFullName;
            CardProduct                  = cardProduct;
            OutputType                   = outputType;
            FormatterKey                 = formatterKey;
            ProcessingModes              = processingModes;
            CardType                     = cardType;
            StatementType                = statementType;
            StatementDate                = statementDate;
            OutputRootPath               = outputRootPath;
            AppendMode                   = appendMode;
            RewardContractCondition      = rewardContractCondition;
            InstallmentContractCondition = installmentContractCondition;
            UseCorporateAccountNumber    = useCorporateAccountNumber;
            StatementTypeSuffix          = statementTypeSuffix;
            ReportTemplateName           = reportTemplateName;
            PdfPasswordScheme            = pdfPasswordScheme;
            EmailFromAddress             = emailFromAddress;
            BankWebsiteUrl               = bankWebsiteUrl;
            FacebookUrl                  = facebookUrl;
            LinkedInUrl                  = linkedInUrl;
            YouTubeUrl                   = youTubeUrl;
            MaxTransactionsPerPage       = maxTransactionsPerPage;
            MaxTransactionsLastPage      = maxTransactionsLastPage;
            EmailWaitPeriodMs            = emailWaitPeriodMs;
            FieldSeparator               = fieldSeparator;
        }

        /// <summary>Returns true when the given mode flag is active.</summary>
        public bool HasMode(ProcessingMode mode) => (ProcessingModes & mode) == mode;
    }
}

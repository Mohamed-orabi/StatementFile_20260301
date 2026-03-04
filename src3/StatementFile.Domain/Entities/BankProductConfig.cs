using System;
using StatementFile.Domain.Enums;

namespace StatementFile.Domain.Entities
{
    /// <summary>
    /// Represents one row in the STAT_BANK_PRODUCT_CONFIG Oracle table.
    ///
    /// Each row describes a single bank + card-product combination and carries
    /// every property that was previously hardcoded in the 500-case switch statement
    /// inside frmStatementFileExtn.runStatement().
    ///
    /// The <see cref="FormatterKey"/> string (e.g. "HTML_UBA", "PDF_BDCA") is the
    /// lookup key used by <see cref="StatementFile.Infrastructure.Formatters.FormatterRegistry"/>
    /// to resolve the correct <see cref="StatementFile.Domain.Interfaces.IStatementFormatter"/>.
    /// </summary>
    public sealed class BankProductConfig
    {
        // ── Identity ────────────────────────────────────────────────────────────

        public int    Id                    { get; private set; }
        public bool   IsActive              { get; private set; }

        // ── Bank identification ─────────────────────────────────────────────────

        /// <summary>Short mnemonic used in file names (e.g. "UBA", "BDCA").</summary>
        public string BankName              { get; private set; }

        /// <summary>Full legal name printed in emails / reports.</summary>
        public string BankFullName          { get; private set; }

        /// <summary>Numeric bank code stored in Oracle BANKCODE columns.</summary>
        public string BankCode              { get; private set; }

        /// <summary>Physical branch / data-partition code.</summary>
        public int    BranchCode            { get; private set; }

        // ── Product / statement type ────────────────────────────────────────────

        /// <summary>
        /// File-name suffix that identifies the card portfolio and statement variant,
        /// e.g. "CR", "DB", "CP", "CRCORP", "DBREWARD".
        /// Together with <see cref="BranchCode"/> it selects the Oracle table set.
        /// </summary>
        public string StatementTypeSuffix   { get; private set; }

        /// <summary>Card portfolio category — drives master/detail table selection.</summary>
        public CardType CardType            { get; private set; }

        /// <summary>Free-text product label displayed to operators.</summary>
        public string CardProduct           { get; private set; }

        // ── Output / formatter ──────────────────────────────────────────────────

        public StatementOutputType OutputType { get; private set; }

        /// <summary>
        /// Registry key that resolves to a concrete IStatementFormatter.
        /// Convention: "{OUTPUT}_{BANK}", e.g. "HTML_UBA", "PDF_BDCA", "TXT_NSGB".
        /// </summary>
        public string FormatterKey          { get; private set; }

        // ── Branding assets ─────────────────────────────────────────────────────

        public string BankWebLink           { get; private set; }
        public string BankLogo              { get; private set; }
        public string BackgroundImage       { get; private set; }
        public string MidBannerImage        { get; private set; }
        public string BottomBannerImage     { get; private set; }

        // ── Email settings ──────────────────────────────────────────────────────

        public string EmailFromAddress      { get; private set; }
        public string EmailFromName         { get; private set; }

        // ── Oracle query filters ────────────────────────────────────────────────

        /// <summary>SQL WHERE fragment appended to product-specific queries.</summary>
        public string WhereCondition        { get; private set; }

        /// <summary>SQL fragment for VIP-customer filtering.</summary>
        public string VipCondition          { get; private set; }

        /// <summary>SQL fragment identifying reward-eligible records.</summary>
        public string RewardCondition       { get; private set; }

        /// <summary>SQL fragment identifying reward-contract accounts.</summary>
        public string RewardContractCondition { get; private set; }

        /// <summary>SQL fragment for currency-based filtering.</summary>
        public string CurrencyFilter        { get; private set; }

        /// <summary>SQL fragment used when instalment data must be excluded.</summary>
        public string InstallmentCondition  { get; private set; }

        /// <summary>Payment-network system code (e.g. "VISA", "MC").</summary>
        public string PaymentSystem         { get; private set; }

        // ── Processing flags ────────────────────────────────────────────────────

        public ProcessingMode ProcessingModes   { get; private set; }

        public bool IsRewardRun             { get; private set; }
        public bool IsSplitOutput           { get; private set; }
        public bool HasAttachment           { get; private set; }
        public bool SaveDataset             { get; private set; }
        public bool ShowMessageBox          { get; private set; }
        public bool RunNullCardDelete       { get; private set; }
        public bool RunCardBranchMatch      { get; private set; }
        public bool ExcludeReward           { get; private set; }

        /// <summary>
        /// Seconds to pause between batches when splitting a large run.
        /// Mirrors the <c>waitPeriod</c> property in the legacy code.
        /// </summary>
        public int WaitPeriodSeconds        { get; private set; }

        // ── Metadata ────────────────────────────────────────────────────────────

        public DateTime CreatedAt           { get; private set; }
        public DateTime UpdatedAt           { get; private set; }

        // ── Private constructor (use factory) ────────────────────────────────────

        private BankProductConfig() { }

        // ── Factory method ───────────────────────────────────────────────────────

        public static BankProductConfig Create(
            string bankName,
            string bankFullName,
            string bankCode,
            int    branchCode,
            string statementTypeSuffix,
            CardType cardType,
            string cardProduct,
            StatementOutputType outputType,
            string formatterKey,
            string bankWebLink               = null,
            string bankLogo                  = null,
            string backgroundImage           = null,
            string midBannerImage            = null,
            string bottomBannerImage         = null,
            string emailFromAddress          = "cardservices@emp-group.com",
            string emailFromName             = null,
            string whereCondition            = null,
            string vipCondition              = null,
            string rewardCondition           = null,
            string rewardContractCondition   = "'New Reward Contract'",
            string currencyFilter            = null,
            string installmentCondition      = null,
            string paymentSystem             = null,
            ProcessingMode processingModes   = ProcessingMode.None,
            bool   isRewardRun               = false,
            bool   isSplitOutput             = false,
            bool   hasAttachment             = false,
            bool   saveDataset               = false,
            bool   showMessageBox            = false,
            bool   runNullCardDelete         = true,
            bool   runCardBranchMatch        = true,
            bool   excludeReward             = true,
            int    waitPeriodSeconds         = 0)
        {
            if (string.IsNullOrWhiteSpace(bankName))
                throw new ArgumentException("BankName is required.", nameof(bankName));
            if (string.IsNullOrWhiteSpace(formatterKey))
                throw new ArgumentException("FormatterKey is required.", nameof(formatterKey));

            return new BankProductConfig
            {
                IsActive                 = true,
                BankName                 = bankName,
                BankFullName             = bankFullName ?? bankName,
                BankCode                 = bankCode,
                BranchCode               = branchCode,
                StatementTypeSuffix      = statementTypeSuffix,
                CardType                 = cardType,
                CardProduct              = cardProduct,
                OutputType               = outputType,
                FormatterKey             = formatterKey,
                BankWebLink              = bankWebLink,
                BankLogo                 = bankLogo,
                BackgroundImage          = backgroundImage,
                MidBannerImage           = midBannerImage,
                BottomBannerImage        = bottomBannerImage,
                EmailFromAddress         = emailFromAddress,
                EmailFromName            = emailFromName ?? bankFullName ?? bankName,
                WhereCondition           = whereCondition,
                VipCondition             = vipCondition,
                RewardCondition          = rewardCondition,
                RewardContractCondition  = rewardContractCondition ?? "'New Reward Contract'",
                CurrencyFilter           = currencyFilter,
                InstallmentCondition     = installmentCondition,
                PaymentSystem            = paymentSystem,
                ProcessingModes          = processingModes,
                IsRewardRun              = isRewardRun,
                IsSplitOutput            = isSplitOutput,
                HasAttachment            = hasAttachment,
                SaveDataset              = saveDataset,
                ShowMessageBox           = showMessageBox,
                RunNullCardDelete        = runNullCardDelete,
                RunCardBranchMatch       = runCardBranchMatch,
                ExcludeReward            = excludeReward,
                WaitPeriodSeconds        = waitPeriodSeconds,
                CreatedAt                = DateTime.UtcNow,
                UpdatedAt                = DateTime.UtcNow,
            };
        }

        // ── Update method ────────────────────────────────────────────────────────

        public void Update(
            string bankName,
            string bankFullName,
            string bankCode,
            int    branchCode,
            string statementTypeSuffix,
            CardType cardType,
            string cardProduct,
            StatementOutputType outputType,
            string formatterKey,
            string bankWebLink,
            string bankLogo,
            string backgroundImage,
            string midBannerImage,
            string bottomBannerImage,
            string emailFromAddress,
            string emailFromName,
            string whereCondition,
            string vipCondition,
            string rewardCondition,
            string rewardContractCondition,
            string currencyFilter,
            string installmentCondition,
            string paymentSystem,
            ProcessingMode processingModes,
            bool   isRewardRun,
            bool   isSplitOutput,
            bool   hasAttachment,
            bool   saveDataset,
            bool   showMessageBox,
            bool   runNullCardDelete,
            bool   runCardBranchMatch,
            bool   excludeReward,
            int    waitPeriodSeconds,
            bool   isActive)
        {
            BankName                 = bankName;
            BankFullName             = bankFullName;
            BankCode                 = bankCode;
            BranchCode               = branchCode;
            StatementTypeSuffix      = statementTypeSuffix;
            CardType                 = cardType;
            CardProduct              = cardProduct;
            OutputType               = outputType;
            FormatterKey             = formatterKey;
            BankWebLink              = bankWebLink;
            BankLogo                 = bankLogo;
            BackgroundImage          = backgroundImage;
            MidBannerImage           = midBannerImage;
            BottomBannerImage        = bottomBannerImage;
            EmailFromAddress         = emailFromAddress;
            EmailFromName            = emailFromName;
            WhereCondition           = whereCondition;
            VipCondition             = vipCondition;
            RewardCondition          = rewardCondition;
            RewardContractCondition  = rewardContractCondition;
            CurrencyFilter           = currencyFilter;
            InstallmentCondition     = installmentCondition;
            PaymentSystem            = paymentSystem;
            ProcessingModes          = processingModes;
            IsRewardRun              = isRewardRun;
            IsSplitOutput            = isSplitOutput;
            HasAttachment            = hasAttachment;
            SaveDataset              = saveDataset;
            ShowMessageBox           = showMessageBox;
            RunNullCardDelete        = runNullCardDelete;
            RunCardBranchMatch       = runCardBranchMatch;
            ExcludeReward            = excludeReward;
            WaitPeriodSeconds        = waitPeriodSeconds;
            IsActive                 = isActive;
            UpdatedAt                = DateTime.UtcNow;
        }

        public void Deactivate()
        {
            IsActive  = false;
            UpdatedAt = DateTime.UtcNow;
        }
    }
}

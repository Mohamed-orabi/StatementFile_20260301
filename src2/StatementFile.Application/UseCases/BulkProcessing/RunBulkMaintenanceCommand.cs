namespace StatementFile.Application.UseCases.BulkProcessing
{
    /// <summary>
    /// Instructs the bulk maintenance use case to run the pre-generation data
    /// fixes for one bank branch.
    ///
    /// All operations are optional flags; the caller sets them based on the
    /// product configuration for the branch.
    ///
    /// Operation mapping to legacy clsMaintainData methods:
    ///   RunNullCardDelete       → CleanNullCards()
    ///   ExcludeReward           → excludes contracttype = RewardContractCondition
    ///   ExcludeInstallment      → excludes contracttype IN InstallmentCondition
    ///   RunCardBranchMatch      → MatchCardBranchForAccount()
    ///   RunArabicAddressFix     → FixArabicAddress()          [strips "???" prefix]
    ///   RunFixAddress           → FixAddress()                [split >50-char addresses]
    ///   RunFixArabicAddressLang → FixArabicAddressLang()      [set companycode=1/0]
    ///   RunOnHoldDelete         → DeleteOnHoldTransactions()  [POSTINGDATE IS NULL AND DOCNO IS NULL]
    ///   IsRewardRun             → passed as isReward=true to DeleteOnHoldTransactions
    ///   RunMergeMarkUpFees      → MergeMarkUpFees()           [consolidate Mark-Up Fee rows]
    ///   RunRewardFix            → FixReward()
    /// </summary>
    public sealed class RunBulkMaintenanceCommand
    {
        public int BranchCode { get; }

        // ── NULL-card cleanup ─────────────────────────────────────────────────────
        public bool   RunNullCardDelete       { get; }
        public bool   ExcludeReward           { get; }
        public bool   ExcludeInstallment      { get; }
        public string InstallmentCondition    { get; }
        public string RewardContractCondition { get; }

        // ── Card-branch alignment ─────────────────────────────────────────────────
        public bool RunCardBranchMatch { get; }

        // ── Arabic address corruption fix ─────────────────────────────────────────
        /// <summary>
        /// Strips "???" prefix from all 9 address fields.
        /// Legacy: clsMaintainData.fixArbicAddress()
        /// </summary>
        public bool RunArabicAddressFix { get; }

        // ── Long address split ────────────────────────────────────────────────────
        /// <summary>
        /// Splits customeraddress1 longer than 50 chars at a word boundary.
        /// Applied to customer, account, and card address fields.
        /// Legacy: clsMaintainData.fixAddress()
        /// </summary>
        public bool RunFixAddress { get; }

        // ── Arabic language detection ─────────────────────────────────────────────
        /// <summary>
        /// Sets companycode = 1 (Arabic) or 0 (non-Arabic) for each statement row.
        /// Legacy: clsMaintainData.fixArbicAddressLang()
        /// </summary>
        public bool RunFixArabicAddressLang { get; }

        // ── On-hold transaction delete ────────────────────────────────────────────
        /// <summary>
        /// Deletes detail rows WHERE POSTINGDATE IS NULL AND DOCNO IS NULL.
        /// Legacy: clsMaintainData.deleteOnHoldTrans(int pBranch, bool isReward)
        /// </summary>
        public bool RunOnHoldDelete { get; }

        /// <summary>
        /// When true, also excludes 'Calculated Points' rows from the on-hold delete.
        /// Passed as isReward=true to DeleteOnHoldTransactions().
        /// </summary>
        public bool IsRewardRun { get; }

        // ── Mark-Up Fee merge ─────────────────────────────────────────────────────
        /// <summary>
        /// Consolidates Mark-Up Fee transactions in the detail table by docno.
        /// Legacy: clsMaintainData.mergeMarkUpFees()
        /// </summary>
        public bool RunMergeMarkUpFees { get; }

        // ── Reward programme fix ──────────────────────────────────────────────────
        public bool RunRewardFix { get; }

        public RunBulkMaintenanceCommand(
            int    branchCode,
            bool   runNullCardDelete       = true,
            bool   excludeReward           = true,
            bool   excludeInstallment      = false,
            string installmentCondition    = null,
            string rewardContractCondition = "'Reward Program (Airmile)'",
            bool   runCardBranchMatch      = true,
            bool   runArabicAddressFix     = false,
            bool   runFixAddress           = false,
            bool   runFixArabicAddressLang = false,
            bool   runOnHoldDelete         = false,
            bool   isRewardRun             = false,
            bool   runMergeMarkUpFees      = false,
            bool   runRewardFix            = false)
        {
            BranchCode              = branchCode;
            RunNullCardDelete       = runNullCardDelete;
            ExcludeReward           = excludeReward;
            ExcludeInstallment      = excludeInstallment;
            InstallmentCondition    = installmentCondition;
            RewardContractCondition = rewardContractCondition;
            RunCardBranchMatch      = runCardBranchMatch;
            RunArabicAddressFix     = runArabicAddressFix;
            RunFixAddress           = runFixAddress;
            RunFixArabicAddressLang = runFixArabicAddressLang;
            RunOnHoldDelete         = runOnHoldDelete;
            IsRewardRun             = isRewardRun;
            RunMergeMarkUpFees      = runMergeMarkUpFees;
            RunRewardFix            = runRewardFix;
        }
    }
}

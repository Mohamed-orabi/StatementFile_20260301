namespace StatementFile.Application.UseCases.BulkProcessing
{
    /// <summary>
    /// Instructs the bulk maintenance use case to run the pre-generation data
    /// fixes for one bank branch.
    ///
    /// All operations are optional flags; the caller (frmGenerateStatement)
    /// sets them based on the product configuration for the branch.
    ///
    /// Operation mapping to legacy clsMaintainData methods:
    ///   RunNullCardDelete      → CleanNullCards()
    ///   ExcludeReward          → excludes contracttype = RewardContractCondition
    ///   ExcludeInstallment     → excludes contracttype IN InstallmentCondition
    ///   RunCardBranchMatch     → MatchCardBranchForAccount()
    ///   RunArabicAddressFix    → FixArabicAddress()
    ///   RunOnHoldDelete        → DeleteOnHoldTransactions()
    ///   RunRewardFix           → FixReward()  [resets reward flags before generation]
    /// </summary>
    public sealed class RunBulkMaintenanceCommand
    {
        public int BranchCode { get; }

        // ── NULL-card cleanup ─────────────────────────────────────────────────────

        /// <summary>
        /// Execute the NULL-card row delete pass.
        /// Should be true for all credit/debit/prepaid statement runs.
        /// Skipped only when the branch has no card-number requirement.
        /// </summary>
        public bool RunNullCardDelete { get; }

        /// <summary>
        /// Exclude rows whose contracttype matches the reward condition from the delete.
        /// Set to true when the batch also includes reward-programme statements.
        /// Legacy: notRwardVal = false in clsMaintainData.
        /// </summary>
        public bool ExcludeReward { get; }

        /// <summary>
        /// Oracle NOT IN condition for installment contract types.
        /// e.g. "('Purchase Installment With Interest Rate','BuyNow Installment')"
        /// Used only when ExcludeInstallment is true.
        /// </summary>
        public string InstallmentCondition { get; }

        /// <summary>
        /// Exclude rows whose contracttype is in InstallmentCondition from the delete.
        /// Set to true for branches that process instalment statements.
        /// </summary>
        public bool ExcludeInstallment { get; }

        /// <summary>
        /// Oracle = condition value for the reward contract type filter,
        /// e.g. "'Reward Program (Airmile)'" or "'New Reward Contract'".
        /// </summary>
        public string RewardContractCondition { get; }

        // ── Card-branch alignment ─────────────────────────────────────────────────

        /// <summary>
        /// Run the card-branch-part alignment pass.
        /// Aligns cardbranchpart / cardbranchpartname across all cards for each client
        /// to the most recently created card's branch.
        /// Should be true for most production runs.
        /// Legacy: clsMaintainData.matchCardBranch4Account().
        /// </summary>
        public bool RunCardBranchMatch { get; }

        // ── Arabic address fix ────────────────────────────────────────────────────

        /// <summary>
        /// Run the Arabic address corruption fix.
        /// Detects "???" prefix on address fields and strips it.
        /// Required for: Egyptian banks (AAIB, CMB, ALXB, QNB, BDCA, AIBK, AUB).
        /// Legacy: clsMaintainData.fixArbicAddress().
        /// </summary>
        public bool RunArabicAddressFix { get; }

        // ── On-hold transaction delete ────────────────────────────────────────────

        /// <summary>
        /// Delete detail rows where HOLSTMT = 'Y' before generation.
        /// Required for: AUB (Branch 25), any bank that sets HOLSTMT flags.
        /// Legacy: clsMaintainData.deleteOnHoldTrans().
        /// </summary>
        public bool RunOnHoldDelete { get; }

        // ── Reward programme fix ──────────────────────────────────────────────────

        /// <summary>
        /// Run the reward-programme pre-fix before generating reward statements.
        /// Ensures the reward programme fields are consistent before the formatter
        /// reads them.
        /// Only applies when the run includes reward-programme products.
        /// Legacy: clsMaintainData.fixReward().
        /// </summary>
        public bool RunRewardFix { get; }

        public RunBulkMaintenanceCommand(
            int    branchCode,
            bool   runNullCardDelete      = true,
            bool   excludeReward          = true,
            bool   excludeInstallment     = false,
            string installmentCondition   = null,
            string rewardContractCondition = "'Reward Program (Airmile)'",
            bool   runCardBranchMatch     = true,
            bool   runArabicAddressFix    = false,
            bool   runOnHoldDelete        = false,
            bool   runRewardFix           = false)
        {
            BranchCode              = branchCode;
            RunNullCardDelete       = runNullCardDelete;
            ExcludeReward           = excludeReward;
            ExcludeInstallment      = excludeInstallment;
            InstallmentCondition    = installmentCondition;
            RewardContractCondition = rewardContractCondition;
            RunCardBranchMatch      = runCardBranchMatch;
            RunArabicAddressFix     = runArabicAddressFix;
            RunOnHoldDelete         = runOnHoldDelete;
            RunRewardFix            = runRewardFix;
        }
    }
}

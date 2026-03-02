namespace StatementFile.Application.UseCases.BulkProcessing
{
    /// <summary>
    /// Command that instructs the bulk-maintenance use case to run
    /// pre-statement-generation data fixes for a given bank branch.
    /// </summary>
    public sealed class RunBulkMaintenanceCommand
    {
        public int    BranchCode           { get; }

        /// <summary>When true, rows with contracttype = 'Reward Program (Airmile)' are excluded from the NULL-card delete.</summary>
        public bool   ExcludeReward        { get; }

        /// <summary>When true, rows matching the installment contract list are excluded from the NULL-card delete.</summary>
        public bool   ExcludeInstallment   { get; }

        /// <summary>IN-list SQL fragment for installment contract types, e.g. "('Installment','BuyNow')".</summary>
        public string InstallmentCondition { get; }

        /// <summary>Run Arabic-address fix pass in addition to card-branch matching.</summary>
        public bool   RunArabicAddressFix  { get; }

        public RunBulkMaintenanceCommand(
            int    branchCode,
            bool   excludeReward        = true,
            bool   excludeInstallment   = false,
            string installmentCondition = null,
            bool   runArabicAddressFix  = false)
        {
            BranchCode           = branchCode;
            ExcludeReward        = excludeReward;
            ExcludeInstallment   = excludeInstallment;
            InstallmentCondition = installmentCondition;
            RunArabicAddressFix  = runArabicAddressFix;
        }
    }
}

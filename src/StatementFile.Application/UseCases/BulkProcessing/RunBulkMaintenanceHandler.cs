using System;
using StatementFile.Domain.Common;
using StatementFile.Domain.Interfaces.Services;

namespace StatementFile.Application.UseCases.BulkProcessing
{
    /// <summary>
    /// Orchestrates ALL pre-statement-generation maintenance steps for one branch.
    ///
    /// Execution order mirrors the legacy clsMaintainData call sequence:
    ///
    ///   1. Delete on-hold transactions (POSTINGDATE IS NULL AND DOCNO IS NULL)
    ///      → clsMaintainData.deleteOnHoldTrans(int, bool)
    ///
    ///   2. Merge Mark-Up Fee transactions by docno
    ///      → clsMaintainData.mergeMarkUpFees()
    ///
    ///   3. Delete NULL-card rows (excluding reward / installment contracts)
    ///      → clsMaintainData.CleanNullCards()
    ///
    ///   4. Card-branch-part alignment
    ///      → clsMaintainData.matchCardBranch4Account()
    ///
    ///   5. Arabic address corruption fix (strip "???" prefix, 9 fields)
    ///      → clsMaintainData.fixArbicAddress()
    ///
    ///   6. Long address split (>50 chars split into addr1 + addr2)
    ///      → clsMaintainData.fixAddress()
    ///
    ///   7. Arabic language code assignment (companycode = 1 or 0)
    ///      → clsMaintainData.fixArbicAddressLang()
    ///
    ///   8. Reward programme pre-fix
    ///      → clsMaintainData.fixReward()
    ///
    /// All steps are optional flags; the command encodes which steps apply.
    /// Returns a summary result with row counts for each completed step.
    /// </summary>
    public sealed class RunBulkMaintenanceHandler
    {
        private readonly IDataMaintenanceService _maintenance;

        public RunBulkMaintenanceHandler(IDataMaintenanceService maintenance)
        {
            _maintenance = maintenance ?? throw new ArgumentNullException(nameof(maintenance));
        }

        public Result<BulkMaintenanceResult> Handle(RunBulkMaintenanceCommand command)
        {
            try
            {
                var result = new BulkMaintenanceResult { BranchCode = command.BranchCode };

                // ── Step 1: On-hold transaction delete ─────────────────────────────
                if (command.RunOnHoldDelete)
                {
                    result.OnHoldRowsDeleted = _maintenance.DeleteOnHoldTransactions(
                        command.BranchCode,
                        command.IsRewardRun);
                }

                // ── Step 2: Merge Mark-Up Fee transactions ─────────────────────────
                if (command.RunMergeMarkUpFees)
                {
                    _maintenance.MergeMarkUpFees(command.BranchCode);
                }

                // ── Step 3: NULL-card row delete ───────────────────────────────────
                if (command.RunNullCardDelete)
                {
                    result.NullCardsDeleted = _maintenance.CleanNullCards(
                        command.BranchCode,
                        command.ExcludeReward,
                        command.ExcludeInstallment,
                        command.InstallmentCondition);
                }

                // ── Step 4: Card-branch-part alignment ─────────────────────────────
                if (command.RunCardBranchMatch)
                {
                    result.CardBranchRecordsUpdated = _maintenance.MatchCardBranchForAccount(
                        command.BranchCode);
                }

                // ── Step 5: Arabic address corruption fix ──────────────────────────
                if (command.RunArabicAddressFix)
                {
                    result.ArabicAddressRecordsFixed = _maintenance.FixArabicAddress(
                        command.BranchCode);
                }

                // ── Step 6: Long address split ─────────────────────────────────────
                if (command.RunFixAddress)
                {
                    result.AddressesSplit = _maintenance.FixAddress(command.BranchCode);
                }

                // ── Step 7: Arabic language code assignment ────────────────────────
                if (command.RunFixArabicAddressLang)
                {
                    result.ArabicLangRowsUpdated = _maintenance.FixArabicAddressLang(
                        command.BranchCode);
                }

                // ── Step 8: Reward programme pre-fix ──────────────────────────────
                if (command.RunRewardFix)
                {
                    result.RewardRowsFixed = _maintenance.FixReward(
                        command.BranchCode,
                        command.RewardContractCondition);
                }

                return Result.Ok(result);
            }
            catch (Exception ex)
            {
                return Result.Fail<BulkMaintenanceResult>(
                    $"Bulk maintenance failed for branch {command.BranchCode}: {ex.Message}");
            }
        }
    }

    public sealed class BulkMaintenanceResult
    {
        public int BranchCode                { get; set; }
        public int OnHoldRowsDeleted         { get; set; }
        public int NullCardsDeleted          { get; set; }
        public int CardBranchRecordsUpdated  { get; set; }
        public int ArabicAddressRecordsFixed { get; set; }
        public int AddressesSplit            { get; set; }
        public int ArabicLangRowsUpdated     { get; set; }
        public int RewardRowsFixed           { get; set; }
    }
}

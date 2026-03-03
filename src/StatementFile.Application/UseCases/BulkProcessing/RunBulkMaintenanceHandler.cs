using System;
using StatementFile.Domain.Common;
using StatementFile.Domain.Interfaces.Services;

namespace StatementFile.Application.UseCases.BulkProcessing
{
    /// <summary>
    /// Orchestrates ALL pre-statement-generation maintenance steps for one branch.
    ///
    /// Execution order is fixed and mirrors the legacy clsMaintainData call sequence
    /// found in clsBasStatementHtml.Statement() and the individual bank Statement()
    /// overrides in the Banks/ folder:
    ///
    ///   1. Delete on-hold transactions (HOLSTMT = 'Y') if requested
    ///      → clsMaintainData.deleteOnHoldTrans()  [AUB, Branch 25]
    ///
    ///   2. Delete NULL-card rows (excluding reward / installment contracts)
    ///      → clsMaintainData.CleanNullCards()  [all credit/debit/prepaid runs]
    ///
    ///   3. Card-branch-part alignment
    ///      → clsMaintainData.matchCardBranch4Account()  [virtually all runs]
    ///
    ///   4. Arabic address corruption fix
    ///      → clsMaintainData.fixArbicAddress()  [Egyptian banks]
    ///
    ///   5. Reward programme pre-fix
    ///      → clsMaintainData.fixReward()  [runs with Reward mode enabled]
    ///
    /// All steps are optional flags; the command encodes which steps apply
    /// to the current bank/product combination.
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
                        command.BranchCode);
                }

                // ── Step 2: NULL-card row delete ───────────────────────────────────
                if (command.RunNullCardDelete)
                {
                    result.NullCardsDeleted = _maintenance.CleanNullCards(
                        command.BranchCode,
                        command.ExcludeReward,
                        command.ExcludeInstallment,
                        command.InstallmentCondition);
                }

                // ── Step 3: Card-branch-part alignment ─────────────────────────────
                if (command.RunCardBranchMatch)
                {
                    result.CardBranchRecordsUpdated = _maintenance.MatchCardBranchForAccount(
                        command.BranchCode);
                }

                // ── Step 4: Arabic address fix ─────────────────────────────────────
                if (command.RunArabicAddressFix)
                {
                    result.ArabicAddressRecordsFixed = _maintenance.FixArabicAddress(
                        command.BranchCode);
                }

                // ── Step 5: Reward programme pre-fix ──────────────────────────────
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
        public int RewardRowsFixed           { get; set; }
    }
}

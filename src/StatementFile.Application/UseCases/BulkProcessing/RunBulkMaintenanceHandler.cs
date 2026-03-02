using System;
using StatementFile.Domain.Common;
using StatementFile.Domain.Interfaces.Services;

namespace StatementFile.Application.UseCases.BulkProcessing
{
    /// <summary>
    /// Orchestrates the bulk pre-processing maintenance steps that must run
    /// before statement generation for any branch.
    ///
    /// Steps (order preserved from legacy clsMaintainData):
    ///   1. Delete NULL-card rows (excluding reward / installment as configured)
    ///   2. Match card-branch parts across all accounts for the branch
    ///   3. Optionally fix garbled Arabic address fields
    ///
    /// Returns a summary result with row counts per step.
    /// </summary>
    public sealed class RunBulkMaintenanceHandler
    {
        private readonly IDataMaintenanceService _maintenanceService;

        public RunBulkMaintenanceHandler(IDataMaintenanceService maintenanceService)
        {
            _maintenanceService = maintenanceService
                ?? throw new ArgumentNullException(nameof(maintenanceService));
        }

        public Result<BulkMaintenanceResult> Handle(RunBulkMaintenanceCommand command)
        {
            try
            {
                var summary = new BulkMaintenanceResult { BranchCode = command.BranchCode };

                // Step 1: Clean null-card rows
                summary.NullCardsDeleted = _maintenanceService.CleanNullCards(
                    command.BranchCode,
                    command.ExcludeReward,
                    command.ExcludeInstallment,
                    command.InstallmentCondition);

                // Step 2: Card-branch matching
                summary.CardBranchRecordsUpdated = _maintenanceService.MatchCardBranchForAccount(
                    command.BranchCode);

                // Step 3: Arabic address fix (optional)
                if (command.RunArabicAddressFix)
                {
                    summary.ArabicAddressRecordsFixed = _maintenanceService.FixArabicAddress(
                        command.BranchCode);
                }

                return Result.Ok(summary);
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
        public int BranchCode                  { get; set; }
        public int NullCardsDeleted            { get; set; }
        public int CardBranchRecordsUpdated    { get; set; }
        public int ArabicAddressRecordsFixed   { get; set; }
    }
}

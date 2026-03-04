using System;
using StatementFile.Domain.Enums;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Application.Commands
{
    /// <summary>
    /// Executes the bulk data-maintenance Oracle operations that frmStatementFile
    /// ran before every statement generation cycle.
    ///
    /// Delegates each maintenance step to <see cref="IBulkMaintenanceService"/>
    /// which is implemented in the Infrastructure layer (Oracle procedures).
    /// </summary>
    public sealed class RunMaintenanceHandler
    {
        private readonly IBulkMaintenanceService _maintenance;

        public RunMaintenanceHandler(IBulkMaintenanceService maintenance)
        {
            _maintenance = maintenance ?? throw new ArgumentNullException(nameof(maintenance));
        }

        public MaintenanceResult Handle(RunMaintenanceCommand cmd)
        {
            if (cmd == null) throw new ArgumentNullException(nameof(cmd));

            try
            {
                var modes = cmd.ProcessingModes;

                if (cmd.RunNullCardDelete)
                    _maintenance.DeleteNullCardRecords(cmd.BranchCode, cmd.ConnectionString);

                if (cmd.RunCardBranchMatch)
                    _maintenance.MatchCardBranch(cmd.BranchCode, cmd.ConnectionString);

                if (modes.HasFlag(ProcessingMode.FixArabicAddress))
                    _maintenance.FixArabicAddress(cmd.BranchCode, cmd.ConnectionString);

                if (modes.HasFlag(ProcessingMode.FixArabicAddressLang))
                    _maintenance.FixArabicAddressLanguage(cmd.BranchCode, cmd.ConnectionString);

                if (modes.HasFlag(ProcessingMode.FixAddress))
                    _maintenance.FixAddress(cmd.BranchCode, cmd.ConnectionString);

                if (modes.HasFlag(ProcessingMode.DeleteOnHold))
                    _maintenance.DeleteOnHoldRecords(cmd.BranchCode, cmd.ConnectionString);

                if (modes.HasFlag(ProcessingMode.MergeMarkUpFees))
                    _maintenance.MergeMarkUpFees(cmd.BranchCode, cmd.ConnectionString);

                if (modes.HasFlag(ProcessingMode.Reward))
                    _maintenance.ProcessRewardData(
                        cmd.BranchCode,
                        cmd.RewardContractCondition ?? "'New Reward Contract'",
                        cmd.ConnectionString);

                if (modes.HasFlag(ProcessingMode.Installment) && !string.IsNullOrWhiteSpace(cmd.InstallmentCondition))
                    _maintenance.ExcludeInstallmentData(
                        cmd.BranchCode,
                        cmd.InstallmentCondition,
                        cmd.ConnectionString);

                return MaintenanceResult.Ok();
            }
            catch (Exception ex)
            {
                return MaintenanceResult.Failure(ex.Message);
            }
        }
    }

    public sealed class MaintenanceResult
    {
        public bool   IsSuccess { get; private set; }
        public string Error     { get; private set; }

        public static MaintenanceResult Ok()              => new MaintenanceResult { IsSuccess = true };
        public static MaintenanceResult Failure(string e) => new MaintenanceResult { IsSuccess = false, Error = e };
    }
}

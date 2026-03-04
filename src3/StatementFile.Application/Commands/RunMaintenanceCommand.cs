using StatementFile.Domain.Enums;

namespace StatementFile.Application.Commands
{
    /// <summary>
    /// Drives the bulk data-maintenance routines that frmStatementFile ran before each generation.
    /// The handler executes Oracle stored procedures for null-card delete, address fixes, etc.
    /// </summary>
    public sealed class RunMaintenanceCommand
    {
        public int            BranchCode                { get; init; }
        public ProcessingMode ProcessingModes           { get; init; }
        public bool           RunNullCardDelete         { get; init; }
        public bool           ExcludeReward             { get; init; }
        public bool           RunCardBranchMatch        { get; init; }
        public string         RewardContractCondition   { get; init; }
        public string         InstallmentCondition      { get; init; }
        public string         ConnectionString          { get; init; }
    }
}

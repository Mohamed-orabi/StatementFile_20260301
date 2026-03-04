namespace StatementFile.Application.DTOs
{
    /// <summary>
    /// Request body for POST /api/maintenance/run.
    /// Replaces the direct call to bulk maintenance routines in frmStatementFile.
    /// </summary>
    public sealed class MaintenanceRequest
    {
        public int    BranchCode                { get; set; }
        public long   ProcessingModes           { get; set; }
        public bool   RunNullCardDelete         { get; set; } = true;
        public bool   ExcludeReward             { get; set; } = true;
        public bool   RunCardBranchMatch        { get; set; } = true;
        public string RewardContractCondition   { get; set; } = "'New Reward Contract'";
        public string InstallmentCondition      { get; set; }
    }
}

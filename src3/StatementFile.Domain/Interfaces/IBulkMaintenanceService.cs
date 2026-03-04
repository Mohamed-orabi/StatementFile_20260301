namespace StatementFile.Domain.Interfaces
{
    /// <summary>
    /// Bulk data-maintenance operations executed before statement generation.
    /// Each method maps to an Oracle stored procedure or SQL statement that was
    /// previously called directly from frmStatementFile.
    /// Implemented in <see cref="StatementFile.Infrastructure.Oracle.OracleBulkMaintenanceService"/>.
    /// </summary>
    public interface IBulkMaintenanceService
    {
        void DeleteNullCardRecords(int branchCode, string connectionString);
        void MatchCardBranch(int branchCode, string connectionString);
        void FixArabicAddress(int branchCode, string connectionString);
        void FixArabicAddressLanguage(int branchCode, string connectionString);
        void FixAddress(int branchCode, string connectionString);
        void DeleteOnHoldRecords(int branchCode, string connectionString);
        void MergeMarkUpFees(int branchCode, string connectionString);
        void ProcessRewardData(int branchCode, string rewardContractCondition, string connectionString);
        void ExcludeInstallmentData(int branchCode, string installmentCondition, string connectionString);
    }
}

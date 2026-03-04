using System;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Infrastructure.Oracle
{
    /// <summary>
    /// Executes the Oracle stored procedures and SQL commands that frmStatementFile
    /// ran before each statement generation cycle.
    ///
    /// Each method uses a short-lived connection created from the supplied
    /// connection string so that maintenance operations are independent of the
    /// main singleton factory connection.
    /// </summary>
    public sealed class OracleBulkMaintenanceService : IBulkMaintenanceService
    {
        public void DeleteNullCardRecords(int branchCode, string cs)
        {
            ExecProc(cs, "STMT.ZM_STMT_APP.DeleteNullCardRecords",
                p => { AddIn(p, "p_branch_code", branchCode); });
        }

        public void MatchCardBranch(int branchCode, string cs)
        {
            ExecProc(cs, "STMT.ZM_STMT_APP.MatchCardBranch",
                p => { AddIn(p, "p_branch_code", branchCode); });
        }

        public void FixArabicAddress(int branchCode, string cs)
        {
            ExecProc(cs, "STMT.ZM_STMT_APP.FixArabicAddress",
                p => { AddIn(p, "p_branch_code", branchCode); });
        }

        public void FixArabicAddressLanguage(int branchCode, string cs)
        {
            ExecProc(cs, "STMT.ZM_STMT_APP.FixArabicAddressLanguage",
                p => { AddIn(p, "p_branch_code", branchCode); });
        }

        public void FixAddress(int branchCode, string cs)
        {
            ExecProc(cs, "STMT.ZM_STMT_APP.FixAddress",
                p => { AddIn(p, "p_branch_code", branchCode); });
        }

        public void DeleteOnHoldRecords(int branchCode, string cs)
        {
            ExecProc(cs, "STMT.ZM_STMT_APP.DeleteOnHoldRecords",
                p => { AddIn(p, "p_branch_code", branchCode); });
        }

        public void MergeMarkUpFees(int branchCode, string cs)
        {
            ExecProc(cs, "STMT.ZM_STMT_APP.MergeMarkUpFees",
                p => { AddIn(p, "p_branch_code", branchCode); });
        }

        public void ProcessRewardData(int branchCode, string rewardContractCondition, string cs)
        {
            ExecProc(cs, "STMT.ZM_STMT_APP.ProcessRewardData", p =>
            {
                AddIn(p, "p_branch_code",              branchCode);
                AddIn(p, "p_reward_contract_condition", rewardContractCondition);
            });
        }

        public void ExcludeInstallmentData(int branchCode, string installmentCondition, string cs)
        {
            ExecProc(cs, "STMT.ZM_STMT_APP.ExcludeInstallmentData", p =>
            {
                AddIn(p, "p_branch_code",          branchCode);
                AddIn(p, "p_installment_condition", installmentCondition);
            });
        }

        // ── Helper ───────────────────────────────────────────────────────────────

        private static void ExecProc(string cs, string procName, Action<OracleParameterCollection> bind)
        {
            using var conn = new OracleConnection(cs);
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = procName;
            bind(cmd.Parameters);
            cmd.ExecuteNonQuery();
        }

        private static void AddIn(OracleParameterCollection p, string name, object value)
        {
            p.Add(new OracleParameter(name, value ?? DBNull.Value)
            {
                Direction = ParameterDirection.Input
            });
        }
    }
}

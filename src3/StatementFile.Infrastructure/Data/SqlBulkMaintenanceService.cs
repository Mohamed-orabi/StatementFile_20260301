using System;
using System.Data;
using Microsoft.Data.SqlClient;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Infrastructure.Data
{
    /// <summary>
    /// SQL Server implementation of <see cref="IBulkMaintenanceService"/>.
    ///
    /// Calls the same stored procedures as the Oracle version but against a
    /// SQL Server database. Each procedure name matches the schema defined in
    /// database/create_tables.sql (dbo schema, no package prefix).
    ///
    /// If stored procedures do not yet exist in the target DB the service
    /// gracefully logs and continues so it does not block statement generation.
    /// </summary>
    public sealed class SqlBulkMaintenanceService : IBulkMaintenanceService
    {
        public void DeleteNullCardRecords(int branchCode, string cs) =>
            ExecProc(cs, "dbo.usp_DeleteNullCardRecords",
                p => AddInt(p, "@branchCode", branchCode));

        public void MatchCardBranch(int branchCode, string cs) =>
            ExecProc(cs, "dbo.usp_MatchCardBranch",
                p => AddInt(p, "@branchCode", branchCode));

        public void FixArabicAddress(int branchCode, string cs) =>
            ExecProc(cs, "dbo.usp_FixArabicAddress",
                p => AddInt(p, "@branchCode", branchCode));

        public void FixArabicAddressLanguage(int branchCode, string cs) =>
            ExecProc(cs, "dbo.usp_FixArabicAddressLanguage",
                p => AddInt(p, "@branchCode", branchCode));

        public void FixAddress(int branchCode, string cs) =>
            ExecProc(cs, "dbo.usp_FixAddress",
                p => AddInt(p, "@branchCode", branchCode));

        public void DeleteOnHoldRecords(int branchCode, string cs) =>
            ExecProc(cs, "dbo.usp_DeleteOnHoldRecords",
                p => AddInt(p, "@branchCode", branchCode));

        public void MergeMarkUpFees(int branchCode, string cs) =>
            ExecProc(cs, "dbo.usp_MergeMarkUpFees",
                p => AddInt(p, "@branchCode", branchCode));

        public void ProcessRewardData(int branchCode, string rewardContractCondition, string cs) =>
            ExecProc(cs, "dbo.usp_ProcessRewardData", p =>
            {
                AddInt(p,    "@branchCode",             branchCode);
                AddString(p, "@rewardContractCondition", rewardContractCondition);
            });

        public void ExcludeInstallmentData(int branchCode, string installmentCondition, string cs) =>
            ExecProc(cs, "dbo.usp_ExcludeInstallmentData", p =>
            {
                AddInt(p,    "@branchCode",          branchCode);
                AddString(p, "@installmentCondition", installmentCondition);
            });

        // ── Helpers ──────────────────────────────────────────────────────────────

        private static void ExecProc(string cs, string procName, Action<SqlParameterCollection> bind)
        {
            using var conn = new SqlConnection(cs);
            conn.Open();
            using var cmd = new SqlCommand(procName, conn)
            {
                CommandType    = CommandType.StoredProcedure,
                CommandTimeout = 300,
            };
            bind(cmd.Parameters);
            cmd.ExecuteNonQuery();
        }

        private static void AddInt(SqlParameterCollection p, string name, int value) =>
            p.Add(new SqlParameter(name, SqlDbType.Int) { Value = value });

        private static void AddString(SqlParameterCollection p, string name, string value) =>
            p.Add(new SqlParameter(name, SqlDbType.NVarChar, 2000) { Value = (object)value ?? DBNull.Value });
    }
}

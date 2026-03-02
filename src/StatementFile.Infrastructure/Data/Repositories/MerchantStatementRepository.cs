using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using StatementFile.Domain.Entities;
using StatementFile.Domain.Interfaces.Repositories;

namespace StatementFile.Infrastructure.Data.Repositories
{
    /// <summary>
    /// MS-Access (OleDb) implementation of <see cref="IMerchantStatementRepository"/>.
    /// Persists merchant statement aggregates into the .mdb staging database,
    /// mirroring the exact insert/fixup logic from the legacy clsStatementMrchXml.
    /// </summary>
    public sealed class MerchantStatementRepository : IMerchantStatementRepository
    {
        private static string BuildConnectionString(string mdbPath)
            => $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={mdbPath};";

        public void SaveMerchantStatements(
            IEnumerable<MerchantStatement> statements,
            string                         mdbFilePath)
        {
            using (var conn = new OleDbConnection(BuildConnectionString(mdbFilePath)))
            {
                conn.Open();
                foreach (var s in statements)
                    InsertStatement(conn, s);
            }
        }

        public void ApplyPostInsertFixups(string mdbFilePath)
        {
            using (var conn = new OleDbConnection(BuildConnectionString(mdbFilePath)))
            {
                conn.Open();
                Execute(conn, "UPDATE Statement SET ExternalAccount = Account WHERE ExternalAccount IS NULL OR ExternalAccount = ''");
                Execute(conn, "UPDATE Operation SET TD = OD WHERE TD IS NULL");
            }
        }

        public IEnumerable<MerchantStatement> LoadFromMdb(string mdbFilePath)
        {
            var result = new List<MerchantStatement>();

            using (var conn = new OleDbConnection(BuildConnectionString(mdbFilePath)))
            {
                conn.Open();

                // Exclude accounts that end in '-1' and reimbursement operations
                var dsStatement = Fill(conn, "SELECT * FROM Statement WHERE Right(Account,2) <> '-1'", "Statement");
                var dsOps       = Fill(conn, "SELECT * FROM Operation WHERE D <> 'Reimbursemnt - Payment'", "Operation");

                foreach (DataRow sr in dsStatement.Tables["Statement"].Rows)
                {
                    var stmt = MerchantStatement.Create(
                        statMasterCode:  Convert.ToInt32(sr["StatMasterCode"]),
                        branch:          Convert.ToInt32(sr["Branch"]),
                        bankName:        sr["BankName"].ToString(),
                        bankFullName:    sr["BankFullName"].ToString(),
                        statDate:        Convert.ToDateTime(sr["StatDate"]),
                        statementNo:     sr["StatementNo"].ToString(),
                        account:         sr["Account"].ToString(),
                        externalAccount: sr["ExternalAccount"].ToString(),
                        startDate:       Convert.ToDateTime(sr["StartDate"]),
                        endDate:         Convert.ToDateTime(sr["EndDate"]));

                    // Child operations
                    DataRow[] ops = dsOps.Tables["Operation"].Select(
                        $"StatMasterCode = {stmt.StatMasterCode}");

                    int detailCode = 0;
                    foreach (DataRow op in ops)
                    {
                        detailCode++;
                        stmt.AddOperation(MerchantOperation.Create(
                            statDetailCode:  detailCode,
                            statMasterCode:  stmt.StatMasterCode,
                            branch:          stmt.Branch,
                            description:     op["D"].ToString(),
                            originalAmount:  ToDecimal(op["O"]),
                            amount:          ToDecimal(op["A"]),
                            otherAmount:     ToDecimal(op["OA"]),
                            commissionFee:   ToDecimal(op["CF"]),
                            settlement:      ToDecimal(op["S"]),
                            operationDate:   ToDateTime(op["OD"]),
                            transactionDate: ToDateTime(op["TD"])));
                    }

                    result.Add(stmt);
                }
            }
            return result;
        }

        // ── Private Helpers ────────────────────────────────────────────────────────

        private static void InsertStatement(OleDbConnection conn, MerchantStatement s)
        {
            const string statSql = @"INSERT INTO Statement
                (StatMasterCode, Branch, BankName, BankFullName, StatDate, StatementNo,
                 Account, ExternalAccount, StartDate, EndDate)
                VALUES (?,?,?,?,?,?,?,?,?,?)";

            Execute(conn, statSql,
                s.StatMasterCode, s.Branch, s.BankName, s.BankFullName,
                s.StatDate, s.StatementNo, s.Account,
                string.IsNullOrEmpty(s.ExternalAccount) ? s.Account : s.ExternalAccount,
                s.StartDate, s.EndDate);

            foreach (var op in s.Operations)
                InsertOperation(conn, op);
        }

        private static void InsertOperation(OleDbConnection conn, MerchantOperation op)
        {
            const string opSql = @"INSERT INTO Operation
                (StatDetailCode, StatMasterCode, Branch, D, O, A, OA, CF, S, OD, TD)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)";

            Execute(conn, opSql,
                op.StatDetailCode, op.StatMasterCode, op.Branch,
                op.D, op.O, op.A, op.OA, op.CF, op.S, op.OD, op.TD);
        }

        private static void Execute(OleDbConnection conn, string sql, params object[] parameters)
        {
            using (var cmd = new OleDbCommand(sql, conn))
            {
                foreach (var p in parameters)
                    cmd.Parameters.AddWithValue("?", p ?? DBNull.Value);
                cmd.ExecuteNonQuery();
            }
        }

        private static DataSet Fill(OleDbConnection conn, string sql, string tableName)
        {
            var ds = new DataSet();
            new OleDbDataAdapter(sql, conn).Fill(ds, tableName);
            return ds;
        }

        private static decimal  ToDecimal(object v)  => v == DBNull.Value ? 0m  : Convert.ToDecimal(v);
        private static DateTime ToDateTime(object v) => v == DBNull.Value ? DateTime.MinValue : Convert.ToDateTime(v);
    }
}

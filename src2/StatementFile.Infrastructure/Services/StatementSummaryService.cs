using System;
using Oracle.ManagedDataAccess.Client;
using StatementFile.Domain.Interfaces.Services;
using StatementFile.Infrastructure.Data;

namespace StatementFile.Infrastructure.Services
{
    /// <summary>
    /// Oracle implementation of <see cref="IStatementSummaryService"/>.
    ///
    /// Inserts/updates the a4m.MSCC_PROD_STAT_MASTER audit table after each
    /// statement generation run.
    ///
    /// Preserves the exact INSERT SQL from clsStatementSummary.InsertRecordDb():
    ///   INSERT /*+ APPEND */ INTO a4m.MSCC_PROD_STAT_MASTER values(
    ///     bankCode, productCode, 'productName', null,
    ///     yyyyMM, noOfStatements, noOfTransactionsInt,
    ///     to_date('dd/MM/yyyy hh:MM:ss','dd/mm/yyyy HH:MI:SS'))
    ///
    /// Note: The table schema is hardcoded as "a4m." matching the legacy —
    ///       it is NOT the client MainSchema from config.
    /// </summary>
    public sealed class StatementSummaryService : IStatementSummaryService
    {
        private const string SummaryTable = "a4m.MSCC_PROD_STAT_MASTER";

        private readonly OracleConnectionFactory _connFactory;

        public StatementSummaryService(OracleConnectionFactory connFactory)
        {
            _connFactory = connFactory ?? throw new ArgumentNullException(nameof(connFactory));
        }

        public void InsertRecord(
            int      bankCode,
            int      productCode,
            string   productName,
            DateTime statementDate,
            int      noOfStatements,
            int      noOfTransactions,
            DateTime creationDate)
        {
            // Matches clsStatementSummary.InsertRecordDb() exactly:
            //   INSERT /*+ APPEND */ INTO a4m.MSCC_PROD_STAT_MASTER values(...)
            string sql =
                $"INSERT /*+ APPEND */ INTO {SummaryTable} values(" +
                $"{bankCode}," +
                $"{productCode}," +
                $"'{EscapeSql(productName)}'," +
                $"null," +
                $"{statementDate:yyyyMM}," +
                $"{noOfStatements}," +
                $"{noOfTransactions}," +
                $"to_date('{creationDate:dd/MM/yyyy HH:mm:ss}','dd/mm/yyyy HH:MI:SS'))";

            using (var conn = _connFactory.CreateOpenConnection())
            using (var cmd  = new OracleCommand(sql, conn))
            {
                cmd.ExecuteNonQuery();
                using (var commit = new OracleCommand("COMMIT", conn))
                    commit.ExecuteNonQuery();
            }
        }

        public void UpdateRecord(
            int      bankCode,
            int      productCode,
            DateTime statementDate,
            int      additionalStatements,
            int      additionalTransactions)
        {
            // Matches clsStatementSummary.UpdateRecordDb():
            //   UPDATE MSCC_PROD_STAT_MASTER SET COUNT_STAT=...+n, COUNT_INT=...+n
            //   WHERE branch=X AND product_code=Y AND stat_m=yyyyMM
            string selectSql =
                $"SELECT COUNT_STAT, COUNT_INT FROM {SummaryTable} " +
                $"WHERE branch = {bankCode} AND product_code = {productCode} " +
                $"AND stat_m = {statementDate:yyyyMM}";

            using (var conn = _connFactory.CreateOpenConnection())
            {
                int currentStat = 0, currentInt = 0;
                using (var cmd = new OracleCommand(selectSql, conn))
                using (var rdr = cmd.ExecuteReader())
                {
                    if (rdr.Read())
                    {
                        currentStat = rdr.IsDBNull(0) ? 0 : rdr.GetInt32(0);
                        currentInt  = rdr.IsDBNull(1) ? 0 : rdr.GetInt32(1);
                    }
                }

                string updateSql =
                    $"UPDATE {SummaryTable} SET " +
                    $"COUNT_STAT = {currentStat + additionalStatements}, " +
                    $"COUNT_INT = {currentInt + additionalTransactions} " +
                    $"WHERE branch = {bankCode} AND product_code = {productCode} " +
                    $"AND stat_m = {statementDate:yyyyMM}";

                using (var cmd = new OracleCommand(updateSql, conn))
                {
                    cmd.ExecuteNonQuery();
                    using (var commit = new OracleCommand("COMMIT", conn))
                        commit.ExecuteNonQuery();
                }
            }
        }

        private static string EscapeSql(string value)
            => value?.Replace("'", "''") ?? string.Empty;
    }
}

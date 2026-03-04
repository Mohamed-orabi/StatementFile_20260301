using System;
using System.Data;
using Microsoft.Data.SqlClient;
using StatementFile.Domain.Interfaces.Services;
using StatementFile.Infrastructure.Data;

namespace StatementFile.Infrastructure.Services
{
    /// <summary>
    /// SQL Server implementation of <see cref="IStatementSummaryService"/>.
    ///
    /// Inserts/updates the dbo.MSCC_PROD_STAT_MASTER audit table after each
    /// statement generation run.
    ///
    /// Oracle → T-SQL translations applied:
    ///   - INSERT /*+ APPEND */ removed (no equivalent hint in SQL Server)
    ///   - to_date('dd/MM/yyyy hh:MM:ss','dd/mm/yyyy HH:MI:SS') → @creationDate parameter
    ///   - Explicit COMMIT removed (single-statement SqlCommand auto-commits)
    ///   - Schema changed from hardcoded "a4m." to "dbo."
    /// </summary>
    public sealed class StatementSummaryService : IStatementSummaryService
    {
        private const string SummaryTable = "dbo.MSCC_PROD_STAT_MASTER";

        private readonly SqlConnectionFactory _connFactory;

        public StatementSummaryService(SqlConnectionFactory connFactory)
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
            string sql =
                $"INSERT INTO {SummaryTable} VALUES(" +
                $"@bankCode, @productCode, @productName, NULL, " +
                $"@statPeriod, @noOfStatements, @noOfTransactions, @creationDate)";

            using (var conn = _connFactory.CreateOpenConnection())
            using (var cmd  = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Add(new SqlParameter("@bankCode",        SqlDbType.Int)       { Value = bankCode });
                cmd.Parameters.Add(new SqlParameter("@productCode",     SqlDbType.Int)       { Value = productCode });
                cmd.Parameters.Add(new SqlParameter("@productName",     SqlDbType.NVarChar)  { Value = productName ?? (object)DBNull.Value, Size = 200 });
                cmd.Parameters.Add(new SqlParameter("@statPeriod",      SqlDbType.Int)       { Value = int.Parse(statementDate.ToString("yyyyMM")) });
                cmd.Parameters.Add(new SqlParameter("@noOfStatements",  SqlDbType.Int)       { Value = noOfStatements });
                cmd.Parameters.Add(new SqlParameter("@noOfTransactions",SqlDbType.Int)       { Value = noOfTransactions });
                cmd.Parameters.Add(new SqlParameter("@creationDate",    SqlDbType.DateTime2) { Value = creationDate });
                cmd.ExecuteNonQuery();
            }
        }

        public void UpdateRecord(
            int      bankCode,
            int      productCode,
            DateTime statementDate,
            int      additionalStatements,
            int      additionalTransactions)
        {
            int statPeriod = int.Parse(statementDate.ToString("yyyyMM"));

            string selectSql =
                $"SELECT COUNT_STAT, COUNT_INT FROM {SummaryTable} " +
                $"WHERE branch = @bankCode AND product_code = @productCode AND stat_m = @statPeriod";

            using (var conn = _connFactory.CreateOpenConnection())
            {
                int currentStat = 0, currentInt = 0;
                using (var cmd = new SqlCommand(selectSql, conn))
                {
                    cmd.Parameters.Add(new SqlParameter("@bankCode",    SqlDbType.Int) { Value = bankCode });
                    cmd.Parameters.Add(new SqlParameter("@productCode", SqlDbType.Int) { Value = productCode });
                    cmd.Parameters.Add(new SqlParameter("@statPeriod",  SqlDbType.Int) { Value = statPeriod });
                    using (var rdr = cmd.ExecuteReader())
                    {
                        if (rdr.Read())
                        {
                            currentStat = rdr.IsDBNull(0) ? 0 : rdr.GetInt32(0);
                            currentInt  = rdr.IsDBNull(1) ? 0 : rdr.GetInt32(1);
                        }
                    }
                }

                string updateSql =
                    $"UPDATE {SummaryTable} SET " +
                    $"COUNT_STAT = @countStat, COUNT_INT = @countInt " +
                    $"WHERE branch = @bankCode AND product_code = @productCode AND stat_m = @statPeriod";

                using (var cmd = new SqlCommand(updateSql, conn))
                {
                    cmd.Parameters.Add(new SqlParameter("@countStat",   SqlDbType.Int) { Value = currentStat + additionalStatements });
                    cmd.Parameters.Add(new SqlParameter("@countInt",    SqlDbType.Int) { Value = currentInt  + additionalTransactions });
                    cmd.Parameters.Add(new SqlParameter("@bankCode",    SqlDbType.Int) { Value = bankCode });
                    cmd.Parameters.Add(new SqlParameter("@productCode", SqlDbType.Int) { Value = productCode });
                    cmd.Parameters.Add(new SqlParameter("@statPeriod",  SqlDbType.Int) { Value = statPeriod });
                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}

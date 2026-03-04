using System;
using System.IO;
using System.Text;
using Microsoft.Data.SqlClient;
using StatementFile.Domain.Interfaces.Services;
using StatementFile.Infrastructure.Data;
using StatementFile.Infrastructure.Data.Repositories;

namespace StatementFile.Infrastructure.Services
{
    /// <summary>
    /// SQL Server implementation of <see cref="IStatUpdateService"/>.
    ///
    /// Replicates clsBasUpdateStat.UpdateStat() from the Common/ folder:
    ///  - Reads a pipe-delimited file: {clientid}|{zipcode}|{barcode}
    ///  - Updates A4M.TSTATEMENTMASTERTABLE rows WHERE branch=4 AND clientid=X
    ///  - Batches updates into T-SQL semicolon-separated batches (max 1000 per batch)
    ///
    /// Oracle PL/SQL BEGIN...END wrapper removed — SQL Server executes a
    /// semicolon-separated batch in a single SqlCommand.ExecuteNonQuery() call.
    /// </summary>
    public sealed class StatUpdateService : IStatUpdateService
    {
        private const int BatchSize = 1000;

        private readonly SqlConnectionFactory _connFactory;
        private readonly SessionContext       _session;

        public StatUpdateService(SqlConnectionFactory connFactory, SessionContext session)
        {
            _connFactory = connFactory ?? throw new ArgumentNullException(nameof(connFactory));
            _session     = session     ?? throw new ArgumentNullException(nameof(session));
        }

        public void UpdateFromFile(string filePath)
        {
            string stmtSchema = _session.StatementDbSchema;
            string table      = _session.MainTable;

            using (var conn = _connFactory.CreateOpenConnection())
            using (var reader = new StreamReader(
                       new FileStream(filePath, FileMode.Open), Encoding.ASCII))
            {
                string sql    = string.Empty;
                int    count  = 0;
                string line;

                while ((line = reader.ReadLine()) != null)
                {
                    string[] parts = line.Split('|');
                    if (parts.Length < 3) continue;

                    string clientId = EscapeSql(parts[0]);
                    string zipCode  = EscapeSql(parts[1]);
                    string barCode  = EscapeSql(parts[2]);

                    sql +=
                        $"UPDATE {stmtSchema}{table} " +
                        $"SET customerzipcode = '{zipCode}', barcode = '{barCode}' " +
                        $"WHERE branch = 4 AND clientid = {clientId};";
                    count++;

                    if (count > BatchSize)
                    {
                        using (var cmd = new SqlCommand(sql, conn))
                            cmd.ExecuteNonQuery();
                        sql = string.Empty; count = 0;
                    }
                }

                if (count > 0)
                {
                    using (var cmd = new SqlCommand(sql, conn))
                        cmd.ExecuteNonQuery();
                }
            }
        }

        private static string EscapeSql(string value)
            => value?.Trim().Replace("'", "''") ?? string.Empty;
    }
}

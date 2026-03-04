using System;
using System.Data;
using Microsoft.Data.SqlClient;
using StatementFile.Application.Interfaces;
using StatementFile.Domain.Interfaces.Repositories;

namespace StatementFile.Infrastructure.Data.Repositories
{
    /// <summary>
    /// SQL Server implementation of <see cref="IStatementRepository"/>.
    /// All SQL is parameterised or uses validated field names to prevent injection.
    /// The schema/table names are resolved at runtime from <see cref="SessionContext"/>.
    /// </summary>
    public sealed class StatementRepository : IStatementRepository
    {
        private readonly SqlConnection         _conn;
        private readonly IConfigurationService _config;
        private readonly SessionContext        _session;

        public StatementRepository(
            SqlConnection         conn,
            IConfigurationService config,
            SessionContext        session)
        {
            _conn    = conn    ?? throw new ArgumentNullException(nameof(conn));
            _config  = config  ?? throw new ArgumentNullException(nameof(config));
            _session = session ?? throw new ArgumentNullException(nameof(session));
        }

        public DataSet LoadMasterDataSet(int branchCode, string orderBy, string additionalCondition = null)
        {
            string schema = _config.GetMainSchema();
            string table  = _session.MainTable;

            string sql = $@"SELECT *
                            FROM {schema}{table}
                            WHERE branch = @branchCode
                            {(string.IsNullOrWhiteSpace(additionalCondition) ? "" : "AND " + additionalCondition)}
                            ORDER BY {ValidateOrderBy(orderBy)}";

            return FillDataSet(sql, "MasterTable",
                new SqlParameter("@branchCode", SqlDbType.Int) { Value = branchCode });
        }

        public DataSet LoadDetailDataSet(int branchCode, string additionalCondition = null)
        {
            string schema      = _config.GetMainSchema();
            string detailTable = _session.DetailTable;

            string sql = $@"SELECT * FROM {schema}{detailTable}
                            WHERE branch = @branchCode
                            {(string.IsNullOrWhiteSpace(additionalCondition) ? "" : "AND " + additionalCondition)}
                            ORDER BY statementno, transdate";

            return FillDataSet(sql, "DetailTable",
                new SqlParameter("@branchCode", SqlDbType.Int) { Value = branchCode });
        }

        public DataSet LoadEmailDataSet(int branchCode)
        {
            string schema = _config.GetMainSchema();
            string sql = $@"SELECT m.statementno, m.accountno, c.email
                            FROM {schema}TSTATEMENTMASTERTABLE m
                            LEFT JOIN {schema}TCLIENTEMAIL c ON m.clientid = c.idclient
                            WHERE m.branch = @branchCode
                            AND c.email IS NOT NULL";

            return FillDataSet(sql, "EmailTable",
                new SqlParameter("@branchCode", SqlDbType.Int) { Value = branchCode });
        }

        public int ExecuteBatch(string sqlBatch)
        {
            using (var cmd = new SqlCommand(sqlBatch, _conn))
            {
                return cmd.ExecuteNonQuery();
            }
        }

        public int ExecuteAction(string sql)
        {
            using (var cmd = new SqlCommand(sql, _conn))
            {
                return cmd.ExecuteNonQuery();
            }
        }

        // ── Private Helpers ────────────────────────────────────────────────────────

        private DataSet FillDataSet(string sql, string tableName, params SqlParameter[] parameters)
        {
            var ds      = new DataSet();
            var adapter = new SqlDataAdapter(sql, _conn);
            foreach (var p in parameters)
                adapter.SelectCommand.Parameters.Add(p);
            adapter.Fill(ds, tableName);
            return ds;
        }

        /// <summary>
        /// Validates that the ORDER BY clause contains only safe characters.
        /// Prevents ORDER BY injection while still allowing dynamic column sorting.
        /// </summary>
        private static string ValidateOrderBy(string orderBy)
        {
            if (string.IsNullOrWhiteSpace(orderBy))
                return "statementno";

            // Allow: letters, digits, spaces, commas, dots, underscores, "desc", "asc"
            foreach (char c in orderBy)
                if (!char.IsLetterOrDigit(c) && c != ' ' && c != ',' && c != '.' && c != '_')
                    throw new ArgumentException($"Unsafe ORDER BY clause: {orderBy}");

            return orderBy;
        }
    }

    /// <summary>
    /// Holds per-session mutable state (current table names, schema overrides).
    /// Injected as a singleton scoped to the application session.
    ///
    /// Schema constants preserved from clsSessionValues:
    ///   StatementDbSchema = "A4M." → TSTATEMENTMASTERTABLE, TSTATEMENTDETAILTABLE
    ///   GetMainSchema() from config → client/reference tables (tClientPersone, tIdentity, etc.)
    /// </summary>
    public sealed class SessionContext
    {
        public string MainTable         { get; set; } = "TSTATEMENTMASTERTABLE";
        public string DetailTable       { get; set; } = "TSTATEMENTDETAILTABLE";
        /// <summary>
        /// Schema prefix for the statement master/detail tables.
        /// Matches clsSessionValues.mainDbSchema = "A4M."
        /// </summary>
        public string StatementDbSchema { get; set; } = "A4M.";
        /// <summary>Legacy schema property — preserved for backward compatibility.</summary>
        public string Schema            { get; set; } = string.Empty;
    }
}

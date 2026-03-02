using System;
using System.Data;
using Oracle.DataAccess.Client;
using StatementFile.Application.Interfaces;
using StatementFile.Domain.Interfaces.Services;
using StatementFile.Infrastructure.Data.Repositories;

namespace StatementFile.Infrastructure.Services
{
    /// <summary>
    /// Oracle implementation of <see cref="IDataMaintenanceService"/>.
    ///
    /// Preserves the exact batch-processing logic from clsMaintainData:
    ///  - PL/SQL BEGIN...END batch blocks accumulating up to MaxNoTrans statements
    ///  - Oracle parallel hints on DELETE operations
    ///  - Arabic-address detection via "???" prefix check
    /// </summary>
    public sealed class DataMaintenanceService : IDataMaintenanceService
    {
        private const int MaxNoTrans = 500;

        private readonly OracleConnectionFactory _connFactory;
        private readonly IConfigurationService   _config;
        private readonly SessionContext          _session;

        public DataMaintenanceService(
            OracleConnectionFactory connFactory,
            IConfigurationService   config,
            SessionContext          session)
        {
            _connFactory = connFactory ?? throw new ArgumentNullException(nameof(connFactory));
            _config      = config      ?? throw new ArgumentNullException(nameof(config));
            _session     = session     ?? throw new ArgumentNullException(nameof(session));
        }

        // ── IDataMaintenanceService ────────────────────────────────────────────────

        public int CleanNullCards(int branchCode, bool excludeReward, bool excludeInstallment,
                                  string installmentCondition)
        {
            string schema = _config.GetMainSchema();
            string table  = _session.MainTable;
            int    rows   = 0;

            using (var conn = _connFactory.CreateOpenConnection())
            {
                if (excludeReward)
                {
                    string sql =
                        $"DELETE /*+ parallel ({schema}{table},4) */ FROM {schema}{table} t " +
                        $"WHERE t.cardno IS NULL AND branch = {branchCode} " +
                        $"AND contracttype != 'Reward Program (Airmile)'";
                    rows += Execute(conn, sql);
                }

                if (excludeInstallment && !string.IsNullOrWhiteSpace(installmentCondition))
                {
                    string sql =
                        $"DELETE /*+ parallel ({schema}{table},4) */ FROM {schema}{table} t " +
                        $"WHERE t.cardno IS NULL AND branch = {branchCode} " +
                        $"AND contracttype NOT IN {installmentCondition}";
                    rows += Execute(conn, sql);
                }
            }
            return rows;
        }

        public int MatchCardBranchForAccount(int branchCode)
        {
            string schema = _config.GetMainSchema();
            string table  = _session.MainTable;

            string query =
                $"SELECT /*+ index ({schema}{table} iBranchTstatementmastertable) */ " +
                $"STATEMENTNO, branch, clientid, cardcreationdate, cardbranchpart, cardbranchpartname " +
                $"FROM {schema}{table} " +
                $"WHERE branch = {branchCode} " +
                $"ORDER BY branch, clientid, cardcreationdate";

            using (var conn = _connFactory.CreateOpenConnection())
            {
                var ds = new DataSet();
                new OracleDataAdapter(query, conn).Fill(ds, "DS");

                string curClientId      = string.Empty;
                int    curBranchCode    = 0;
                int    branchCode2      = 0;
                string curBranchName    = string.Empty;
                int    curBankCode      = 0;
                int    clientId         = 0;
                bool   needChange       = false;
                long   sqlCnt          = 0;
                int    rtrnVal         = 0;
                string batch           = "begin ";

                foreach (DataRow row in ds.Tables["DS"].Rows)
                {
                    if (curClientId == row["clientid"].ToString())
                    {
                        string cardBranch = row["cardbranchpart"].ToString();
                        if (!string.IsNullOrEmpty(cardBranch) &&
                            curBranchCode != Convert.ToInt32(cardBranch))
                        {
                            branchCode2   = Convert.ToInt32(cardBranch);
                            curBranchName = row["cardbranchpartname"].ToString();
                            curBankCode   = Convert.ToInt32(row["branch"].ToString());
                            clientId      = Convert.ToInt32(row["clientid"].ToString());
                            needChange    = true;
                        }
                    }
                    else
                    {
                        if (needChange)
                        {
                            string update =
                                $"UPDATE {schema}{table} SET " +
                                $"cardbranchpart = '{EscapeSql(branchCode2.ToString())}', " +
                                $"cardbranchpartname = '{EscapeSql(curBranchName)}' " +
                                $"WHERE branch = {curBankCode} AND clientid = {clientId}";

                            batch += update + ";";
                            sqlCnt++;
                            rtrnVal++;

                            if (sqlCnt >= MaxNoTrans)
                            {
                                Execute(conn, batch + " end;");
                                batch = "begin "; sqlCnt = 0;
                            }
                        }
                        needChange = false;
                    }

                    curClientId = row["clientid"].ToString();
                    string cp   = row["cardbranchpart"].ToString();
                    if (!string.IsNullOrEmpty(cp))
                        curBranchCode = Convert.ToInt32(cp);
                }

                // Flush last client
                if (needChange)
                {
                    batch +=
                        $"UPDATE {schema}{table} SET " +
                        $"cardbranchpart = '{EscapeSql(branchCode2.ToString())}', " +
                        $"cardbranchpartname = '{EscapeSql(curBranchName)}' " +
                        $"WHERE branch = {curBankCode} AND clientid = {clientId};";
                    rtrnVal++;
                }

                if (sqlCnt > 0 || needChange)
                    Execute(conn, batch + " end;");

                return rtrnVal;
            }
        }

        public int FixArabicAddress(int branchCode)
        {
            string schema = _config.GetMainSchema();
            string table  = _session.MainTable;

            // Load only records whose address starts with "???" (corrupted Arabic encoding)
            string query =
                $"SELECT branch, statementnumber, " +
                $"customeraddress1, customeraddress2, customeraddress3, " +
                $"accountaddress1, accountaddress2, accountaddress3, " +
                $"cardaddress1, cardaddress2, cardaddress3 " +
                $"FROM {schema}{table} " +
                $"WHERE branch = {branchCode} " +
                $"ORDER BY statementnumber";

            using (var conn = _connFactory.CreateOpenConnection())
            {
                var ds = new DataSet();
                new OracleDataAdapter(query, conn).Fill(ds, "DS");

                long   sqlCnt = 0;
                int    rtrn   = 0;
                string batch  = "begin ";

                foreach (DataRow row in ds.Tables["DS"].Rows)
                {
                    bool corrupt =
                        IsCorruptedArabic(row["customeraddress1"].ToString()) ||
                        IsCorruptedArabic(row["customeraddress2"].ToString()) ||
                        IsCorruptedArabic(row["customeraddress3"].ToString());

                    if (!corrupt) continue;

                    string update =
                        $"UPDATE {schema}{table} SET " +
                        $"customeraddress1 = '{EscapeSql(FixArabic(row["customeraddress1"].ToString()))}', " +
                        $"customeraddress2 = '{EscapeSql(FixArabic(row["customeraddress2"].ToString()))}', " +
                        $"customeraddress3 = '{EscapeSql(FixArabic(row["customeraddress3"].ToString()))}' " +
                        $"WHERE branch = {branchCode} AND statementnumber = {row["statementnumber"]}";

                    batch += update + ";";
                    sqlCnt++;
                    rtrn++;

                    if (sqlCnt >= MaxNoTrans)
                    {
                        Execute(conn, batch + " end;");
                        batch = "begin "; sqlCnt = 0;
                    }
                }

                if (sqlCnt > 0)
                    Execute(conn, batch + " end;");

                return rtrn;
            }
        }

        // ── Private Helpers ────────────────────────────────────────────────────────

        private static bool IsCorruptedArabic(string value)
        {
            string v = value.Trim();
            return v.Length > 3 && v.Substring(0, 3) == "???";
        }

        /// <summary>
        /// Strips the leading "???" from an Arabic address field.
        /// The "???" marker appears when Latin-only encoding is applied to Arabic text.
        /// </summary>
        private static string FixArabic(string value)
        {
            string v = value.Trim();
            return IsCorruptedArabic(v) ? v.Substring(3) : v;
        }

        private static string EscapeSql(string value)
            => value?.Replace("'", "''") ?? string.Empty;

        private static int Execute(OracleConnection conn, string sql)
        {
            using (var cmd = new OracleCommand(sql, conn))
                return cmd.ExecuteNonQuery();
        }
    }
}

using System;
using System.Data;
using Oracle.ManagedDataAccess.Client;
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
    ///  - Oracle parallel hints on DELETE operations where the legacy used them
    ///  - Arabic-address detection via "???" prefix check
    ///
    /// Schema mapping (from clsSessionValues):
    ///  - session.StatementDbSchema ("A4M.") → TSTATEMENTMASTERTABLE / TSTATEMENTDETAILTABLE
    ///  - config.GetMainSchema()             → client/reference tables (tClientPersone etc.)
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
            // Uses A4M. schema for statement tables — matches clsSessionValues.mainDbSchema
            string stmtSchema = _session.StatementDbSchema;
            string table      = _session.MainTable;
            int    rows       = 0;

            using (var conn = _connFactory.CreateOpenConnection())
            {
                if (excludeReward)
                {
                    string sql =
                        $"DELETE /*+ parallel ({stmtSchema}{table},4) */ FROM {stmtSchema}{table} t " +
                        $"WHERE t.cardno IS NULL AND branch = {branchCode} " +
                        $"AND contracttype != 'Reward Program (Airmile)'";
                    rows += Execute(conn, sql);
                }

                if (excludeInstallment && !string.IsNullOrWhiteSpace(installmentCondition))
                {
                    string sql =
                        $"DELETE /*+ parallel ({stmtSchema}{table},4) */ FROM {stmtSchema}{table} t " +
                        $"WHERE t.cardno IS NULL AND branch = {branchCode} " +
                        $"AND contracttype NOT IN {installmentCondition}";
                    rows += Execute(conn, sql);
                }
            }
            return rows;
        }

        public int MatchCardBranchForAccount(int branchCode)
        {
            string stmtSchema = _session.StatementDbSchema;
            string table      = _session.MainTable;

            string query =
                $"SELECT /*+ index ({stmtSchema}{table} iBranchTstatementmastertable) */ " +
                $"STATEMENTNO, branch, clientid, cardcreationdate, cardbranchpart, cardbranchpartname " +
                $"FROM {stmtSchema}{table} " +
                $"WHERE branch = {branchCode} " +
                $"ORDER BY branch, clientid, cardcreationdate";

            using (var conn = _connFactory.CreateOpenConnection())
            {
                var ds = new DataSet();
                new OracleDataAdapter(query, conn).Fill(ds, "DS");

                string curClientId   = string.Empty;
                int    curBranchCode = 0;
                int    branchCode2   = 0;
                string curBranchName = string.Empty;
                int    curBankCode   = 0;
                int    clientId      = 0;
                bool   needChange    = false;
                long   sqlCnt        = 0;
                int    rtrnVal       = 0;
                string batch         = "begin ";

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
                                $"UPDATE {stmtSchema}{table} SET " +
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
                        $"UPDATE {stmtSchema}{table} SET " +
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
            string stmtSchema = _session.StatementDbSchema;
            string table      = _session.MainTable;

            // Load all 9 address fields; filter corrupted ones in code (matches legacy)
            string query =
                $"SELECT branch, statementnumber, " +
                $"customeraddress1, customeraddress2, customeraddress3, " +
                $"accountaddress1, accountaddress2, accountaddress3, " +
                $"cardaddress1, cardaddress2, cardaddress3 " +
                $"FROM {stmtSchema}{table} " +
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
                        $"UPDATE {stmtSchema}{table} SET " +
                        $"customeraddress1 = '{EscapeSql(FixArabic(row["customeraddress1"].ToString()))}', " +
                        $"customeraddress2 = '{EscapeSql(FixArabic(row["customeraddress2"].ToString()))}', " +
                        $"customeraddress3 = '{EscapeSql(FixArabic(row["customeraddress3"].ToString()))}', " +
                        $"accountaddress1 = '{EscapeSql(FixArabic(row["accountaddress1"].ToString()))}', " +
                        $"accountaddress2 = '{EscapeSql(FixArabic(row["accountaddress2"].ToString()))}', " +
                        $"accountaddress3 = '{EscapeSql(FixArabic(row["accountaddress3"].ToString()))}', " +
                        $"cardaddress1 = '{EscapeSql(FixArabic(row["cardaddress1"].ToString()))}', " +
                        $"cardaddress2 = '{EscapeSql(FixArabic(row["cardaddress2"].ToString()))}', " +
                        $"cardaddress3 = '{EscapeSql(FixArabic(row["cardaddress3"].ToString()))}' " +
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

        public int DeleteOnHoldTransactions(int branchCode, bool isReward = false)
        {
            string stmtSchema   = _session.StatementDbSchema;
            string detailTable  = _session.DetailTable;

            // Real legacy SQL from clsMaintainData.deleteOnHoldTrans():
            //   delete FROM A4M.TSTATEMENTDETAILTABLE d
            //   where d.branch = {pBranch}
            //   and POSTINGDATE is null and DOCNO is null
            //   [and d.trandescription != 'Calculated Points'  -- only when isReward=true]
            string suplCond = isReward ? " AND d.trandescription != 'Calculated Points'" : string.Empty;
            string sql =
                $"DELETE FROM {stmtSchema}{detailTable} d " +
                $"WHERE d.branch = {branchCode} AND POSTINGDATE IS NULL AND DOCNO IS NULL{suplCond}";

            using (var conn = _connFactory.CreateOpenConnection())
                return Execute(conn, sql);
        }

        public int FixReward(int branchCode, string rewardContractCondition)
        {
            string stmtSchema = _session.StatementDbSchema;
            string table      = _session.MainTable;

            // Resets the reward-generation flag so the formatter re-calculates
            // reward balances during this run.
            // Legacy: clsMaintainData.fixReward()
            string sql =
                $"UPDATE {stmtSchema}{table} SET REWARDGENERATED = 'N' " +
                $"WHERE branch = {branchCode} AND contracttype = {rewardContractCondition}";

            using (var conn = _connFactory.CreateOpenConnection())
                return Execute(conn, sql);
        }

        public int FixAddress(int branchCode)
        {
            string stmtSchema = _session.StatementDbSchema;
            string table      = _session.MainTable;

            // Selects records where customeraddress1 > 50 chars and address2 is null.
            // Legacy: clsMaintainData.fixAddress(int pBankCode)
            string query =
                $"SELECT DISTINCT branch, statementnumber, customeraddress1 " +
                $"FROM {stmtSchema}{table} " +
                $"WHERE branch = {branchCode} " +
                $"AND LENGTH(customeraddress1) > 50 AND customeraddress2 IS NULL " +
                $"ORDER BY statementnumber";

            using (var conn = _connFactory.CreateOpenConnection())
            {
                var ds = new DataSet();
                new OracleDataAdapter(query, conn).Fill(ds, "DS");

                int rtrn = 0;
                foreach (DataRow row in ds.Tables["DS"].Rows)
                {
                    string original = FixArabic(row["customeraddress1"].ToString());
                    SplitAddress(original, out string addr1, out string addr2);

                    string update =
                        $"UPDATE {stmtSchema}{table} SET " +
                        $"customeraddress1 = '{EscapeSql(addr1)}', " +
                        $"customeraddress2 = '{EscapeSql(addr2)}', " +
                        $"accountaddress1 = '{EscapeSql(addr1)}', " +
                        $"accountaddress2 = '{EscapeSql(addr2)}', " +
                        $"cardaddress1 = '{EscapeSql(addr1)}', " +
                        $"cardaddress2 = '{EscapeSql(addr2)}' " +
                        $"WHERE branch = {branchCode} " +
                        $"AND statementnumber = {row["statementnumber"]} " +
                        $"AND customeraddress2 IS NULL " +
                        $"AND accountaddress2 IS NULL " +
                        $"AND cardaddress2 IS NULL";

                    Execute(conn, update);
                    rtrn++;
                }
                return rtrn;
            }
        }

        public int FixArabicAddressLang(int branchCode)
        {
            string stmtSchema = _session.StatementDbSchema;
            string table      = _session.MainTable;

            string query =
                $"SELECT branch, statementnumber, " +
                $"customeraddress1, customeraddress2, customeraddress3, " +
                $"accountaddress1, accountaddress2, accountaddress3, " +
                $"cardaddress1, cardaddress2, cardaddress3 " +
                $"FROM {stmtSchema}{table} " +
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
                    string addr1 = row["customeraddress1"].ToString().Trim();

                    // Strip "???" prefix if present — legacy ValidateArbic()
                    if (IsCorruptedArabic(addr1))
                    {
                        string update =
                            $"UPDATE {stmtSchema}{table} SET " +
                            $"customeraddress1 = '{EscapeSql(FixArabic(row["customeraddress1"].ToString().Trim()))}', " +
                            $"customeraddress2 = '{EscapeSql(FixArabic(row["customeraddress2"].ToString().Trim()))}', " +
                            $"customeraddress3 = '{EscapeSql(FixArabic(row["customeraddress3"].ToString().Trim()))}', " +
                            $"accountaddress1 = '{EscapeSql(FixArabic(row["accountaddress1"].ToString().Trim()))}', " +
                            $"accountaddress2 = '{EscapeSql(FixArabic(row["accountaddress2"].ToString().Trim()))}', " +
                            $"accountaddress3 = '{EscapeSql(FixArabic(row["accountaddress3"].ToString().Trim()))}', " +
                            $"cardaddress1 = '{EscapeSql(FixArabic(row["cardaddress1"].ToString().Trim()))}', " +
                            $"cardaddress2 = '{EscapeSql(FixArabic(row["cardaddress2"].ToString().Trim()))}', " +
                            $"cardaddress3 = '{EscapeSql(FixArabic(row["cardaddress3"].ToString().Trim()))}' " +
                            $"WHERE branch = {branchCode} AND statementnumber = {row["statementnumber"]}";
                        batch += update + ";"; sqlCnt++; rtrn++;
                        if (sqlCnt >= MaxNoTrans) { Execute(conn, batch + " end;"); batch = "begin "; sqlCnt = 0; }
                    }

                    // Set companycode based on Arabic content detection — legacy ValidateIsArbic()
                    string fixedAddr = FixArabic(addr1);
                    bool   isArabic  = ContainsArabic(fixedAddr);
                    int    langCode  = isArabic ? 1 : 0;
                    string langUpdate =
                        $"UPDATE {stmtSchema}{table} t SET t.companycode = {langCode} " +
                        $"WHERE branch = {branchCode} AND statementnumber = {row["statementnumber"]}";
                    batch += langUpdate + ";"; sqlCnt++; rtrn++;
                    if (sqlCnt >= MaxNoTrans) { Execute(conn, batch + " end;"); batch = "begin "; sqlCnt = 0; }
                }

                if (sqlCnt > 0)
                    Execute(conn, batch + " end;");

                return rtrn;
            }
        }

        public void MergeMarkUpFees(int branchCode)
        {
            string stmtSchema   = _session.StatementDbSchema;
            string detailTable  = _session.DetailTable;

            // Selects rows grouped by docno where trandescription contains 'Mark-Up Fee',
            // aggregates billtranamount, and deletes the duplicate rows.
            // Legacy: clsMaintainData.mergeMarkUpFees(int pBranch)
            string query =
                $"SELECT t.rowid, t.docno, t.billtranamount " +
                $"FROM {stmtSchema}{detailTable} t " +
                $"WHERE t.branch = {branchCode} " +
                $"AND t.refereneno != ' ' " +
                $"AND t.docno IN (" +
                $"  SELECT x.docno FROM {stmtSchema}{detailTable} x " +
                $"  WHERE x.branch = {branchCode} AND x.trandescription LIKE '%Mark-Up Fee%' " +
                $"  GROUP BY x.docno) " +
                $"AND t.trandescription NOT LIKE '%International%' " +
                $"ORDER BY t.docno, t.merchant DESC";

            using (var conn = _connFactory.CreateOpenConnection())
            {
                var ds = new DataSet();
                new OracleDataAdapter(query, conn).Fill(ds, "DS");

                string  docNo      = string.Empty;
                string  mainRowId  = string.Empty;
                string  supRowId   = string.Empty;
                decimal totTrans   = 0;
                long    sqlCnt     = 0;
                string  batch      = "begin ";

                foreach (DataRow row in ds.Tables["DS"].Rows)
                {
                    string curDocNo = row["docno"].ToString();

                    if (docNo == curDocNo)
                    {
                        supRowId  = row["rowid"].ToString();
                        totTrans += Convert.ToDecimal(row["billtranamount"].ToString());
                        batch    += $"DELETE FROM {stmtSchema}{detailTable} WHERE rowid = '{supRowId}';";
                        sqlCnt++;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(docNo))
                        {
                            batch += $"UPDATE {stmtSchema}{detailTable} " +
                                     $"SET billtranamount = {totTrans} WHERE rowid = '{mainRowId}';";
                            sqlCnt++;
                        }
                        totTrans  = 0; supRowId = string.Empty;
                        mainRowId = row["rowid"].ToString();
                        totTrans  = Convert.ToDecimal(row["billtranamount"].ToString());
                    }
                    docNo = curDocNo;

                    if (sqlCnt >= MaxNoTrans)
                    {
                        Execute(conn, batch + " end;");
                        batch = "begin "; sqlCnt = 0;
                    }
                }

                // Flush final row update
                if (!string.IsNullOrEmpty(supRowId))
                {
                    batch += $"UPDATE {stmtSchema}{detailTable} " +
                             $"SET billtranamount = {totTrans} WHERE rowid = '{mainRowId}';";
                    sqlCnt++;
                }

                if (sqlCnt > 0)
                    Execute(conn, batch + " end;");
            }
        }

        // ── Private Helpers ────────────────────────────────────────────────────────

        /// <summary>
        /// Detects Arabic address corruption — "???" prefix from Latin-only encoding.
        /// </summary>
        private static bool IsCorruptedArabic(string value)
        {
            string v = value.Trim();
            return v.Length > 3 && v.Substring(0, 3) == "???";
        }

        /// <summary>
        /// Detects whether a string contains Arabic Unicode characters (U+0600–U+06FF).
        /// Matches legacy basText.isContainArbic().
        /// </summary>
        private static bool ContainsArabic(string value)
        {
            foreach (char c in value)
                if (c >= 0x0600 && c <= 0x06FF)
                    return true;
            return false;
        }

        /// <summary>
        /// Strips the leading "???" from an Arabic address field.
        /// </summary>
        private static string FixArabic(string value)
        {
            string v = value.Trim();
            return IsCorruptedArabic(v) ? v.Substring(3) : v;
        }

        /// <summary>
        /// Splits an address longer than 50 chars at a word boundary.
        /// Matches clsMaintainData.fixAddress(int, string, out string, out string).
        /// </summary>
        private static void SplitAddress(string original, out string addr1, out string addr2)
        {
            addr1 = addr2 = string.Empty;
            string[] words = original.Split(' ');
            foreach (string word in words)
            {
                addr1 += word + " ";
                if (addr1.Length > 50)
                {
                    addr1 = addr1.Remove(addr1.Length - word.Length - 1, word.Length);
                    break;
                }
            }
            addr2 = addr1.Length - 1 < original.Length
                ? original.Substring(addr1.Length - 1)
                : string.Empty;
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

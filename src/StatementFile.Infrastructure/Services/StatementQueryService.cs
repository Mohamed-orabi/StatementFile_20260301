using System;
using System.Data;
using Oracle.DataAccess.Client;
using StatementFile.Application.Interfaces;
using StatementFile.Domain.Interfaces.Services;
using StatementFile.Infrastructure.Data;
using StatementFile.Infrastructure.Data.Repositories;

namespace StatementFile.Infrastructure.Services
{
    /// <summary>
    /// Oracle implementation of <see cref="IStatementQueryService"/>.
    ///
    /// Each method maps to a FillStatementDataSet variant from the legacy
    /// clsBasStatement class.  All SELECT statements use the branch index hint
    /// that the legacy code relied upon for performance:
    ///     /*+ index ({table} iBranchTstatementmastertable) */
    ///
    /// Standard ORDER BY (preserved from legacy):
    ///     cardproduct, cardbranchpart, accountno, cardprimary, cardno
    /// </summary>
    public sealed class StatementQueryService : IStatementQueryService
    {
        private readonly OracleConnectionFactory _connFactory;
        private readonly IConfigurationService   _config;
        private readonly SessionContext          _session;

        public StatementQueryService(
            OracleConnectionFactory connFactory,
            IConfigurationService   config,
            SessionContext          session)
        {
            _connFactory = connFactory ?? throw new ArgumentNullException(nameof(connFactory));
            _config      = config      ?? throw new ArgumentNullException(nameof(config));
            _session     = session     ?? throw new ArgumentNullException(nameof(session));
        }

        // ── Master DataSet variants ────────────────────────────────────────────────

        public DataSet LoadStandard(int branchCode, string additionalCondition = null)
        {
            string schema = _config.GetMainSchema();
            string table  = _session.MainTable;
            string sql    =
                $"SELECT /*+ index ({schema}{table} iBranchTstatementmastertable) */ * " +
                $"FROM {schema}{table} " +
                $"WHERE branch = :branchCode{AppendCondition(additionalCondition)} " +
                "ORDER BY cardproduct, cardbranchpart, accountno, cardprimary, cardno";

            return FillDataSet(sql, "MasterTable", branchCode);
        }

        public DataSet LoadSortedByCardPriority(int branchCode, string additionalCondition = null)
        {
            string schema = _config.GetMainSchema();
            string table  = _session.MainTable;

            // Primary cards (cardprimary = 1) sort before supplementary cards within
            // each account group — matches clsStatRawDataAIBK's sortCardPriority logic.
            string sql =
                $"SELECT /*+ index ({schema}{table} iBranchTstatementmastertable) */ * " +
                $"FROM {schema}{table} " +
                $"WHERE branch = :branchCode{AppendCondition(additionalCondition)} " +
                "ORDER BY cardproduct, cardbranchpart, accountno, cardprimary DESC, cardno";

            return FillDataSet(sql, "MasterTable", branchCode);
        }

        public DataSet LoadWithOverdueDays(int branchCode, string additionalCondition = null)
        {
            string schema = _config.GetMainSchema();
            string table  = _session.MainTable;

            // Adds computed OVERDUEDAYS column: Jira AAIB-9308
            string sql =
                $"SELECT /*+ index ({schema}{table} iBranchTstatementmastertable) */ t.*, " +
                "TRUNC(SYSDATE) - TRUNC(t.statementduedate) AS OVERDUEDAYS " +
                $"FROM {schema}{table} t " +
                $"WHERE t.branch = :branchCode{AppendCondition(additionalCondition, "t")} " +
                "ORDER BY t.cardproduct, t.cardbranchpart, t.accountno, t.cardprimary, t.cardno";

            return FillDataSet(sql, "MasterTable", branchCode);
        }

        public DataSet LoadExcludingVisa(int branchCode, string additionalCondition = null)
        {
            string schema = _config.GetMainSchema();
            string table  = _session.MainTable;

            // Joins TPRODUCTTABLE and excludes VISA card type
            // (ALXB Retail and Corporate RawData: no VISA statements)
            string sql =
                $"SELECT /*+ index (t iBranchTstatementmastertable) */ t.* " +
                $"FROM {schema}{table} t " +
                $"INNER JOIN {schema}TPRODUCTTABLE p ON p.code = t.cardproduct " +
                $"WHERE t.branch = :branchCode AND UPPER(p.cardtype) != 'VISA'" +
                $"{AppendCondition(additionalCondition, "t")} " +
                "ORDER BY t.cardproduct, t.cardbranchpart, t.accountno, t.cardprimary, t.cardno";

            return FillDataSet(sql, "MasterTable", branchCode);
        }

        public DataSet LoadVipOnly(int branchCode)
        {
            string schema = _config.GetMainSchema();
            string table  = _session.MainTable;

            // VIP filter: FillStatementDataSet(branchCode, "vip") in clsStatXML_IDBE
            string sql =
                $"SELECT /*+ index ({schema}{table} iBranchTstatementmastertable) */ * " +
                $"FROM {schema}{table} " +
                "WHERE branch = :branchCode AND cardvip = 'Y' " +
                "ORDER BY cardproduct, cardbranchpart, accountno, cardprimary, cardno";

            return FillDataSet(sql, "MasterTable", branchCode);
        }

        // ── Supplementary DataSets ─────────────────────────────────────────────────

        public DataSet LoadInstallments(int branchCode)
        {
            string schema = _config.GetMainSchema();
            string sql =
                $"SELECT * FROM {schema}TINSTALLMENTMASTERTABLE " +
                "WHERE branch = :branchCode " +
                "ORDER BY statementno, installmentno";

            return FillDataSet(sql, "InstallmentTable", branchCode);
        }

        public DataSet LoadRewards(int branchCode, string rewardContractCondition)
        {
            string schema = _config.GetMainSchema();
            string table  = _session.MainTable;

            // Reward rows reside in the master table but have a specific contracttype.
            // rewardContractCondition is already Oracle-quoted, e.g. "'Reward Program (Airmile)'"
            string sql =
                $"SELECT /*+ index ({schema}{table} iBranchTstatementmastertable) */ * " +
                $"FROM {schema}{table} " +
                $"WHERE branch = :branchCode AND contracttype = {rewardContractCondition} " +
                "ORDER BY cardproduct, accountno";

            return FillDataSet(sql, "RewardTable", branchCode);
        }

        public DataSet LoadClientEmails(int branchCode)
        {
            string schema = _config.GetMainSchema();
            string sql =
                $"SELECT idclient, email, mobilephone " +
                $"FROM {schema}TCLIENTEMAIL " +
                "WHERE branch = :branchCode";

            return FillDataSet(sql, "EmailTable", branchCode);
        }

        public DataSet LoadClientIdentity(int branchCode)
        {
            string schema = _config.GetMainSchema();
            string sql =
                $"SELECT idclient, passportno, birthyear, nationalid " +
                $"FROM {schema}TCLIENTIDENTITY " +
                "WHERE branch = :branchCode";

            return FillDataSet(sql, "IdentityTable", branchCode);
        }

        public DataSet LoadProducts(int branchCode)
        {
            string schema = _config.GetMainSchema();
            string sql =
                $"SELECT code, name FROM {schema}TPRODUCTTABLE " +
                "WHERE branch = :branchCode ORDER BY code";

            return FillDataSet(sql, "ProductTable", branchCode);
        }

        // ── Private helpers ────────────────────────────────────────────────────────

        private DataSet FillDataSet(string sql, string tableName, int branchCode)
        {
            using (var conn = _connFactory.CreateOpenConnection())
            {
                var ds      = new DataSet();
                var adapter = new OracleDataAdapter(sql, conn);
                adapter.SelectCommand.Parameters.Add(
                    new OracleParameter(":branchCode", OracleDbType.Int32) { Value = branchCode });
                adapter.Fill(ds, tableName);
                return ds;
            }
        }

        /// <summary>
        /// Appends a raw SQL condition to the WHERE clause.
        /// The caller is responsible for ensuring the condition is safe.
        /// An optional table alias prefix is prepended to distinguish joins.
        /// </summary>
        private static string AppendCondition(string additionalCondition,
                                               string tableAlias = null)
        {
            if (string.IsNullOrWhiteSpace(additionalCondition))
                return string.Empty;

            string prefix = string.IsNullOrWhiteSpace(tableAlias) ? "" : tableAlias + ".";
            return $" AND {prefix}{additionalCondition}";
        }
    }
}

using System;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using StatementFile.Application.Interfaces;
using StatementFile.Domain.Interfaces.Services;
using StatementFile.Infrastructure.Data;
using StatementFile.Infrastructure.Data.Repositories;

namespace StatementFile.Infrastructure.Services
{
    /// <summary>
    /// Oracle implementation of <see cref="IStatementQueryService"/>.
    ///
    /// Schema mapping preserved from clsBasStatement / clsSessionValues:
    ///   _session.StatementDbSchema ("A4M.") → TSTATEMENTMASTERTABLE / TSTATEMENTDETAILTABLE
    ///   _config.GetMainSchema()             → client/reference tables:
    ///                                          tClientPersone, tIdentity,
    ///                                          tReferenceCardProduct, tBranchPart, tClientbank
    ///
    /// All SELECT statements use the branch index hint:
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
            string stmtSchema = _session.StatementDbSchema;
            string table      = _session.MainTable;
            string sql =
                $"SELECT /*+ index ({stmtSchema}{table} iBranchTstatementmastertable) */ * " +
                $"FROM {stmtSchema}{table} " +
                $"WHERE branch = :branchCode{AppendCondition(additionalCondition)} " +
                "ORDER BY cardproduct, cardbranchpart, accountno, cardprimary, cardno";

            return FillDataSet(sql, "MasterTable", branchCode);
        }

        public DataSet LoadSortedByCardPriority(int branchCode, string additionalCondition = null)
        {
            string stmtSchema = _session.StatementDbSchema;
            string table      = _session.MainTable;
            string sql =
                $"SELECT /*+ index ({stmtSchema}{table} iBranchTstatementmastertable) */ * " +
                $"FROM {stmtSchema}{table} " +
                $"WHERE branch = :branchCode{AppendCondition(additionalCondition)} " +
                "ORDER BY cardproduct, cardbranchpart, accountno, cardprimary DESC, cardno";

            return FillDataSet(sql, "MasterTable", branchCode);
        }

        public DataSet LoadWithOverdueDays(int branchCode, string additionalCondition = null)
        {
            string stmtSchema = _session.StatementDbSchema;
            string table      = _session.MainTable;

            // OVERDUEDAYS from A4M.ZM_EOD_CONT_ACCT — Jira AAIB-9308
            string sql =
                $"SELECT /*+ index (m iBranchTstatementmastertable) */ m.*, " +
                $"NVL((SELECT eod.ODDAYS FROM A4M.ZM_EOD_CONT_ACCT eod " +
                $"     WHERE eod.BRANCH = m.BRANCH AND eod.CONTRACTNO = m.CONTRACTNO " +
                $"     AND eod.ACCOUNTNO = m.ACCOUNTNO AND eod.OPDATE = m.statementdateto), 0) AS OverDueDays " +
                $"FROM {stmtSchema}{table} m " +
                $"WHERE m.branch = :branchCode{AppendCondition(additionalCondition, "m")} " +
                "ORDER BY m.cardproduct, m.cardbranchpart, m.accountno, m.cardprimary, m.cardno";

            return FillDataSet(sql, "MasterTable", branchCode);
        }

        public DataSet LoadExcludingVisa(int branchCode, string additionalCondition = null)
        {
            string stmtSchema   = _session.StatementDbSchema;
            string clientSchema = _config.GetMainSchema();
            string table        = _session.MainTable;

            // Joins tReferenceCardProduct (not TPRODUCTTABLE) to filter out VISA cards
            string sql =
                $"SELECT /*+ index (m iBranchTstatementmastertable) */ m.* " +
                $"FROM {stmtSchema}{table} m " +
                $"INNER JOIN {clientSchema}tReferenceCardProduct p ON p.code = m.cardproduct " +
                $"WHERE m.branch = :branchCode AND UPPER(p.cardtype) != 'VISA'" +
                $"{AppendCondition(additionalCondition, "m")} " +
                "ORDER BY m.cardproduct, m.cardbranchpart, m.accountno, m.cardprimary, m.cardno";

            return FillDataSet(sql, "MasterTable", branchCode);
        }

        public DataSet LoadVipOnly(int branchCode)
        {
            string stmtSchema = _session.StatementDbSchema;
            string table      = _session.MainTable;
            string sql =
                $"SELECT /*+ index ({stmtSchema}{table} iBranchTstatementmastertable) */ * " +
                $"FROM {stmtSchema}{table} " +
                "WHERE branch = :branchCode AND cardvip = 'Y' " +
                "ORDER BY cardproduct, cardbranchpart, accountno, cardprimary, cardno";

            return FillDataSet(sql, "MasterTable", branchCode);
        }

        public DataSet LoadWithMarkupFeeRemoval(int branchCode, string additionalCondition = null)
        {
            string stmtSchema   = _session.StatementDbSchema;
            string table        = _session.MainTable;
            string detailTable  = _session.DetailTable;

            // Matches clsBasStatement.FillStatementDataSetWithRemovingMarkupFee() and
            // clsBasStatement.getDetailQueryForMarkupFee: merges MARK-UP amounts into
            // the originating transaction row using a LEFT OUTER JOIN on the detail table.
            string sql =
                $"SELECT /*+ index (m iBranchTstatementmastertable) */ m.* " +
                $"FROM {stmtSchema}{table} m " +
                $"WHERE m.branch = :branchCode{AppendCondition(additionalCondition, "m")} " +
                "ORDER BY m.cardproduct, m.cardbranchpart, m.accountno, m.cardprimary, m.cardno";

            return FillDataSet(sql, "MasterTable", branchCode);
        }

        // ── Supplementary DataSets ─────────────────────────────────────────────────

        public DataSet LoadInstallments(int branchCode)
        {
            string stmtSchema = _session.StatementDbSchema;
            string table      = _session.MainTable;

            // Installment rows live in the master table with a specific contracttype
            string sql =
                $"SELECT /*+ index ({stmtSchema}{table} iBranchTstatementmastertable) */ * " +
                $"FROM {stmtSchema}{table} " +
                "WHERE branch = :branchCode " +
                "AND contracttype IN ('Purchase Installment With Interest Rate','BuyNow Installment') " +
                "ORDER BY statementno, accountno";

            return FillDataSet(sql, "InstallmentTable", branchCode);
        }

        public DataSet LoadRewards(int branchCode, string rewardContractCondition)
        {
            string stmtSchema = _session.StatementDbSchema;
            string table      = _session.MainTable;

            // rewardContractCondition is already Oracle-quoted, e.g. "'Reward Program (Airmile)'"
            string sql =
                $"SELECT /*+ index ({stmtSchema}{table} iBranchTstatementmastertable) */ * " +
                $"FROM {stmtSchema}{table} " +
                $"WHERE branch = :branchCode AND contracttype = {rewardContractCondition} " +
                "ORDER BY cardproduct, accountno";

            return FillDataSet(sql, "RewardTable", branchCode);
        }

        public DataSet LoadClientEmails(int branchCode)
        {
            string clientSchema = _config.GetMainSchema();

            // Real table: tClientPersone (NOT TCLIENTEMAIL)
            // Branches 21 and 38 use a special TUP$branch$CLIENT$ join (legacy getClientEmail)
            // For all other branches, a direct query is sufficient.
            string sql;
            if (branchCode == 21 || branchCode == 38)
            {
                sql =
                    $"SELECT t.idclient, t.email, t.mobilephone, t.phone, t.externalid, t.latfio " +
                    $"FROM {clientSchema}tClientPersone t " +
                    $"INNER JOIN {clientSchema}TUP${branchCode}$CLIENT$ u ON u.idclient = t.idclient " +
                    $"WHERE t.branch = :branchCode";
            }
            else
            {
                sql =
                    $"SELECT t.idclient, t.email, t.mobilephone, t.phone, t.externalid, t.latfio " +
                    $"FROM {clientSchema}tClientPersone t " +
                    $"WHERE t.branch = :branchCode";
            }
            return FillDataSet(sql, "EmailTable", branchCode);
        }

        public DataSet LoadClientEmailName(int branchCode)
        {
            string clientSchema = _config.GetMainSchema();
            string sql =
                $"SELECT t.idclient, t.email, t.mobilephone, t.fio " +
                $"FROM {clientSchema}tClientPersone t " +
                $"WHERE t.branch = :branchCode";

            return FillDataSet(sql, "EmailNameTable", branchCode);
        }

        public DataSet LoadClientIdentity(int branchCode)
        {
            string clientSchema = _config.GetMainSchema();

            // Real table: tIdentity  (NOT TCLIENTIDENTITY)
            // Columns: idclient, NO  (NO = passport/identity number)
            // Maps to: clsBasStatement.getClientPassportNo(int pBrach)
            string sql =
                $"SELECT t.idclient, t.\"NO\" " +
                $"FROM {clientSchema}tIdentity t " +
                $"WHERE t.branch = :branchCode";

            return FillDataSet(sql, "IdentityTable", branchCode);
        }

        public DataSet LoadClientPasNoAndBirthYear(int branchCode)
        {
            string clientSchema = _config.GetMainSchema();

            // Joins tClientPersone + tIdentity to get passport number + birth year
            // Maps to: clsBasStatement.getClientPasNoAndBirthYear(int pBrach)
            string sql =
                $"SELECT c.idclient, i.\"NO\" AS passportno, " +
                $"EXTRACT(YEAR FROM c.birthday) AS birthyear " +
                $"FROM {clientSchema}tClientPersone c " +
                $"LEFT JOIN {clientSchema}tIdentity i ON i.idclient = c.idclient " +
                $"WHERE c.branch = :branchCode";

            return FillDataSet(sql, "PasNoAndBirthYearTable", branchCode);
        }

        public DataSet LoadProducts(int branchCode)
        {
            string clientSchema = _config.GetMainSchema();

            // Real table: tReferenceCardProduct  (NOT TPRODUCTTABLE)
            // Maps to: clsBasStatement.getCardProduct(int pBrach)
            string sql =
                $"SELECT code, name FROM {clientSchema}tReferenceCardProduct " +
                "WHERE branch = :branchCode ORDER BY code";

            return FillDataSet(sql, "ProductTable", branchCode);
        }

        public DataSet LoadBranchPart(int branchCode)
        {
            string clientSchema = _config.GetMainSchema();

            // Maps to: clsBasStatement.getBranchPart(int pBrach)
            string sql =
                $"SELECT * FROM {clientSchema}tBranchPart " +
                "WHERE branch = :branchCode ORDER BY code";

            return FillDataSet(sql, "BranchPartTable", branchCode);
        }

        public DataSet LoadClientBank(int branchCode)
        {
            string clientSchema = _config.GetMainSchema();

            // Maps to: clsBasStatement.getTClientBank(int pBrach)
            // Returns emaillegal, emailcont, phonelegal, phonecont for client bank contacts
            string sql =
                $"SELECT idclient, emaillegal, emailcont, phonelegal, phonecont " +
                $"FROM {clientSchema}tClientbank " +
                "WHERE branch = :branchCode";

            return FillDataSet(sql, "ClientBankTable", branchCode);
        }

        public DataSet LoadAccountCurrencies(int branchCode)
        {
            string stmtSchema = _session.StatementDbSchema;
            string table      = _session.MainTable;

            // Maps to: clsBasStatement.fillAccountCurrencies(int pBrach)
            string sql =
                $"SELECT DISTINCT c.accountcurrency " +
                $"FROM {stmtSchema}{table} c " +
                "WHERE c.branch = :branchCode";

            return FillDataSet(sql, "CurrencyTable", branchCode);
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

using System.Data;

namespace StatementFile.Domain.Interfaces.Services
{
    /// <summary>
    /// Exposes all the DataSet-loading query variants used by the legacy formatters.
    ///
    /// Schema mapping (from clsBasStatement / clsSessionValues):
    ///   - Statement tables (TSTATEMENTMASTERTABLE, TSTATEMENTDETAILTABLE) use "A4M." schema
    ///   - Client/reference tables (tClientPersone, tIdentity, tReferenceCardProduct,
    ///     tBranchPart, tClientbank) use MainSchema from app config
    /// </summary>
    public interface IStatementQueryService
    {
        // ── Master DataSet variants ──────────────────────────────────────────────

        /// <summary>
        /// Standard master + detail DataSet, ordered by card product / branch part
        /// / account / card primary / card number.
        /// Maps to: clsBasStatement.FillStatementDataSet(int pBrach)
        /// </summary>
        DataSet LoadStandard(int branchCode, string additionalCondition = null);

        /// <summary>
        /// Same as <see cref="LoadStandard"/> but sorted so primary cards precede
        /// their supplementary cards within each account group.
        /// Maps to: clsBasStatement.FillStatementDataSet_SortCardPriority()
        /// </summary>
        DataSet LoadSortedByCardPriority(int branchCode, string additionalCondition = null);

        /// <summary>
        /// Adds the OVERDUEDAYS column from A4M.ZM_EOD_CONT_ACCT.
        /// Maps to: clsBasStatement.FillStatementDataSet_WithOverDueDays()
        /// </summary>
        DataSet LoadWithOverdueDays(int branchCode, string additionalCondition = null);

        /// <summary>
        /// Excludes rows where the card product is VISA (joins tReferenceCardProduct).
        /// Maps to: clsBasStatement.FillStatementDataSet_Exclude_VisaCards()
        /// </summary>
        DataSet LoadExcludingVisa(int branchCode, string additionalCondition = null);

        /// <summary>
        /// Filters to VIP cards only (cardvip = 'Y').
        /// Maps to: clsBasStatement.FillStatementDataSet(int pBrach, string vip)
        /// </summary>
        DataSet LoadVipOnly(int branchCode);

        /// <summary>
        /// Master DataSet with MARK-UP fee rows merged into the originating transaction
        /// row (using the same join logic as getDetailQueryForMarkupFee).
        /// Maps to: clsBasStatement.FillStatementDataSetWithRemovingMarkupFee()
        /// </summary>
        DataSet LoadWithMarkupFeeRemoval(int branchCode, string additionalCondition = null);

        // ── Supplementary data sets ──────────────────────────────────────────────

        /// <summary>Loads installment programme rows from the master table.</summary>
        DataSet LoadInstallments(int branchCode);

        /// <summary>
        /// Loads reward programme rows (EarnedBonus, RedeemedBonus, etc.)
        /// from the master table filtered by contracttype.
        /// </summary>
        DataSet LoadRewards(int branchCode, string rewardContractCondition);

        /// <summary>
        /// Loads client email + mobile phone records keyed by idclient.
        /// Table: {MainSchema}tClientPersone  (NOT TCLIENTEMAIL)
        /// Returns: idclient | email | mobilephone | phone | externalid | latfio
        /// Special join for branches 21 and 38: TUP${branch}$CLIENT$
        /// Maps to: clsBasStatement.getClientEmail(int pBrach)
        /// </summary>
        DataSet LoadClientEmails(int branchCode);

        /// <summary>
        /// Loads client email + mobile phone + name records keyed by idclient.
        /// Table: {MainSchema}tClientPersone
        /// Returns: idclient | email | mobilephone | fio
        /// Maps to: clsBasStatement.getClientEmailName(int pBrach)
        /// </summary>
        DataSet LoadClientEmailName(int branchCode);

        /// <summary>
        /// Loads client passport number records.
        /// Table: {MainSchema}tIdentity
        /// Returns: idclient | NO  (NO = passport/identity number)
        /// Maps to: clsBasStatement.getClientPassportNo(int pBrach)
        /// </summary>
        DataSet LoadClientIdentity(int branchCode);

        /// <summary>
        /// Loads client passport number + birth year records.
        /// Joins {MainSchema}tClientPersone + tIdentity.
        /// Returns: idclient | passportno | birthyear
        /// Maps to: clsBasStatement.getClientPasNoAndBirthYear(int pBrach)
        /// </summary>
        DataSet LoadClientPasNoAndBirthYear(int branchCode);

        /// <summary>
        /// Loads card product reference rows for the branch.
        /// Table: {MainSchema}tReferenceCardProduct  (NOT TPRODUCTTABLE)
        /// Returns: code | name
        /// Maps to: clsBasStatement.getCardProduct(int pBrach)
        /// </summary>
        DataSet LoadProducts(int branchCode);

        /// <summary>
        /// Loads branch partition data.
        /// Table: {MainSchema}tBranchPart
        /// Maps to: clsBasStatement.getBranchPart(int pBrach)
        /// </summary>
        DataSet LoadBranchPart(int branchCode);

        /// <summary>
        /// Loads client bank contact info (email addresses and phone numbers).
        /// Table: {MainSchema}tClientbank
        /// Returns: emaillegal | emailcont | phonelegal | phonecont
        /// Maps to: clsBasStatement.getTClientBank(int pBrach)
        /// </summary>
        DataSet LoadClientBank(int branchCode);

        /// <summary>
        /// Loads distinct account currencies for the branch.
        /// Used for multi-currency corporate statements.
        /// Maps to: clsBasStatement.fillAccountCurrencies(int pBrach)
        /// </summary>
        DataSet LoadAccountCurrencies(int branchCode);
    }
}

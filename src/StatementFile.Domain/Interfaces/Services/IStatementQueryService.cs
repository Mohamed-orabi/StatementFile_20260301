using System.Data;

namespace StatementFile.Domain.Interfaces.Services
{
    /// <summary>
    /// Exposes all the DataSet-loading query variants used by the legacy formatters.
    /// Each method maps 1:1 to a specialised FillStatementDataSet_* overload
    /// from clsBasStatement, which is the Oracle data-access root class.
    ///
    /// Keeping them as explicit interface members (rather than parameters) ensures
    /// callers cannot accidentally use the wrong variant and makes unit-testing
    /// straightforward by swapping in a fake implementation.
    /// </summary>
    public interface IStatementQueryService
    {
        // ── Master DataSet variants ──────────────────────────────────────────────

        /// <summary>
        /// Standard master + detail DataSet, ordered by card product / branch part
        /// / account / card primary / card number.
        /// </summary>
        DataSet LoadStandard(int branchCode, string additionalCondition = null);

        /// <summary>
        /// Same as <see cref="LoadStandard"/> but sorted so primary cards precede
        /// their supplementary cards within each account group.
        /// Used by: AIBK Raw Data, AIBK2 Raw Data.
        /// </summary>
        DataSet LoadSortedByCardPriority(int branchCode, string additionalCondition = null);

        /// <summary>
        /// Adds the OVERDUEDAYS column calculated from STETEMENTDUEDATE vs today.
        /// Used by: AAIB Raw Data (Jira AAIB-9308).
        /// </summary>
        DataSet LoadWithOverdueDays(int branchCode, string additionalCondition = null);

        /// <summary>
        /// Excludes rows whose card brand is 'VISA' (joins TPRODUCTTABLE).
        /// Used by: ALXB Retail Raw Data, ALXB Corporate Raw Data.
        /// </summary>
        DataSet LoadExcludingVisa(int branchCode, string additionalCondition = null);

        /// <summary>
        /// Filters to VIP cards only (cardvip = 'Y' or bank-specific VIP flag).
        /// Used by: IDBE XML (branch 16).
        /// </summary>
        DataSet LoadVipOnly(int branchCode);

        // ── Supplementary data sets ──────────────────────────────────────────────

        /// <summary>Loads installment programme rows from the installment tables.</summary>
        DataSet LoadInstallments(int branchCode);

        /// <summary>Loads reward programme rows (EarnedBonus, RedeemedBonus, etc.).</summary>
        DataSet LoadRewards(int branchCode, string rewardContractCondition);

        /// <summary>
        /// Loads client email + mobile phone records keyed by clientid.
        /// Returns: idclient | email | mobilephone.
        /// </summary>
        DataSet LoadClientEmails(int branchCode);

        /// <summary>
        /// Loads client passport number and birth year records.
        /// Used by: AIBK Raw Data, BDCA, EGB.
        /// Returns: idclient | passportno | birthyear | nationalid.
        /// </summary>
        DataSet LoadClientIdentity(int branchCode);

        /// <summary>
        /// Loads product reference rows for the branch.
        /// Used by: BRKA, AIBK Raw Data.
        /// Returns: code | name.
        /// </summary>
        DataSet LoadProducts(int branchCode);
    }
}

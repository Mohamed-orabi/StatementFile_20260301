namespace StatementFile.Domain.Interfaces.Services
{
    /// <summary>
    /// Encapsulates bulk pre-processing operations that must run before
    /// statement generation begins (card-branch matching, Arabic address fix, etc.).
    /// </summary>
    public interface IDataMaintenanceService
    {
        /// <summary>
        /// Aligns the cardbranchpart/cardbranchpartname of every card for a client
        /// to the most recently created card's branch part.
        /// Returns the number of records updated.
        /// </summary>
        int MatchCardBranchForAccount(int branchCode);

        /// <summary>
        /// Detects and fixes garbled Arabic address fields (replaced with "???").
        /// Returns the number of records updated.
        /// </summary>
        int FixArabicAddress(int branchCode);

        /// <summary>
        /// Removes NULL-card rows that are not part of reward or installment programmes.
        /// Returns the number of records deleted.
        /// </summary>
        int CleanNullCards(int branchCode, bool excludeReward, bool excludeInstallment,
                           string installmentCondition);
    }
}

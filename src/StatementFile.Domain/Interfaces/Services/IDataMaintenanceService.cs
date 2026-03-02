namespace StatementFile.Domain.Interfaces.Services
{
    /// <summary>
    /// Encapsulates all bulk pre-processing operations that must run before
    /// statement generation begins for a branch.
    ///
    /// Method mapping to legacy clsMaintainData:
    ///   CleanNullCards()            → DELETE rows where cardno IS NULL
    ///   MatchCardBranchForAccount() → UPDATE cardbranchpart / cardbranchpartname
    ///   FixArabicAddress()          → UPDATE address fields with "???" prefix
    ///   DeleteOnHoldTransactions()  → DELETE rows where HOLSTMT = 'Y'
    ///   FixReward()                 → UPDATE reward-programme fields before generation
    /// </summary>
    public interface IDataMaintenanceService
    {
        /// <summary>
        /// Deletes master-table rows where cardno IS NULL.
        /// Rows belonging to reward or installment contracts are excluded as specified.
        /// Uses Oracle parallel hint: DELETE /*+ parallel (...,4) */
        /// Returns the number of rows deleted.
        /// </summary>
        int CleanNullCards(int branchCode, bool excludeReward, bool excludeInstallment,
                           string installmentCondition);

        /// <summary>
        /// Aligns cardbranchpart / cardbranchpartname across all cards for each client
        /// to the most recently created card's branch part.
        /// Executed as a PL/SQL batch (up to 500 UPDATE statements per BEGIN...END block).
        /// Returns the number of records updated.
        /// </summary>
        int MatchCardBranchForAccount(int branchCode);

        /// <summary>
        /// Finds address fields that start with "???" (corrupted Arabic encoding) and
        /// strips the prefix. Applied to customeraddress1/2/3.
        /// Returns the number of records updated.
        /// </summary>
        int FixArabicAddress(int branchCode);

        /// <summary>
        /// Deletes detail-table rows that are marked on-hold (HOLSTMT = 'Y').
        /// Required by: AUB (Branch 25), any bank that sets the HOLSTMT flag.
        /// Legacy: clsMaintainData.deleteOnHoldTrans().
        /// Returns the number of rows deleted.
        /// </summary>
        int DeleteOnHoldTransactions(int branchCode);

        /// <summary>
        /// Runs the reward-programme consistency fix before generating reward statements.
        /// Resets / recalculates reward fields that may be stale from previous runs.
        /// Legacy: clsMaintainData.fixReward().
        /// Returns the number of records updated.
        /// </summary>
        int FixReward(int branchCode, string rewardContractCondition);
    }
}

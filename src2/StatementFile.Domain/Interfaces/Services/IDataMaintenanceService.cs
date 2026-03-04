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
    ///   DeleteOnHoldTransactions()  → DELETE rows WHERE POSTINGDATE IS NULL AND DOCNO IS NULL
    ///   FixReward()                 → UPDATE reward-programme fields before generation
    ///   FixAddress()                → Split long addresses (>50 chars) into addr1 + addr2
    ///   FixArabicAddressLang()      → Set companycode=1 (Arabic) or companycode=0 (non-Arabic)
    ///   MergeMarkUpFees()           → Consolidate Mark-Up Fee transactions per docno
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
        /// strips the prefix. Applied to all 9 address fields
        /// (customeraddress1/2/3, accountaddress1/2/3, cardaddress1/2/3).
        /// Returns the number of records updated.
        /// </summary>
        int FixArabicAddress(int branchCode);

        /// <summary>
        /// Deletes detail-table rows where POSTINGDATE IS NULL AND DOCNO IS NULL.
        /// When isReward is true, also excludes rows where trandescription = 'Calculated Points'.
        /// Legacy: clsMaintainData.deleteOnHoldTrans(int pBranch, bool isReward).
        /// Returns the number of rows deleted.
        /// </summary>
        int DeleteOnHoldTransactions(int branchCode, bool isReward = false);

        /// <summary>
        /// Runs the reward-programme consistency fix before generating reward statements.
        /// Resets / recalculates reward fields that may be stale from previous runs.
        /// Legacy: clsMaintainData.fixReward().
        /// Returns the number of records updated.
        /// </summary>
        int FixReward(int branchCode, string rewardContractCondition);

        /// <summary>
        /// For each record where customeraddress1 exceeds 50 characters and
        /// customeraddress2 is NULL, splits the address at a word boundary into
        /// customeraddress1/2, accountaddress1/2, and cardaddress1/2.
        /// Legacy: clsMaintainData.fixAddress(int pBankCode).
        /// Returns the number of records updated.
        /// </summary>
        int FixAddress(int branchCode);

        /// <summary>
        /// Sets companycode = 1 (Arabic) or companycode = 0 (non-Arabic) for each
        /// record, based on whether the customer address contains Arabic characters.
        /// Also strips the "???" corruption prefix where present.
        /// Legacy: clsMaintainData.fixArbicAddressLang(int pBankCode).
        /// Returns the number of records updated.
        /// </summary>
        int FixArabicAddressLang(int branchCode);

        /// <summary>
        /// Groups Mark-Up Fee transactions in the detail table by docno, accumulates
        /// the total billtranamount onto the first row, and deletes duplicate rows.
        /// Legacy: clsMaintainData.mergeMarkUpFees(int pBranch).
        /// </summary>
        void MergeMarkUpFees(int branchCode);
    }
}

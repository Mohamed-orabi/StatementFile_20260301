using System;

namespace StatementFile.Domain.Interfaces.Services
{
    /// <summary>
    /// Records a statement generation run in the a4m.MSCC_PROD_STAT_MASTER audit table.
    ///
    /// Called after each successful statement generation run.
    /// Legacy: clsStatementSummary.InsertRecordDb()
    ///
    /// Target table: a4m.MSCC_PROD_STAT_MASTER
    /// Columns: branch, product_code, product_name, (null), stat_m (yyyyMM),
    ///          count_stat, count_int, creation_date
    /// </summary>
    public interface IStatementSummaryService
    {
        /// <summary>
        /// Inserts one audit row for the completed statement run.
        /// </summary>
        void InsertRecord(
            int      bankCode,
            int      productCode,
            string   productName,
            DateTime statementDate,
            int      noOfStatements,
            int      noOfTransactions,
            DateTime creationDate);

        /// <summary>
        /// Updates an existing audit row by incrementing count_stat and count_int.
        /// Used when a run is split across multiple batches.
        /// Legacy: clsStatementSummary.UpdateRecordDb()
        /// </summary>
        void UpdateRecord(
            int      bankCode,
            int      productCode,
            DateTime statementDate,
            int      additionalStatements,
            int      additionalTransactions);
    }
}

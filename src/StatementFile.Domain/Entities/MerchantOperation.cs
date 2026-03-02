using System;

namespace StatementFile.Domain.Entities
{
    /// <summary>
    /// A single merchant transaction line (Operation) belonging to a MerchantStatement.
    /// Corresponds to the "Operation" table in MerchantStatementTemplate.mdb.
    /// </summary>
    public class MerchantOperation
    {
        public int      StatDetailCode  { get; private set; }
        public int      StatMasterCode  { get; private set; }
        public int      Branch          { get; private set; }

        // Core fields (single-letter column names preserved from legacy schema)
        public string   D               { get; private set; }  // Description
        public decimal  O               { get; private set; }  // Original amount
        public decimal  A               { get; private set; }  // Amount
        public decimal  OA              { get; private set; }  // Other amount
        public decimal  CF              { get; private set; }  // Commission fee
        public decimal  S               { get; private set; }  // Settlement
        public DateTime OD              { get; private set; }  // Operation date
        public DateTime TD              { get; private set; }  // Transaction date

        private MerchantOperation() { }

        public static MerchantOperation Create(
            int statDetailCode, int statMasterCode, int branch,
            string description,
            decimal originalAmount, decimal amount,
            decimal otherAmount, decimal commissionFee, decimal settlement,
            DateTime operationDate, DateTime transactionDate)
        {
            return new MerchantOperation
            {
                StatDetailCode = statDetailCode,
                StatMasterCode = statMasterCode,
                Branch         = branch,
                D              = description,
                O              = originalAmount,
                A              = amount,
                OA             = otherAmount,
                CF             = commissionFee,
                S              = settlement,
                OD             = operationDate,
                TD             = transactionDate,
            };
        }
    }
}

using System;
using System.Collections.Generic;

namespace StatementFile.Domain.Entities
{
    /// <summary>
    /// Aggregate root for a merchant's billing statement.
    /// Loaded from an XML source file and persisted into an MS-Access template (.mdb).
    /// </summary>
    public class MerchantStatement
    {
        public int      StatMasterCode  { get; private set; }
        public int      Branch          { get; private set; }
        public string   BankName        { get; private set; }
        public string   BankFullName    { get; private set; }
        public DateTime StatDate        { get; private set; }
        public string   StatementNo     { get; private set; }
        public string   Account         { get; private set; }
        public string   ExternalAccount { get; private set; }
        public DateTime StartDate       { get; private set; }
        public DateTime EndDate         { get; private set; }

        private readonly List<MerchantOperation> _operations = new List<MerchantOperation>();
        public IReadOnlyList<MerchantOperation> Operations => _operations.AsReadOnly();

        private MerchantStatement() { }

        public static MerchantStatement Create(
            int statMasterCode, int branch,
            string bankName, string bankFullName,
            DateTime statDate, string statementNo,
            string account, string externalAccount,
            DateTime startDate, DateTime endDate)
        {
            return new MerchantStatement
            {
                StatMasterCode  = statMasterCode,
                Branch          = branch,
                BankName        = bankName,
                BankFullName    = bankFullName,
                StatDate        = statDate,
                StatementNo     = statementNo,
                Account         = account,
                ExternalAccount = string.IsNullOrEmpty(externalAccount) ? account : externalAccount,
                StartDate       = startDate,
                EndDate         = endDate,
            };
        }

        public void AddOperation(MerchantOperation operation)
        {
            if (operation == null) throw new ArgumentNullException(nameof(operation));
            _operations.Add(operation);
        }
    }
}

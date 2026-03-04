using System;

namespace StatementFile.Domain.Entities
{
    /// <summary>
    /// Represents a single transaction line (debit/credit) within a statement period.
    /// Sourced from TSTATEMENTDETAILTABLE.
    /// </summary>
    public class StatementTransaction
    {
        public string  StatementNo                  { get; private set; }
        public string  StatementNumber              { get; private set; }
        public string  AccountNo                    { get; private set; }
        public string  AccountCurrency              { get; private set; }
        public string  CardNo                       { get; private set; }

        public DateTime TransDate                   { get; private set; }
        public DateTime PostingDate                 { get; private set; }

        public string  TranDescription              { get; private set; }
        public string  Merchant                     { get; private set; }
        public decimal OrigTranAmount               { get; private set; }
        public string  OrigTranCurrency             { get; private set; }
        public decimal BillTranAmount               { get; private set; }
        public string  BillTranAmountSign           { get; private set; }
        public string  DocNo                        { get; private set; }
        public string  ReferenceNo                  { get; private set; }
        public string  ApprovalCode                 { get; private set; }
        public string  PackageName                  { get; private set; }
        public long    EntryNo                      { get; private set; }

        // Installment-specific fields
        public string  InstallmentData              { get; private set; }
        public string  OrigDocNo                    { get; private set; }
        public string  InstallmentMerchant          { get; private set; }
        public string  InstallmentMerchantLocation  { get; private set; }
        public DateTime? InstallmentStartDate        { get; private set; }
        public decimal InstallmentLoan              { get; private set; }
        public int     InstallmentCycleNo           { get; private set; }
        public int     InstallmentCycles            { get; private set; }
        public decimal InstallmentRegRePayment      { get; private set; }
        public DateTime? InstallmentEndDate          { get; private set; }
        public decimal InstallmentOutBalance        { get; private set; }

        private StatementTransaction() { }

        public static StatementTransaction Create(
            string statementNo, string statementNumber,
            string accountNo, string accountCurrency, string cardNo,
            DateTime transDate, DateTime postingDate,
            string tranDescription, string merchant,
            decimal origTranAmount, string origTranCurrency,
            decimal billTranAmount, string billTranAmountSign,
            string docNo, string referenceNo, string approvalCode,
            string packageName, long entryNo,
            string installmentData, string origDocNo,
            string installmentMerchant, string installmentMerchantLocation,
            DateTime? installmentStartDate, decimal installmentLoan,
            int installmentCycleNo, int installmentCycles,
            decimal installmentRegRePayment, DateTime? installmentEndDate,
            decimal installmentOutBalance)
        {
            return new StatementTransaction
            {
                StatementNo = statementNo,
                StatementNumber = statementNumber,
                AccountNo = accountNo,
                AccountCurrency = accountCurrency,
                CardNo = cardNo,
                TransDate = transDate,
                PostingDate = postingDate,
                TranDescription = tranDescription,
                Merchant = merchant,
                OrigTranAmount = origTranAmount,
                OrigTranCurrency = origTranCurrency,
                BillTranAmount = billTranAmount,
                BillTranAmountSign = billTranAmountSign,
                DocNo = docNo,
                ReferenceNo = referenceNo,
                ApprovalCode = approvalCode,
                PackageName = packageName,
                EntryNo = entryNo,
                InstallmentData = installmentData,
                OrigDocNo = origDocNo,
                InstallmentMerchant = installmentMerchant,
                InstallmentMerchantLocation = installmentMerchantLocation,
                InstallmentStartDate = installmentStartDate,
                InstallmentLoan = installmentLoan,
                InstallmentCycleNo = installmentCycleNo,
                InstallmentCycles = installmentCycles,
                InstallmentRegRePayment = installmentRegRePayment,
                InstallmentEndDate = installmentEndDate,
                InstallmentOutBalance = installmentOutBalance,
            };
        }
    }
}

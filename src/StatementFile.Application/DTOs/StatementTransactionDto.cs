using System;

namespace StatementFile.Application.DTOs
{
    /// <summary>
    /// Read-only projection of a StatementTransaction for use cases and UI rendering.
    /// </summary>
    public sealed class StatementTransactionDto
    {
        public string   StatementNo        { get; set; }
        public string   CardNo             { get; set; }
        public string   AccountNo          { get; set; }
        public DateTime TransDate          { get; set; }
        public DateTime PostingDate        { get; set; }
        public string   TranDescription    { get; set; }
        public string   Merchant           { get; set; }
        public decimal  OrigTranAmount     { get; set; }
        public string   OrigTranCurrency   { get; set; }
        public decimal  BillTranAmount     { get; set; }
        public string   BillTranAmountSign { get; set; }
        public string   DocNo              { get; set; }
        public string   ApprovalCode       { get; set; }
        public string   PackageName        { get; set; }
        public long     EntryNo            { get; set; }
        public string   InstallmentData    { get; set; }
    }
}

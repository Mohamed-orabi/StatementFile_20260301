using System;
using System.Collections.Generic;

namespace StatementFile.Application.DTOs
{
    public sealed class MerchantStatementDto
    {
        public int      Branch          { get; set; }
        public string   BankName        { get; set; }
        public string   BankFullName    { get; set; }
        public DateTime StatDate        { get; set; }
        public string   StatementNo     { get; set; }
        public string   Account         { get; set; }
        public string   ExternalAccount { get; set; }
        public DateTime StartDate       { get; set; }
        public DateTime EndDate         { get; set; }
        public List<MerchantOperationDto> Operations { get; set; } = new List<MerchantOperationDto>();
    }

    public sealed class MerchantOperationDto
    {
        public string   Description    { get; set; }
        public decimal  Amount         { get; set; }
        public decimal  CommissionFee  { get; set; }
        public DateTime OperationDate  { get; set; }
        public DateTime TransactionDate { get; set; }
    }
}

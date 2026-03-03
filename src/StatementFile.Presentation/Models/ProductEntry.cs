using System;
using StatementFile.Domain.Enums;

namespace StatementFile.Presentation.Models
{
    /// <summary>
    /// UI model representing one bank/product combination in the generation queue.
    /// ConfigId identifies the database record; the API uses it to load the full
    /// configuration when running generation.
    /// </summary>
    public sealed class ProductEntry
    {
        public int    ConfigId            { get; set; }
        public bool   Selected            { get; set; }
        public string BankName            { get; set; } = string.Empty;
        public string BankFullName        { get; set; } = string.Empty;
        public int    BranchCode          { get; set; }
        public string CardProduct         { get; set; } = string.Empty;
        public StatementOutputType OutputType { get; set; } = StatementOutputType.Html;
        public string FormatterKey        { get; set; } = string.Empty;
        public string StatementTypeSuffix { get; set; } = "CR";
        public ProcessingMode Modes       { get; set; } = ProcessingMode.Standard;
        public DateTime StatementDate     { get; set; } = DateTime.Today;

        // Display-only fields (the API reloads authoritative values from the database)
        public string EmailFromAddress        { get; set; } = "cardservices@emp-group.com";
        public string RewardContractCondition { get; set; } = "'New Reward Contract'";

        public override string ToString() => $"{BankName} — {CardProduct}";
    }
}

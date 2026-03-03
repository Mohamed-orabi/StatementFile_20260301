using System;
using StatementFile.Application.UseCases.StatementGeneration;
using StatementFile.Domain.Enums;

namespace StatementFile.Presentation.Models
{
    /// <summary>
    /// UI model representing one bank/product combination in the generation list.
    /// Replaces the <c>ProductSelection</c> class used by the WinForms CheckedListBox.
    /// Each entry maps directly to one <see cref="GenerateStatementCommand"/>.
    /// </summary>
    public sealed class ProductEntry
    {
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

        // Optional overrides (defaults match the command defaults)
        public string EmailFromAddress        { get; set; } = "cardservices@emp-group.com";
        public string RewardContractCondition { get; set; } = "'New Reward Contract'";

        public override string ToString() => $"{BankName} — {CardProduct}";

        /// <summary>
        /// Builds the Application-layer command from this UI model.
        /// </summary>
        public GenerateStatementCommand BuildCommand(string outputRootPath) =>
            new GenerateStatementCommand(
                branchCode:             BranchCode,
                bankName:               BankName,
                bankFullName:           BankFullName,
                cardProduct:            CardProduct,
                outputType:             OutputType,
                formatterKey:           FormatterKey,
                processingModes:        Modes,
                cardType:               CardType.Credit,
                statementType:          StatementType.Normal,
                statementDate:          StatementDate,
                outputRootPath:         outputRootPath,
                statementTypeSuffix:    StatementTypeSuffix,
                emailFromAddress:       EmailFromAddress,
                rewardContractCondition: RewardContractCondition);
    }
}

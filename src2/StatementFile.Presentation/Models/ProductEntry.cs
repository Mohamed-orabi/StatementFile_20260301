using System;
using StatementFile.Application.DTOs;
using StatementFile.Domain.Enums;

namespace StatementFile.Presentation.Models
{
    /// <summary>
    /// UI model representing one bank / product combination in the generation queue.
    ///
    /// Each entry maps to one <see cref="GenerateStatementApiRequest"/> that is POSTed
    /// to the API when the user clicks Generate.
    /// </summary>
    public sealed class ProductEntry
    {
        public bool   Selected            { get; set; }
        public string BankName            { get; set; } = string.Empty;
        public string BankFullName        { get; set; } = string.Empty;
        public int    BranchCode          { get; set; }
        public string CardProduct         { get; set; } = string.Empty;
        public int    OutputType          { get; set; }   // StatementOutputType as int
        public string FormatterKey        { get; set; } = string.Empty;
        public string StatementTypeSuffix { get; set; } = "CR";
        public long   ProcessingModes     { get; set; }   // ProcessingMode flags as long
        public DateTime StatementDate     { get; set; } = DateTime.Today;

        // Optional overrides (defaults match GenerateStatementCommand defaults)
        public string EmailFromAddress        { get; set; } = "cardservices@emp-group.com";
        public string RewardContractCondition { get; set; } = "'New Reward Contract'";

        public override string ToString() => $"{BankName} — {CardProduct}";

        /// <summary>
        /// Builds the API request from this UI model.
        /// <paramref name="outputRootPath"/> is optional — when null or empty
        /// the API will use its configured default path.
        /// </summary>
        public GenerateStatementApiRequest BuildApiRequest(string outputRootPath = null) =>
            new GenerateStatementApiRequest
            {
                BranchCode                   = BranchCode,
                BankName                     = BankName,
                BankFullName                 = BankFullName,
                CardProduct                  = CardProduct,
                OutputType                   = OutputType,
                FormatterKey                 = FormatterKey,
                ProcessingModes              = ProcessingModes,
                CardType                     = (int)CardType.Credit,
                StatementType                = (int)StatementType.Normal,
                StatementDate                = StatementDate,
                OutputRootPath               = outputRootPath,
                StatementTypeSuffix          = StatementTypeSuffix,
                EmailFromAddress             = EmailFromAddress,
                RewardContractCondition      = RewardContractCondition,
            };
    }
}

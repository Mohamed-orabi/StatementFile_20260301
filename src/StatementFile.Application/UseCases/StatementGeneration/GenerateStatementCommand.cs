using System;
using StatementFile.Domain.Enums;

namespace StatementFile.Application.UseCases.StatementGeneration
{
    /// <summary>
    /// Command that drives a single bank/product statement generation run.
    /// </summary>
    public sealed class GenerateStatementCommand
    {
        public int          BranchCode        { get; }
        public string       BankName          { get; }
        public string       BankFullName      { get; }
        public string       CardProduct       { get; }
        public CardType     CardType          { get; }
        public StatementType StatementType    { get; }
        public string       FormatterKey      { get; }    // e.g. "HTML_BAI", "TXT_UBA"
        public DateTime     StatementDate     { get; }
        public string       OutputRootPath    { get; }
        public bool         AppendMode        { get; }    // merge with existing data

        public GenerateStatementCommand(
            int          branchCode,
            string       bankName,
            string       bankFullName,
            string       cardProduct,
            CardType     cardType,
            StatementType statementType,
            string       formatterKey,
            DateTime     statementDate,
            string       outputRootPath,
            bool         appendMode = false)
        {
            BranchCode    = branchCode;
            BankName      = bankName;
            BankFullName  = bankFullName;
            CardProduct   = cardProduct;
            CardType      = cardType;
            StatementType = statementType;
            FormatterKey  = formatterKey;
            StatementDate = statementDate;
            OutputRootPath = outputRootPath;
            AppendMode    = appendMode;
        }
    }
}

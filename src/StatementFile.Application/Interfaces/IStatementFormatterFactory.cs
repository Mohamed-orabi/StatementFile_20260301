using StatementFile.Domain.Interfaces.Services;

namespace StatementFile.Application.Interfaces
{
    /// <summary>
    /// Factory that resolves the correct <see cref="IStatementFormatterService"/>
    /// for a given bank/product key.
    /// Keeps the selection logic database-driven via the formatter key stored in
    /// the product configuration rather than hardcoded if/else chains.
    /// </summary>
    public interface IStatementFormatterFactory
    {
        /// <summary>
        /// Returns the formatter registered for <paramref name="formatterKey"/>.
        /// Key convention: "{FORMAT}_{BANKCODE}" (e.g. "HTML_BAI", "TXT_UBA").
        /// Falls back to a generic formatter when no exact match is found.
        /// </summary>
        IStatementFormatterService GetFormatter(string formatterKey);
    }
}

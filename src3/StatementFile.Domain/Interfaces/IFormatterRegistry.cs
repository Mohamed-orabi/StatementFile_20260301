using StatementFile.Domain.Interfaces;

namespace StatementFile.Domain.Interfaces
{
    /// <summary>
    /// Resolves a formatter by its string key (e.g. "HTML_UBA", "PDF_BDCA").
    /// Implemented in the Infrastructure layer; registered in DI.
    /// </summary>
    public interface IFormatterRegistry
    {
        IStatementFormatter Resolve(string formatterKey);
        bool IsRegistered(string formatterKey);
    }
}

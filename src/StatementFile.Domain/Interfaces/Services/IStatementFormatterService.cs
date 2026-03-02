using System.Data;

namespace StatementFile.Domain.Interfaces.Services
{
    /// <summary>
    /// Renders statement content into a specific output format (HTML, TXT, etc.).
    /// Each bank/product variation supplies its own implementation via the factory.
    /// </summary>
    public interface IStatementFormatterService
    {
        /// <summary>Key that the factory uses to resolve this formatter (e.g. "HTML_BAI").</summary>
        string FormatterKey { get; }

        /// <summary>
        /// Generates the formatted output files for the supplied statement DataSet
        /// and writes them to <paramref name="outputDirectory"/>.
        /// Returns paths of all files produced.
        /// </summary>
        System.Collections.Generic.IEnumerable<string> Format(
            DataSet statementDataSet,
            string  outputDirectory,
            int     branchCode,
            string  cardProduct);
    }
}

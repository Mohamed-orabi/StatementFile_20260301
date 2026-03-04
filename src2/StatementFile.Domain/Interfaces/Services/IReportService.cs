using System.Collections.Generic;
using System.Data;

namespace StatementFile.Domain.Interfaces.Services
{
    /// <summary>
    /// Abstraction over Crystal Reports / RDLC report rendering and export.
    /// </summary>
    public interface IReportService
    {
        /// <summary>
        /// Renders a report template populated with <paramref name="data"/> and exports
        /// the result to <paramref name="outputPath"/> in the requested <paramref name="format"/>.
        /// </summary>
        /// <param name="reportTemplatePath">Full path to the .rpt or .rdlc template.</param>
        /// <param name="data">DataSet containing all named sub-tables the report expects.</param>
        /// <param name="outputPath">Destination file path (e.g. .pdf, .html, .xls).</param>
        /// <param name="format">Export format token: "PDF", "HTML", "EXCEL".</param>
        /// <param name="selectionFormula">Optional Crystal Reports record-selection formula.</param>
        void Export(
            string reportTemplatePath,
            DataSet data,
            string outputPath,
            string format,
            string selectionFormula = null);
    }
}

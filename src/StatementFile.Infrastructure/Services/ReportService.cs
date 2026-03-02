using System;
using System.Data;
using System.IO;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using StatementFile.Domain.Interfaces.Services;

namespace StatementFile.Infrastructure.Services
{
    /// <summary>
    /// Crystal Reports implementation of <see cref="IReportService"/>.
    /// Wraps the legacy CrystalDecisions API; the report template path and
    /// export format are resolved from the use-case command, keeping this
    /// class entirely format-agnostic.
    /// </summary>
    public sealed class ReportService : IReportService
    {
        public void Export(
            string  reportTemplatePath,
            DataSet data,
            string  outputPath,
            string  format,
            string  selectionFormula = null)
        {
            if (!File.Exists(reportTemplatePath))
                throw new FileNotFoundException($"Report template not found: {reportTemplatePath}");

            using (var report = new ReportDocument())
            {
                report.Load(reportTemplatePath);
                report.SetDataSource(data);

                if (!string.IsNullOrWhiteSpace(selectionFormula))
                    report.RecordSelectionFormula = selectionFormula;

                ExportFormatType exportFormat = ResolveFormat(format);

                report.ExportToDisk(exportFormat, outputPath);
            }
        }

        private static ExportFormatType ResolveFormat(string format)
        {
            switch ((format ?? string.Empty).ToUpperInvariant())
            {
                case "PDF":   return ExportFormatType.PortableDocFormat;
                case "HTML":  return ExportFormatType.HTML40;
                case "EXCEL":
                case "XLS":   return ExportFormatType.Excel;
                default:      return ExportFormatType.PortableDocFormat;
            }
        }
    }
}

using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using StatementFile.Application.Interfaces;
using StatementFile.Application.UseCases.StatementGeneration;

namespace StatementFile.Infrastructure.Formatters
{
    /// <summary>
    /// Fallback HTML formatter used when no bank-specific formatter is registered.
    /// Produces a simple but valid HTML statement file — bank-specific classes
    /// in the Banks/ folder extend this via the legacy inheritance hierarchy and
    /// can be registered individually in the factory.
    /// </summary>
    public sealed class GenericHtmlStatementFormatter : IStatementFormatterService
    {
        public string FormatterKey => "HTML_GENERIC";

        public IEnumerable<string> Format(
            StatementDataContext     context,
            string                   outputDirectory,
            GenerateStatementCommand command)
        {
            string outputFile = Path.Combine(
                outputDirectory,
                $"Statement_{command.BranchCode}_{command.CardProduct}.html");

            DataSet masterDataSet = context?.MasterDataSet;

            var sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html><head><meta charset='UTF-8'>");
            sb.AppendLine("<style>body{font-family:Arial,sans-serif;font-size:11px;}" +
                          "table{border-collapse:collapse;width:100%}" +
                          "th,td{border:1px solid #ccc;padding:4px;}</style>");
            sb.AppendLine("</head><body>");

            if (masterDataSet?.Tables?.Count > 0)
            {
                foreach (DataTable table in masterDataSet.Tables)
                {
                    sb.AppendLine($"<h3>{table.TableName}</h3><table><tr>");
                    foreach (DataColumn col in table.Columns)
                        sb.Append($"<th>{col.ColumnName}</th>");
                    sb.AppendLine("</tr>");

                    foreach (DataRow row in table.Rows)
                    {
                        sb.Append("<tr>");
                        foreach (DataColumn col in table.Columns)
                            sb.Append($"<td>{row[col]}</td>");
                        sb.AppendLine("</tr>");
                    }
                    sb.AppendLine("</table><br/>");
                }
            }

            sb.AppendLine("</body></html>");
            File.WriteAllText(outputFile, sb.ToString(), Encoding.UTF8);

            return new[] { outputFile };
        }
    }
}

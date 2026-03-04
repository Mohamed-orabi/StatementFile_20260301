using System;
using System.Data;
using System.IO;
using Oracle.ManagedDataAccess.Client;
using StatementFile.Domain.Entities;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Infrastructure.Formatters
{
    /// <summary>
    /// Exports raw statement data as CSV.
    /// Equivalent to the legacy <c>clsStatement_ExportRpt</c> /
    /// <c>clsStatement_ExportReward</c> classes used in frmStatementFileExtn.
    /// </summary>
    public sealed class RawDataExportFormatter : IStatementFormatter
    {
        public FormatterResult Format(BankProductConfig config, FormatterContext ctx)
        {
            string suffix      = ResolveTableSuffix(config);
            string masterTable = $"TSTATEMENTMASTER{suffix}";

            try
            {
                using var conn = new OracleConnection(ctx.ConnectionString);
                conn.Open();

                var sql = $@"SELECT M.* FROM {masterTable} M
                             WHERE M.BRANCH_CODE = :bc
                               AND TRUNC(M.STATEMENT_DATE) = TRUNC(:sd)";

                if (!string.IsNullOrWhiteSpace(config.WhereCondition))
                    sql += $" AND ({config.WhereCondition})";

                using var cmd = new OracleCommand(sql, conn);
                cmd.Parameters.Add(":bc", config.BranchCode);
                cmd.Parameters.Add(":sd", ctx.StatementDate);

                using var adapter = new OracleDataAdapter(cmd);
                var dt = new DataTable();
                adapter.Fill(dt);

                var fileName = $"{config.BankName}_RawData_{ctx.StatementDate:yyyyMM}.csv";
                var filePath = Path.Combine(ctx.OutputDirectory, fileName);

                WriteCsv(dt, filePath, ctx.AppendData);

                return new FormatterResult
                {
                    Success         = true,
                    FilesGenerated  = 1,
                    StatementsCount = dt.Rows.Count,
                    OutputDirectory = ctx.OutputDirectory,
                };
            }
            catch (Exception ex)
            {
                return new FormatterResult { Success = false, Error = ex.Message };
            }
        }

        private static string ResolveTableSuffix(BankProductConfig c)
        {
            string s = c.StatementTypeSuffix?.ToUpperInvariant() ?? "CR";
            if (s.StartsWith("DB")) return "DB";
            if (s.StartsWith("CP")) return "CP";
            return "CR";
        }

        private static void WriteCsv(DataTable dt, string filePath, bool append)
        {
            bool writeHeader = !append || !File.Exists(filePath);
            using var sw = new StreamWriter(filePath, append, System.Text.Encoding.UTF8);

            if (writeHeader)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (i > 0) sw.Write(",");
                    sw.Write(CsvEscape(dt.Columns[i].ColumnName));
                }
                sw.WriteLine();
            }

            foreach (DataRow row in dt.Rows)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (i > 0) sw.Write(",");
                    sw.Write(CsvEscape(row[i]?.ToString()));
                }
                sw.WriteLine();
            }
        }

        private static string CsvEscape(string value)
        {
            if (value == null) return string.Empty;
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                return $"\"{value.Replace("\"", "\"\"")}\"";
            return value;
        }
    }
}

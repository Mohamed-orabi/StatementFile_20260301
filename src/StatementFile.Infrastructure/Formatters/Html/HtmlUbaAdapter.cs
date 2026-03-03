using System;
using System.Collections.Generic;
using StatementFile.Application.UseCases.StatementGeneration;

namespace StatementFile.Infrastructure.Formatters.Html
{
    /// <summary>HTML e-statement adapter for UBA (United Bank for Africa).</summary>
    public sealed class HtmlUbaAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "HTML_UBA";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fileBasePath, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            var legacy    = new clsStatHtmlUBA();
            legacy.Statement(
                fileBasePath,
                cmd.BankName,
                cmd.BranchCode,
                cmd.StatementTypeSuffix,
                cmd.StatementDate,
                cmd.StatementTypeSuffix,
                cmd.AppendMode,
                pReportName: string.Empty);

            return CollectOutputFiles(
                System.IO.Path.Combine(fileBasePath,
                    cmd.StatementDate.ToString("yyyyMM") + cmd.BankName + "_" + cmd.StatementTypeSuffix),
                startedAt);
        }
    }
}

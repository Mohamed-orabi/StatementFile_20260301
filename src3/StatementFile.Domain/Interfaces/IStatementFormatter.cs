using System;
using StatementFile.Domain.Entities;

namespace StatementFile.Domain.Interfaces
{
    /// <summary>
    /// Contract implemented by every bank-specific statement formatter.
    ///
    /// Replaces the per-case object instantiation pattern in frmStatementFileExtn:
    ///   var obj = new clsStatHtmlUBA();
    ///   obj.emailFrom = ...;
    ///   obj.Statement(outputDir, bankName, bankCode, fileName, date, email, appendData, report);
    ///
    /// Each formatter is registered in <see cref="StatementFile.Infrastructure.Formatters.FormatterRegistry"/>
    /// under its <see cref="BankProductConfig.FormatterKey"/>.
    /// </summary>
    public interface IStatementFormatter
    {
        /// <summary>
        /// Generates the statement output for the given configuration and statement date.
        /// </summary>
        /// <param name="config">The active bank/product configuration row.</param>
        /// <param name="context">Runtime parameters: output directory, date, optional email override.</param>
        /// <returns>Result describing how many files were produced and optional errors.</returns>
        FormatterResult Format(BankProductConfig config, FormatterContext context);
    }

    /// <summary>Runtime parameters passed into every formatter call.</summary>
    public sealed class FormatterContext
    {
        public DateTime StatementDate   { get; init; }
        public string   OutputDirectory { get; init; }
        public string   ReportFilePath  { get; init; }  // path to .rpt Crystal Reports template
        public string   EmailOverride   { get; init; }  // force delivery to this address (testing)
        public bool     AppendData      { get; init; }  // append to existing output (legacy flag)
        public string   ConnectionString{ get; init; }
    }

    /// <summary>Outcome reported back from a formatter run.</summary>
    public sealed class FormatterResult
    {
        public bool   Success           { get; init; }
        public string Error             { get; init; }
        public int    FilesGenerated    { get; init; }
        public int    EmailsSent        { get; init; }
        public int    StatementsCount   { get; init; }
        public string OutputDirectory   { get; init; }
    }
}

using System;
using System.IO;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Application.Commands
{
    /// <summary>
    /// Orchestrates a single statement generation run.
    ///
    /// Steps (mirrors the inner loop of frmStatementFile.createAllStat / runStatement):
    ///   1. Resolve the formatter from the registry using config.FormatterKey.
    ///   2. Ensure the output directory exists.
    ///   3. Call IStatementFormatter.Format().
    ///   4. Return the result.
    /// </summary>
    public sealed class GenerateStatementHandler
    {
        private readonly IFormatterRegistry _formatters;

        public GenerateStatementHandler(IFormatterRegistry formatters)
        {
            _formatters = formatters ?? throw new ArgumentNullException(nameof(formatters));
        }

        public GenerationResult Handle(GenerateStatementCommand cmd)
        {
            if (cmd == null) throw new ArgumentNullException(nameof(cmd));

            var formatter = _formatters.Resolve(cmd.Config.FormatterKey);
            if (formatter == null)
                return GenerationResult.Failure($"No formatter registered for key '{cmd.Config.FormatterKey}'.");

            string outputDir = string.IsNullOrWhiteSpace(cmd.OutputDirectory)
                ? Path.Combine(Path.GetTempPath(), "StatementOutput")
                : cmd.OutputDirectory;

            Directory.CreateDirectory(outputDir);

            var context = new FormatterContext
            {
                StatementDate    = cmd.StatementDate,
                OutputDirectory  = outputDir,
                ReportFilePath   = cmd.ReportTemplate,
                EmailOverride    = cmd.EmailOverride,
                AppendData       = cmd.AppendData,
                ConnectionString = cmd.ConnectionString,
            };

            try
            {
                var result = formatter.Format(cmd.Config, context);
                return result.Success
                    ? GenerationResult.Ok(result.FilesGenerated, result.EmailsSent, result.StatementsCount, result.OutputDirectory)
                    : GenerationResult.Failure(result.Error);
            }
            catch (Exception ex)
            {
                return GenerationResult.Failure(ex.Message);
            }
        }
    }

    public sealed class GenerationResult
    {
        public bool   IsSuccess        { get; private set; }
        public string Error            { get; private set; }
        public int    FilesGenerated   { get; private set; }
        public int    EmailsSent       { get; private set; }
        public int    StatementsCount  { get; private set; }
        public string OutputDirectory  { get; private set; }

        public static GenerationResult Ok(int files, int emails, int statements, string dir) =>
            new GenerationResult
            {
                IsSuccess       = true,
                FilesGenerated  = files,
                EmailsSent      = emails,
                StatementsCount = statements,
                OutputDirectory = dir,
            };

        public static GenerationResult Failure(string error) =>
            new GenerationResult { IsSuccess = false, Error = error };
    }
}

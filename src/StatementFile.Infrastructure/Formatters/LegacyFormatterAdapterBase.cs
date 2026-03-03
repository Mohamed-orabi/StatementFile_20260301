using System;
using System.Collections.Generic;
using System.IO;
using StatementFile.Application.UseCases.StatementGeneration;
using StatementFile.Domain.Interfaces.Services;

namespace StatementFile.Infrastructure.Formatters
{
    /// <summary>
    /// Base class for all legacy formatter adapters.
    ///
    /// The legacy bank classes (clsStatHtmlUBA, clsStatHtmlABP, clsStatRawDataAIBK, etc.)
    /// each expose a Statement(pStrFileName, pBankName, pBankCode, pStrFile, pDate,
    /// pStmntType, pAppendData [, pReportName]) method. This adapter:
    ///
    ///   1. Translates the Clean-Architecture command into the legacy parameter list.
    ///   2. Calls the legacy Statement() method.
    ///   3. Collects the file paths produced by the legacy code.
    ///   4. Returns the file paths to the use-case handler.
    ///
    /// The legacy classes are NOT modified in any way — this is the Adapter pattern.
    /// Each concrete subclass instantiates exactly one legacy class and delegates to it.
    /// </summary>
    public abstract class LegacyFormatterAdapterBase : IStatementFormatterService
    {
        public abstract string FormatterKey { get; }

        public IEnumerable<string> Format(
            StatementDataContext     context,
            string                   outputDirectory,
            GenerateStatementCommand command)
        {
            // Build the file base path the legacy code expects:
            //   {outputRootPath}  (already includes trailing backslash from clsBasFile.makeStrAsPath)
            string fileBasePath = EnsureTrailingBackslash(command.OutputRootPath);

            var files = FormatCore(context, fileBasePath, command);

            return files;
        }

        /// <summary>
        /// Implemented by each concrete adapter to call the legacy Statement() method
        /// and return the paths of files produced.
        /// </summary>
        protected abstract IEnumerable<string> FormatCore(
            StatementDataContext     context,
            string                   fileBasePath,
            GenerateStatementCommand command);

        // ── Helpers used by derived adapters ──────────────────────────────────────

        /// <summary>
        /// Collects all files in the given directory that were produced during the
        /// current run by comparing modification timestamps.
        /// </summary>
        protected static IEnumerable<string> CollectOutputFiles(
            string outputDirectory,
            DateTime runStartedAt)
        {
            if (!Directory.Exists(outputDirectory))
                yield break;

            foreach (string file in Directory.GetFiles(outputDirectory, "*", SearchOption.AllDirectories))
            {
                if (File.GetLastWriteTime(file) >= runStartedAt)
                    yield return file;
            }
        }

        private static string EnsureTrailingBackslash(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return path;
            return path.TrimEnd('\\', '/') + "\\";
        }
    }
}

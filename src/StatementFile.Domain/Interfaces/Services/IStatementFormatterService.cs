using System.Collections.Generic;
using System.Data;
using StatementFile.Application.UseCases.StatementGeneration;

namespace StatementFile.Domain.Interfaces.Services
{
    /// <summary>
    /// Renders statement content into a specific physical output format.
    ///
    /// One implementation exists per bank/output-format combination, registered
    /// in the StatementFormatterFactory by their FormatterKey.
    ///
    /// Key convention: "{OutputType}_{BankCode}[_{Variant}]"
    ///   Examples:
    ///     "HTML_UBA"          – United Bank for Africa HTML e-statement
    ///     "HTML_ABP_SUP"      – Access Bank prepaid supplementary card variant
    ///     "HTML_BAI_PREPAID"  – BAI Angola prepaid card (Portuguese, Kamba branding)
    ///     "HTML_ALXB_CP"      – Alexandria Bank corporate prepaid
    ///     "HTML_FBN_CORP"     – First Bank Nigeria corporate
    ///     "HTML_SBN_SIG"      – Stanbic Bank Nigeria with digital signature
    ///     "RAWDATA_AIBK"      – AIBK raw DAT + CSV delivery files
    ///     "RAWDATA_ALXB_CORP" – ALXB corporate DAT files (VISA excluded)
    ///     "RAWDATA_AAIB"      – AAIB pipe-delimited with OverDueDays
    ///     "RAWDATA_BRKA"      – BRKA pipe-delimited with reward fields
    ///     "TEXTLABEL_FCMB"    – FCMB fixed-width text with page flags (F0/F1/F2/F3)
    ///     "TEXT_EDBE"         – EDBE plain text with control characters
    ///     "XML_IDBE"          – IDBE VIP XML + ZIP + MD5
    ///     "PDF_QNB"           – QNB Crystal Reports PDF
    ///
    /// The complete bank registry is maintained in
    /// Infrastructure.Configuration.DependencyInjection.
    /// </summary>
    public interface IStatementFormatterService
    {
        /// <summary>
        /// Unique registry key for this formatter.
        /// The factory resolves the formatter by matching this key to the
        /// FormatterKey on <see cref="GenerateStatementCommand"/>.
        /// </summary>
        string FormatterKey { get; }

        /// <summary>
        /// Generates all output files for a single bank/product statement run.
        ///
        /// The formatter receives the fully-loaded <see cref="StatementDataContext"/>
        /// (master DataSet plus optional reward, installment, email and identity
        /// supplementary DataSets) so it can access all Oracle data without
        /// re-querying the database.
        ///
        /// Returns the absolute paths of every file written, including the
        /// email-tracking text files ({prefix}_Emails.txt /
        /// {prefix}_WithoutEmails.txt) for HTML runs.
        ///
        /// All page-level logic (MaxDetailInPage = 20, MaxDetailInLastPage = 27),
        /// page-flag markers (F 0 / F 1 / F 2 / F 3), and control characters
        /// (\u000C form-feed, \u000D carriage-return) are handled internally.
        /// </summary>
        IEnumerable<string> Format(
            StatementDataContext context,
            string               outputDirectory,
            GenerateStatementCommand command);
    }
}

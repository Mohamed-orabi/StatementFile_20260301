namespace StatementFile.Domain.Interfaces.Services
{
    /// <summary>
    /// Validates that all pages in a generated text statement file contain
    /// exactly the expected number of lines per page.
    ///
    /// Legacy: clsValidatePageSize.ValidatePageSize(string, int, string)
    ///
    /// Used after generating fixed-width text statement files (TextLabel, EDBE)
    /// to verify the file is correctly paginated before FTP delivery.
    /// </summary>
    public interface IPageSizeValidationService
    {
        /// <summary>
        /// Reads the input file and counts pages separated by <paramref name="endOfPageCharacter"/>.
        /// For each page that does not contain exactly <paramref name="expectedPageSize"/> lines,
        /// the line number is written to a sibling error file named *_Err2.{ext}.
        /// </summary>
        /// <param name="filePath">Full path to the text statement file.</param>
        /// <param name="expectedPageSize">Expected number of lines per page.</param>
        /// <param name="endOfPageCharacter">The form-feed or end-of-page marker (e.g., "\f").</param>
        /// <returns>Number of pages with incorrect line counts.</returns>
        int Validate(string filePath, int expectedPageSize, string endOfPageCharacter);

        /// <summary>Gets the human-readable result message from the last call to Validate.</summary>
        string LastMessage { get; }
    }
}

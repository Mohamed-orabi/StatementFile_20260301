using System;
using System.IO;
using System.Text;
using StatementFile.Domain.Interfaces.Services;

namespace StatementFile.Infrastructure.Services
{
    /// <summary>
    /// Implementation of <see cref="IPageSizeValidationService"/>.
    ///
    /// Replicates clsValidatePageSize.ValidatePageSize() from the Common/ folder:
    ///  - Reads the input file line by line
    ///  - Counts lines within each page (bounded by the end-of-page character)
    ///  - Records incorrect page sizes in a sibling *_Err2.{ext} file
    ///  - Deletes the error file if no errors were found
    /// </summary>
    public sealed class PageSizeValidationService : IPageSizeValidationService
    {
        private string _lastMessage = string.Empty;

        public string LastMessage => _lastMessage;

        public int Validate(string filePath, int expectedPageSize, string endOfPageCharacter)
        {
            int numOfLine = 0, numOfErr = 0, curPageRow = 0;
            string errorFilePath = BuildErrorFilePath(filePath);

            try
            {
                using (var reader = new StreamReader(
                           new FileStream(filePath, FileMode.Open), Encoding.ASCII))
                using (var writer = new StreamWriter(
                           new FileStream(errorFilePath, FileMode.Create), Encoding.Default))
                {
                    // Skip first line (matches legacy behaviour)
                    reader.ReadLine();

                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        numOfLine++;
                        curPageRow++;

                        bool isEndOfPage = line == endOfPageCharacter
                                        || line == endOfPageCharacter + "----------------- END OF STATEMENT -----------------";

                        if (isEndOfPage)
                        {
                            if (curPageRow != expectedPageSize)
                            {
                                writer.WriteLine(numOfLine.ToString());
                                numOfErr++;
                            }
                            curPageRow = 0;
                        }
                    }

                    writer.Flush();
                }
            }
            catch (Exception ex)
            {
                _lastMessage = $"Validation error: {ex.Message}";
                return -1;
            }

            if (numOfErr == 0)
            {
                _lastMessage = "File Contain No Errors";
                if (File.Exists(errorFilePath))
                    File.Delete(errorFilePath);
            }
            else
            {
                _lastMessage = $"File Contain Errors logged to file {errorFilePath}";
            }

            return numOfErr;
        }

        private static string BuildErrorFilePath(string filePath)
        {
            string dir  = Path.GetDirectoryName(filePath) ?? string.Empty;
            string name = Path.GetFileNameWithoutExtension(filePath);
            string ext  = Path.GetExtension(filePath);
            return Path.Combine(dir, $"{name}_Err2{ext}");
        }
    }
}

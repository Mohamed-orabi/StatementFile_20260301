namespace StatementFile.Domain.Interfaces.Services
{
    /// <summary>
    /// Updates TSTATEMENTMASTERTABLE rows with zipcode and barcode values
    /// read from a pipe-delimited file.
    ///
    /// Legacy: clsBasUpdateStat.UpdateStat(string pStrFileNameIn)
    ///
    /// File format (one record per line):
    ///   {clientid}|{zipcode}|{barcode}
    ///
    /// Executed as Oracle PL/SQL BEGIN...END blocks (max 1000 statements per batch).
    /// </summary>
    public interface IStatUpdateService
    {
        /// <summary>
        /// Reads the pipe-delimited file and updates customerzipcode and barcode
        /// on TSTATEMENTMASTERTABLE rows matching branch = 4 and clientid.
        /// </summary>
        /// <param name="filePath">Full path to the pipe-delimited input file.</param>
        void UpdateFromFile(string filePath);
    }
}

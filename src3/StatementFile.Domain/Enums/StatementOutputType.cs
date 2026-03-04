namespace StatementFile.Domain.Enums
{
    /// <summary>
    /// The output format that the statement formatter produces.
    /// Mirrors the hardcoded formatter branch logic in frmStatementFileExtn.runStatement().
    /// </summary>
    public enum StatementOutputType
    {
        Html    = 1,
        Pdf     = 2,
        Text    = 3,
        RawData = 4,
        Xml     = 5,
        Email   = 6
    }
}

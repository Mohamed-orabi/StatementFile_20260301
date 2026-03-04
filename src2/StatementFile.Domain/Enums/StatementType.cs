namespace StatementFile.Domain.Enums
{
    /// <summary>
    /// Defines the high-level type of a statement batch run.
    /// N = Normal individual card statement.
    /// C = Corporate/company-level statement.
    /// M = Merchant XML-based statement.
    /// </summary>
    public enum StatementType
    {
        Normal    = 0,
        Corporate = 1,
        Merchant  = 2,
    }
}

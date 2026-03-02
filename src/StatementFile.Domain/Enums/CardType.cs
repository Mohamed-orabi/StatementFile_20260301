namespace StatementFile.Domain.Enums
{
    /// <summary>
    /// Classifies the card product family.
    /// Drives which master/detail table suffix (CR/DB/CP) is used.
    /// </summary>
    public enum CardType
    {
        Credit  = 0,   // CR tables
        Debit   = 1,   // DB tables
        Prepaid = 2,   // CP tables
        Reward  = 3,   // Credit-class with reward program contract type
    }
}

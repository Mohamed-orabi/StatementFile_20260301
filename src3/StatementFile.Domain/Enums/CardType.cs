namespace StatementFile.Domain.Enums
{
    /// <summary>
    /// Card portfolio type.
    /// CR = Credit, DB = Debit, CP = Corporate / Prepaid.
    /// Determines which Oracle master/detail tables are used:
    ///   CR → TSTATEMENTMASTERCR / TSTATEMENTDETAILCR
    ///   DB → TSTATEMENTMASTERDB / TSTATEMENTDETAILDB
    ///   CP → TSTATEMENTMASTERCP / TSTATEMENTDETAILCP
    /// </summary>
    public enum CardType
    {
        Credit    = 1,  // CR
        Debit     = 2,  // DB
        Corporate = 3   // CP (also covers Prepaid)
    }
}

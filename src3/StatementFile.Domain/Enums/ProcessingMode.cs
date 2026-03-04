using System;

namespace StatementFile.Domain.Enums
{
    /// <summary>
    /// Flags that control optional data-maintenance steps executed before statement generation.
    /// Derived from the processing options visible in frmStatementFile and applied per-product.
    /// </summary>
    [Flags]
    public enum ProcessingMode : long
    {
        None                = 0,
        Installment         = 1 << 0,   // exclude installment data from statement
        FixArabicAddress    = 1 << 1,   // fix Arabic address encoding
        FixAddress          = 1 << 2,   // normalise English address fields
        FixArabicAddressLang= 1 << 3,   // apply language-level Arabic address fix
        DeleteOnHold        = 1 << 4,   // delete on-hold statement records before run
        Reward              = 1 << 5,   // reward-points processing run
        MergeMarkUpFees     = 1 << 6    // merge mark-up fees into transactions
    }
}

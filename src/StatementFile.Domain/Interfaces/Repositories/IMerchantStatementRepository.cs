using System.Collections.Generic;
using StatementFile.Domain.Entities;

namespace StatementFile.Domain.Interfaces.Repositories
{
    /// <summary>
    /// Contract for persisting merchant statement data into the MS-Access staging database.
    /// </summary>
    public interface IMerchantStatementRepository
    {
        /// <summary>
        /// Inserts all master rows and their child operation rows into the .mdb template.
        /// </summary>
        void SaveMerchantStatements(IEnumerable<MerchantStatement> statements, string mdbFilePath);

        /// <summary>
        /// Runs the post-insert fixup queries (ExternalAccount, TD = OD).
        /// </summary>
        void ApplyPostInsertFixups(string mdbFilePath);

        /// <summary>
        /// Loads statement and operation data back from the .mdb after fixup for reporting.
        /// </summary>
        IEnumerable<MerchantStatement> LoadFromMdb(string mdbFilePath);
    }
}

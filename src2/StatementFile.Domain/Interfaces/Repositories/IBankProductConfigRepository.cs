using System.Collections.Generic;
using StatementFile.Domain.Entities;

namespace StatementFile.Domain.Interfaces.Repositories
{
    /// <summary>
    /// Repository for persisted bank/product configuration records.
    /// Data is stored in Oracle table STAT_BANK_PRODUCT_CONFIG.
    /// </summary>
    public interface IBankProductConfigRepository
    {
        /// <summary>Returns all configuration records regardless of active flag.</summary>
        IReadOnlyList<BankProductConfig> GetAll();

        /// <summary>Returns only configurations with IS_ACTIVE = 1, ordered by SORT_ORDER.</summary>
        IReadOnlyList<BankProductConfig> GetActive();

        /// <summary>Returns the configuration with the given primary key, or null.</summary>
        BankProductConfig GetById(int id);

        /// <summary>
        /// Inserts a new configuration row and returns the newly assigned ID.
        /// </summary>
        int Insert(BankProductConfig config);

        /// <summary>Updates all mutable fields of an existing configuration row.</summary>
        void Update(BankProductConfig config);

        /// <summary>Deletes the configuration row with the given primary key.</summary>
        void Delete(int id);
    }
}

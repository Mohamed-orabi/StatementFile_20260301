using System.Collections.Generic;
using StatementFile.Domain.Entities;

namespace StatementFile.Domain.Interfaces
{
    /// <summary>Repository contract for <see cref="BankProductConfig"/> persistence.</summary>
    public interface IBankProductConfigRepository
    {
        IReadOnlyList<BankProductConfig> GetAll();
        IReadOnlyList<BankProductConfig> GetActive();
        BankProductConfig GetById(int id);
        int Add(BankProductConfig config);
        void Update(BankProductConfig config);
        void Delete(int id);
    }
}

using System.Collections.Generic;
using StatementFile.Domain.Entities;

namespace StatementFile.Domain.Interfaces.Repositories
{
    /// <summary>
    /// Contract for querying bank / branch reference data from Oracle.
    /// </summary>
    public interface IBankRepository
    {
        Bank GetByBranchCode(int branchCode);
        IEnumerable<Bank> GetAll();
    }
}

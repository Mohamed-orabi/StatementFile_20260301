using System.Collections.Generic;
using StatementFile.Domain.Entities;

namespace StatementFile.Domain.Interfaces
{
    public interface IStatementRunRepository
    {
        int  Add(StatementRun run);
        void Update(StatementRun run);
        IReadOnlyList<StatementRun> GetByConfigId(int configId, int maxRows = 50);
    }
}

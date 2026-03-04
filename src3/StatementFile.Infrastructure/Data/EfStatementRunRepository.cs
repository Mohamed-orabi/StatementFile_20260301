using System;
using System.Collections.Generic;
using System.Linq;
using StatementFile.Domain.Entities;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Infrastructure.Data
{
    public sealed class EfStatementRunRepository : IStatementRunRepository
    {
        private readonly StatementFileDbContext _db;

        public EfStatementRunRepository(StatementFileDbContext db) =>
            _db = db ?? throw new ArgumentNullException(nameof(db));

        public int Add(StatementRun run)
        {
            _db.StatementRuns.Add(run);
            _db.SaveChanges();
            return run.Id;
        }

        public void Update(StatementRun run)
        {
            _db.StatementRuns.Update(run);
            _db.SaveChanges();
        }

        public IReadOnlyList<StatementRun> GetByConfigId(int configId, int maxRows = 50) =>
            _db.StatementRuns
               .Where(x => x.ConfigId == configId)
               .OrderByDescending(x => x.StartedAt)
               .Take(maxRows)
               .ToList();
    }
}

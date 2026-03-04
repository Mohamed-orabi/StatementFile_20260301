using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.EntityFrameworkCore;
using StatementFile.Domain.Entities;
using StatementFile.Domain.Interfaces.Repositories;

namespace StatementFile.Infrastructure.Data.Repositories
{
    /// <summary>
    /// EF Core implementation of <see cref="IBankProductConfigRepository"/>.
    ///
    /// Each public method creates a short-lived <see cref="AppDbContext"/> from the
    /// supplied options so that no long-lived context (with stale change-tracker state)
    /// is shared across calls.
    /// </summary>
    public sealed class BankProductConfigRepository : IBankProductConfigRepository
    {
        private readonly DbContextOptions<AppDbContext> _dbOptions;

        public BankProductConfigRepository(DbContextOptions<AppDbContext> dbOptions)
        {
            _dbOptions = dbOptions ?? throw new ArgumentNullException(nameof(dbOptions));
        }

        public IReadOnlyList<BankProductConfig> GetAll()
        {
            using var ctx = new AppDbContext(_dbOptions);
            return ctx.BankProductConfigs
                      .OrderBy(c => c.SortOrder)
                      .ThenBy(c => c.BranchCode)
                      .ToList();
        }

        public IReadOnlyList<BankProductConfig> GetActive()
        {
            using var ctx = new AppDbContext(_dbOptions);
            return ctx.BankProductConfigs
                      .Where(c => c.IsActive)
                      .OrderBy(c => c.SortOrder)
                      .ThenBy(c => c.BranchCode)
                      .ToList();
        }

        public BankProductConfig GetById(int id)
        {
            using var ctx = new AppDbContext(_dbOptions);
            return ctx.BankProductConfigs.Find(id);
        }

        public int Insert(BankProductConfig config)
        {
            using var ctx = new AppDbContext(_dbOptions);
            ctx.BankProductConfigs.Add(config);
            ctx.SaveChanges();
            return config.Id;
        }

        public void Update(BankProductConfig config)
        {
            using var ctx = new AppDbContext(_dbOptions);
            ctx.BankProductConfigs.Update(config);
            ctx.SaveChanges();
        }

        public void Delete(int id)
        {
            using var ctx = new AppDbContext(_dbOptions);
            var entity = ctx.BankProductConfigs.Find(id);
            if (entity != null)
            {
                ctx.BankProductConfigs.Remove(entity);
                ctx.SaveChanges();
            }
        }
    }
}

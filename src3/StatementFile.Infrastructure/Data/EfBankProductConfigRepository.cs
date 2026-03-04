using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.EntityFrameworkCore;
using StatementFile.Domain.Entities;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Infrastructure.Data
{
    /// <summary>
    /// EF Core implementation of <see cref="IBankProductConfigRepository"/>.
    /// Replaces the raw-ADO Oracle version.
    /// </summary>
    public sealed class EfBankProductConfigRepository : IBankProductConfigRepository
    {
        private readonly StatementFileDbContext _db;

        public EfBankProductConfigRepository(StatementFileDbContext db) =>
            _db = db ?? throw new ArgumentNullException(nameof(db));

        public IReadOnlyList<BankProductConfig> GetAll() =>
            _db.BankProductConfigs
               .OrderBy(x => x.BankName)
               .ThenBy(x => x.StatementTypeSuffix)
               .ToList();

        public IReadOnlyList<BankProductConfig> GetActive() =>
            _db.BankProductConfigs
               .Where(x => x.IsActive)
               .OrderBy(x => x.BankName)
               .ThenBy(x => x.StatementTypeSuffix)
               .ToList();

        public BankProductConfig GetById(int id) =>
            _db.BankProductConfigs.FirstOrDefault(x => x.Id == id);

        public int Add(BankProductConfig config)
        {
            _db.BankProductConfigs.Add(config);
            _db.SaveChanges();
            return config.Id;
        }

        public void Update(BankProductConfig config)
        {
            _db.BankProductConfigs.Update(config);
            _db.SaveChanges();
        }

        public void Delete(int id)
        {
            var entity = _db.BankProductConfigs.Find(id);
            if (entity != null)
            {
                _db.BankProductConfigs.Remove(entity);
                _db.SaveChanges();
            }
        }
    }
}

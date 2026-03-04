using Microsoft.EntityFrameworkCore;
using StatementFile.Domain.Entities;

namespace StatementFile.Infrastructure.Data
{
    public sealed class StatementFileDbContext : DbContext
    {
        public StatementFileDbContext(DbContextOptions<StatementFileDbContext> options)
            : base(options) { }

        public DbSet<BankProductConfig> BankProductConfigs { get; set; }
        public DbSet<StatementRun>      StatementRuns      { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.ApplyConfiguration(new BankProductConfigConfiguration());
            modelBuilder.ApplyConfiguration(new StatementRunConfiguration());
        }
    }
}

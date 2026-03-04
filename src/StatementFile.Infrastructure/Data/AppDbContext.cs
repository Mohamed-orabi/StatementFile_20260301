using Microsoft.EntityFrameworkCore;
using StatementFile.Domain.Entities;
using StatementFile.Infrastructure.Data.EntityConfigurations;

namespace StatementFile.Infrastructure.Data
{
    /// <summary>
    /// EF Core DbContext for the StatementFile infrastructure.
    /// Currently manages only <see cref="BankProductConfig"/>; all other data
    /// access uses raw SqlClient (DataAdapter/DataSet) to preserve the existing
    /// formatter pipeline.
    /// </summary>
    public sealed class AppDbContext : DbContext
    {
        public DbSet<BankProductConfig> BankProductConfigs { get; set; }

        public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.ApplyConfiguration(new BankProductConfigConfiguration());
        }
    }
}

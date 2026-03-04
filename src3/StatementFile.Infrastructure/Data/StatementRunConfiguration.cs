using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using StatementFile.Domain.Entities;

namespace StatementFile.Infrastructure.Data
{
    internal sealed class StatementRunConfiguration : IEntityTypeConfiguration<StatementRun>
    {
        public void Configure(EntityTypeBuilder<StatementRun> b)
        {
            b.ToTable("STAT_STATEMENT_RUN");

            b.HasKey(x => x.Id);
            b.Property(x => x.Id)             .HasColumnName("ID").UseIdentityColumn();
            b.Property(x => x.ConfigId)       .HasColumnName("CONFIG_ID").IsRequired();
            b.Property(x => x.StatementDate)  .HasColumnName("STATEMENT_DATE").IsRequired();
            b.Property(x => x.StartedAt)      .HasColumnName("STARTED_AT");
            b.Property(x => x.FinishedAt)     .HasColumnName("FINISHED_AT");
            b.Property(x => x.IsSuccess)      .HasColumnName("IS_SUCCESS");
            b.Property(x => x.ErrorMessage)   .HasColumnName("ERROR_MESSAGE").HasMaxLength(4000);
            b.Property(x => x.FilesGenerated) .HasColumnName("FILES_GENERATED");
            b.Property(x => x.EmailsSent)     .HasColumnName("EMAILS_SENT");
            b.Property(x => x.StatementsCount).HasColumnName("STATEMENTS_COUNT");
            b.Property(x => x.OutputDirectory).HasColumnName("OUTPUT_DIRECTORY").HasMaxLength(1000);

            b.HasOne<StatementFile.Domain.Entities.BankProductConfig>()
             .WithMany()
             .HasForeignKey(x => x.ConfigId)
             .OnDelete(DeleteBehavior.Restrict);
        }
    }
}

using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using StatementFile.Domain.Entities;
using StatementFile.Domain.Enums;

namespace StatementFile.Infrastructure.Data.EntityConfigurations
{
    /// <summary>
    /// Fluent API mapping for <see cref="BankProductConfig"/> → STAT_BANK_PRODUCT_CONFIG table.
    ///
    /// EF Core materialises the entity via the private parameterless constructor and
    /// sets all private-setter properties through reflection — no special constructor
    /// or field-backing configuration is required.
    /// </summary>
    internal sealed class BankProductConfigConfiguration
        : IEntityTypeConfiguration<BankProductConfig>
    {
        public void Configure(EntityTypeBuilder<BankProductConfig> builder)
        {
            builder.ToTable("STAT_BANK_PRODUCT_CONFIG");

            // ── Primary key (IDENTITY) ─────────────────────────────────────────────
            builder.HasKey(e => e.Id);
            builder.Property(e => e.Id)
                   .HasColumnName("ID")
                   .ValueGeneratedOnAdd();

            // ── Identity ──────────────────────────────────────────────────────────
            builder.Property(e => e.BranchCode)
                   .HasColumnName("BRANCH_CODE")
                   .IsRequired();

            builder.Property(e => e.BankName)
                   .HasColumnName("BANK_NAME")
                   .HasMaxLength(50)
                   .IsRequired();

            builder.Property(e => e.BankFullName)
                   .HasColumnName("BANK_FULL_NAME")
                   .HasMaxLength(200)
                   .IsRequired();

            builder.Property(e => e.CardProduct)
                   .HasColumnName("CARD_PRODUCT")
                   .HasMaxLength(50);

            // ── Output format ──────────────────────────────────────────────────────
            builder.Property(e => e.OutputType)
                   .HasColumnName("OUTPUT_TYPE")
                   .HasConversion<int>()
                   .IsRequired();

            builder.Property(e => e.FormatterKey)
                   .HasColumnName("FORMATTER_KEY")
                   .HasMaxLength(100)
                   .IsRequired();

            // ── Processing flags (flags enum stored as bigint) ─────────────────────
            builder.Property(e => e.ProcessingModes)
                   .HasColumnName("PROCESSING_MODES")
                   .HasConversion<long>()
                   .HasColumnType("bigint")
                   .IsRequired();

            // ── Card classification ────────────────────────────────────────────────
            builder.Property(e => e.CardType)
                   .HasColumnName("CARD_TYPE")
                   .HasConversion<int>()
                   .IsRequired();

            builder.Property(e => e.StatementType)
                   .HasColumnName("STATEMENT_TYPE")
                   .HasConversion<int>()
                   .IsRequired();

            builder.Property(e => e.StatementTypeSuffix)
                   .HasColumnName("STATEMENT_TYPE_SUFFIX")
                   .HasMaxLength(10);

            // ── Optional conditions ────────────────────────────────────────────────
            builder.Property(e => e.RewardContractCondition)
                   .HasColumnName("REWARD_CONTRACT_COND")
                   .HasMaxLength(200);

            builder.Property(e => e.InstallmentContractCondition)
                   .HasColumnName("INSTALLMENT_CONTRACT_COND")
                   .HasMaxLength(200);

            // ── Flags (bool → bit) ─────────────────────────────────────────────────
            builder.Property(e => e.AppendMode)
                   .HasColumnName("APPEND_MODE")
                   .IsRequired();

            builder.Property(e => e.UseCorporateAccountNumber)
                   .HasColumnName("USE_CORP_ACCT_NO")
                   .IsRequired();

            // ── PDF ────────────────────────────────────────────────────────────────
            builder.Property(e => e.ReportTemplateName)
                   .HasColumnName("REPORT_TEMPLATE_NAME")
                   .HasMaxLength(200);

            builder.Property(e => e.PdfPasswordScheme)
                   .HasColumnName("PDF_PASSWORD_SCHEME")
                   .HasMaxLength(50);

            // ── Email / branding ───────────────────────────────────────────────────
            builder.Property(e => e.EmailFromAddress)
                   .HasColumnName("EMAIL_FROM_ADDRESS")
                   .HasMaxLength(200);

            builder.Property(e => e.BankWebsiteUrl)
                   .HasColumnName("BANK_WEBSITE_URL")
                   .HasMaxLength(500);

            builder.Property(e => e.FacebookUrl)
                   .HasColumnName("FACEBOOK_URL")
                   .HasMaxLength(500);

            builder.Property(e => e.LinkedInUrl)
                   .HasColumnName("LINKEDIN_URL")
                   .HasMaxLength(500);

            builder.Property(e => e.YouTubeUrl)
                   .HasColumnName("YOUTUBE_URL")
                   .HasMaxLength(500);

            // ── Pagination ─────────────────────────────────────────────────────────
            builder.Property(e => e.MaxTransactionsPerPage)
                   .HasColumnName("MAX_TRANS_PER_PAGE")
                   .IsRequired();

            builder.Property(e => e.MaxTransactionsLastPage)
                   .HasColumnName("MAX_TRANS_LAST_PAGE")
                   .IsRequired();

            // ── Email delivery ─────────────────────────────────────────────────────
            builder.Property(e => e.EmailWaitPeriodMs)
                   .HasColumnName("EMAIL_WAIT_MS")
                   .IsRequired();

            // ── Raw data ───────────────────────────────────────────────────────────
            builder.Property(e => e.FieldSeparator)
                   .HasColumnName("FIELD_SEPARATOR")
                   .HasMaxLength(5);

            // ── Scheduling ─────────────────────────────────────────────────────────
            builder.Property(e => e.ScheduledDay)
                   .HasColumnName("SCHEDULED_DAY")
                   .IsRequired();

            // ── Display ────────────────────────────────────────────────────────────
            builder.Property(e => e.DisplayName)
                   .HasColumnName("DISPLAY_NAME")
                   .HasMaxLength(500);

            builder.Property(e => e.SortOrder)
                   .HasColumnName("SORT_ORDER")
                   .IsRequired();

            builder.Property(e => e.IsActive)
                   .HasColumnName("IS_ACTIVE")
                   .IsRequired();

            // ── Audit ──────────────────────────────────────────────────────────────
            builder.Property(e => e.CreatedAt)
                   .HasColumnName("CREATED_AT")
                   .HasColumnType("datetime2")
                   .IsRequired();

            builder.Property(e => e.ModifiedAt)
                   .HasColumnName("MODIFIED_AT")
                   .HasColumnType("datetime2")
                   .IsRequired();
        }
    }
}

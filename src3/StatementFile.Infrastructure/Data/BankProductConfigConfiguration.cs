using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using StatementFile.Domain.Entities;
using StatementFile.Domain.Enums;

namespace StatementFile.Infrastructure.Data
{
    internal sealed class BankProductConfigConfiguration : IEntityTypeConfiguration<BankProductConfig>
    {
        public void Configure(EntityTypeBuilder<BankProductConfig> b)
        {
            b.ToTable("STAT_BANK_PRODUCT_CONFIG");

            b.HasKey(x => x.Id);
            b.Property(x => x.Id)
             .HasColumnName("ID")
             .UseIdentityColumn();

            // EF Core can write private-setter properties via reflection.
            b.Property(x => x.IsActive)             .HasColumnName("IS_ACTIVE");
            b.Property(x => x.BankName)             .HasColumnName("BANK_NAME")              .IsRequired().HasMaxLength(50);
            b.Property(x => x.BankFullName)         .HasColumnName("BANK_FULL_NAME")          .HasMaxLength(200);
            b.Property(x => x.BankCode)             .HasColumnName("BANK_CODE")               .HasMaxLength(20);
            b.Property(x => x.BranchCode)           .HasColumnName("BRANCH_CODE");
            b.Property(x => x.StatementTypeSuffix)  .HasColumnName("STATEMENT_TYPE_SUFFIX")   .HasMaxLength(20);

            // Enum stored as int
            b.Property(x => x.CardType)
             .HasColumnName("CARD_TYPE")
             .HasConversion<int>();

            b.Property(x => x.CardProduct)          .HasColumnName("CARD_PRODUCT")            .HasMaxLength(100);

            b.Property(x => x.OutputType)
             .HasColumnName("OUTPUT_TYPE")
             .HasConversion<int>();

            b.Property(x => x.FormatterKey)         .HasColumnName("FORMATTER_KEY")           .IsRequired().HasMaxLength(100);
            b.Property(x => x.BankWebLink)          .HasColumnName("BANK_WEB_LINK")           .HasMaxLength(500);
            b.Property(x => x.BankLogo)             .HasColumnName("BANK_LOGO")               .HasMaxLength(500);
            b.Property(x => x.BackgroundImage)      .HasColumnName("BACKGROUND_IMAGE")        .HasMaxLength(500);
            b.Property(x => x.MidBannerImage)       .HasColumnName("MID_BANNER_IMAGE")        .HasMaxLength(500);
            b.Property(x => x.BottomBannerImage)    .HasColumnName("BOTTOM_BANNER_IMAGE")     .HasMaxLength(500);
            b.Property(x => x.EmailFromAddress)     .HasColumnName("EMAIL_FROM_ADDRESS")      .HasMaxLength(200);
            b.Property(x => x.EmailFromName)        .HasColumnName("EMAIL_FROM_NAME")         .HasMaxLength(200);
            b.Property(x => x.WhereCondition)       .HasColumnName("WHERE_CONDITION")         .HasMaxLength(2000);
            b.Property(x => x.VipCondition)         .HasColumnName("VIP_CONDITION")           .HasMaxLength(2000);
            b.Property(x => x.RewardCondition)      .HasColumnName("REWARD_CONDITION")        .HasMaxLength(2000);
            b.Property(x => x.RewardContractCondition).HasColumnName("REWARD_CONTRACT_CONDITION").HasMaxLength(2000);
            b.Property(x => x.CurrencyFilter)       .HasColumnName("CURRENCY_FILTER")         .HasMaxLength(2000);
            b.Property(x => x.InstallmentCondition) .HasColumnName("INSTALLMENT_CONDITION")   .HasMaxLength(2000);
            b.Property(x => x.PaymentSystem)        .HasColumnName("PAYMENT_SYSTEM")          .HasMaxLength(50);

            // ProcessingMode [Flags] enum stored as long
            b.Property(x => x.ProcessingModes)
             .HasColumnName("PROCESSING_MODES")
             .HasConversion<long>();

            b.Property(x => x.IsRewardRun)          .HasColumnName("IS_REWARD_RUN");
            b.Property(x => x.IsSplitOutput)        .HasColumnName("IS_SPLIT_OUTPUT");
            b.Property(x => x.HasAttachment)        .HasColumnName("HAS_ATTACHMENT");
            b.Property(x => x.SaveDataset)          .HasColumnName("SAVE_DATASET");
            b.Property(x => x.ShowMessageBox)       .HasColumnName("SHOW_MESSAGE_BOX");
            b.Property(x => x.RunNullCardDelete)    .HasColumnName("RUN_NULL_CARD_DELETE");
            b.Property(x => x.RunCardBranchMatch)   .HasColumnName("RUN_CARD_BRANCH_MATCH");
            b.Property(x => x.ExcludeReward)        .HasColumnName("EXCLUDE_REWARD");
            b.Property(x => x.WaitPeriodSeconds)    .HasColumnName("WAIT_PERIOD_SECONDS");
            b.Property(x => x.CreatedAt)            .HasColumnName("CREATED_AT");
            b.Property(x => x.UpdatedAt)            .HasColumnName("UPDATED_AT");
        }
    }
}

using StatementFile.Domain.Enums;

namespace StatementFile.Application.DTOs
{
    /// <summary>
    /// Flat, JSON-serialisable projection of <see cref="StatementFile.Domain.Entities.BankProductConfig"/>.
    /// Used as the API contract so domain entities with private setters are never exposed directly.
    /// </summary>
    public sealed class BankConfigDto
    {
        public int    Id                        { get; set; }
        public bool   IsActive                  { get; set; }
        public string BankName                  { get; set; }
        public string BankFullName              { get; set; }
        public string BankCode                  { get; set; }
        public int    BranchCode                { get; set; }
        public string StatementTypeSuffix       { get; set; }
        public CardType CardType                { get; set; }
        public string CardProduct               { get; set; }
        public StatementOutputType OutputType   { get; set; }
        public string FormatterKey              { get; set; }
        public string BankWebLink               { get; set; }
        public string BankLogo                  { get; set; }
        public string BackgroundImage           { get; set; }
        public string MidBannerImage            { get; set; }
        public string BottomBannerImage         { get; set; }
        public string EmailFromAddress          { get; set; }
        public string EmailFromName             { get; set; }
        public string WhereCondition            { get; set; }
        public string VipCondition              { get; set; }
        public string RewardCondition           { get; set; }
        public string RewardContractCondition   { get; set; }
        public string CurrencyFilter            { get; set; }
        public string InstallmentCondition      { get; set; }
        public string PaymentSystem             { get; set; }
        public long   ProcessingModes           { get; set; }
        public bool   IsRewardRun               { get; set; }
        public bool   IsSplitOutput             { get; set; }
        public bool   HasAttachment             { get; set; }
        public bool   SaveDataset               { get; set; }
        public bool   ShowMessageBox            { get; set; }
        public bool   RunNullCardDelete         { get; set; }
        public bool   RunCardBranchMatch        { get; set; }
        public bool   ExcludeReward             { get; set; }
        public int    WaitPeriodSeconds         { get; set; }
    }
}

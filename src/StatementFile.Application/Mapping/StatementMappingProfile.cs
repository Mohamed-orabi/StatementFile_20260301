using System.Data;
using StatementFile.Application.DTOs;
using StatementFile.Domain.Entities;

namespace StatementFile.Application.Mapping
{
    /// <summary>
    /// Hand-written mapping helpers (no AutoMapper dependency in Domain/Application).
    /// Converts domain entities → DTOs and DataRow → domain entities.
    /// </summary>
    public static class StatementMappingProfile
    {
        // ── Entity → DTO ──────────────────────────────────────────────────────────

        public static StatementDto ToDto(this Statement s)
        {
            return new StatementDto
            {
                Branch              = s.Branch,
                StatementNo         = s.StatementNo,
                StatementNumber     = s.StatementNumber,
                CardNo              = s.CardNo,
                Mbr                 = s.Mbr,
                CardProduct         = s.CardProduct,
                CardPrimary         = s.CardPrimary,
                PrimaryCardNo       = s.PrimaryCardNo,
                AccountNo           = s.AccountNo,
                ExternalNo          = s.ExternalNo,
                AccountType         = s.AccountType,
                AccountStatus       = s.AccountStatus,
                AccountCurrency     = s.AccountCurrency,
                StatementDateFrom   = s.StatementDateFrom,
                StatementDateTo     = s.StatementDateTo,
                StatementDueDate    = s.StatementDueDate,
                CardState           = s.CardState,
                CardStatus          = s.CardStatus,
                CardExpiryDate      = s.CardExpiryDate,
                CardVip             = s.CardVip,
                CardPaymentMethod   = s.CardPaymentMethod,
                CustomerName        = s.CustomerName,
                CustomerTitle       = s.CustomerTitle,
                CustomerAddress1    = s.CustomerAddress1,
                CustomerAddress2    = s.CustomerAddress2,
                CustomerAddress3    = s.CustomerAddress3,
                CustomerCity        = s.CustomerCity,
                CustomerCountry     = s.CustomerCountry,
                AccountLim          = s.AccountLim,
                AccountAvailableLim = s.AccountAvailableLim,
                CardLimit           = s.CardLimit,
                CardAvailableLimit  = s.CardAvailableLimit,
                TotalOverdueAmount  = s.TotalOverdueAmount,
                TotalDueAmount      = s.TotalDueAmount,
                MinDueAmount        = s.MinDueAmount,
                OpeningBalance      = s.OpeningBalance,
                ClosingBalance      = s.ClosingBalance,
                ContractType        = s.ContractType,
                ContractNo          = s.ContractNo,
                EarnedBonus         = s.EarnedBonus,
                RedeemedBonus       = s.RedeemedBonus,
                ExpiredBonus        = s.ExpiredBonus,
                StatementMessageLine1 = s.StatementMessageLine1,
                StatementMessageLine2 = s.StatementMessageLine2,
                StatementMessageLine3 = s.StatementMessageLine3,
            };
        }

        public static StatementTransactionDto ToDto(this StatementTransaction t)
        {
            return new StatementTransactionDto
            {
                StatementNo         = t.StatementNo,
                CardNo              = t.CardNo,
                AccountNo           = t.AccountNo,
                TransDate           = t.TransDate,
                PostingDate         = t.PostingDate,
                TranDescription     = t.TranDescription,
                Merchant            = t.Merchant,
                OrigTranAmount      = t.OrigTranAmount,
                OrigTranCurrency    = t.OrigTranCurrency,
                BillTranAmount      = t.BillTranAmount,
                BillTranAmountSign  = t.BillTranAmountSign,
                DocNo               = t.DocNo,
                ApprovalCode        = t.ApprovalCode,
                PackageName         = t.PackageName,
                EntryNo             = t.EntryNo,
                InstallmentData     = t.InstallmentData,
            };
        }

        public static BankDto ToDto(this Bank b)
        {
            return new BankDto
            {
                BranchCode = b.BranchCode,
                BranchName = b.BranchName,
                BranchPart = b.BranchPart,
                Ident      = b.Ident,
            };
        }
    }
}

using System;
using StatementFile.Domain.Entities;

namespace StatementFile.Application.UseCases.BankConfiguration
{
    /// <summary>
    /// Read-only data transfer object returned by the bank-config API.
    /// Mirrors every property of <see cref="BankProductConfig"/> using
    /// public setters so System.Text.Json can deserialize it on the client.
    /// </summary>
    public sealed class BankProductConfigDto
    {
        // ── Primary key ───────────────────────────────────────────────────────
        public int Id { get; set; }

        // ── Identity ──────────────────────────────────────────────────────────
        public int    BranchCode   { get; set; }
        public string BankName     { get; set; }
        public string BankFullName { get; set; }
        public string CardProduct  { get; set; }

        // ── Output format ─────────────────────────────────────────────────────
        public int    OutputType   { get; set; }   // StatementOutputType enum value
        public string FormatterKey { get; set; }

        // ── Processing flags ──────────────────────────────────────────────────
        public long ProcessingModes { get; set; }  // ProcessingMode flags bitmask

        // ── Card classification ───────────────────────────────────────────────
        public int CardType      { get; set; }     // CardType enum value
        public int StatementType { get; set; }     // StatementType enum value

        // ── Statement suffix ──────────────────────────────────────────────────
        public string StatementTypeSuffix { get; set; }

        // ── Oracle conditions ─────────────────────────────────────────────────
        public string RewardContractCondition      { get; set; }
        public string InstallmentContractCondition { get; set; }

        // ── Flags ─────────────────────────────────────────────────────────────
        public bool AppendMode               { get; set; }
        public bool UseCorporateAccountNumber { get; set; }

        // ── PDF ───────────────────────────────────────────────────────────────
        public string ReportTemplateName { get; set; }
        public string PdfPasswordScheme  { get; set; }

        // ── Email / branding ──────────────────────────────────────────────────
        public string EmailFromAddress { get; set; }
        public string BankWebsiteUrl   { get; set; }
        public string FacebookUrl      { get; set; }
        public string LinkedInUrl      { get; set; }
        public string YouTubeUrl       { get; set; }

        // ── Pagination / timing ───────────────────────────────────────────────
        public int MaxTransactionsPerPage  { get; set; }
        public int MaxTransactionsLastPage { get; set; }
        public int EmailWaitPeriodMs       { get; set; }

        // ── Raw data ──────────────────────────────────────────────────────────
        public string FieldSeparator { get; set; }

        // ── Scheduling / display ──────────────────────────────────────────────
        public int    ScheduledDay { get; set; }
        public string DisplayName  { get; set; }
        public int    SortOrder    { get; set; }
        public bool   IsActive     { get; set; }

        // ── Audit ─────────────────────────────────────────────────────────────
        public DateTime CreatedAt  { get; set; }
        public DateTime ModifiedAt { get; set; }

        /// <summary>Maps a domain entity to this DTO.</summary>
        public static BankProductConfigDto From(BankProductConfig e) => new()
        {
            Id                           = e.Id,
            BranchCode                   = e.BranchCode,
            BankName                     = e.BankName,
            BankFullName                 = e.BankFullName,
            CardProduct                  = e.CardProduct,
            OutputType                   = (int)e.OutputType,
            FormatterKey                 = e.FormatterKey,
            ProcessingModes              = (long)e.ProcessingModes,
            CardType                     = (int)e.CardType,
            StatementType                = (int)e.StatementType,
            StatementTypeSuffix          = e.StatementTypeSuffix,
            RewardContractCondition      = e.RewardContractCondition,
            InstallmentContractCondition = e.InstallmentContractCondition,
            AppendMode                   = e.AppendMode,
            UseCorporateAccountNumber    = e.UseCorporateAccountNumber,
            ReportTemplateName           = e.ReportTemplateName,
            PdfPasswordScheme            = e.PdfPasswordScheme,
            EmailFromAddress             = e.EmailFromAddress,
            BankWebsiteUrl               = e.BankWebsiteUrl,
            FacebookUrl                  = e.FacebookUrl,
            LinkedInUrl                  = e.LinkedInUrl,
            YouTubeUrl                   = e.YouTubeUrl,
            MaxTransactionsPerPage       = e.MaxTransactionsPerPage,
            MaxTransactionsLastPage      = e.MaxTransactionsLastPage,
            EmailWaitPeriodMs            = e.EmailWaitPeriodMs,
            FieldSeparator               = e.FieldSeparator,
            ScheduledDay                 = e.ScheduledDay,
            DisplayName                  = e.DisplayName,
            SortOrder                    = e.SortOrder,
            IsActive                     = e.IsActive,
            CreatedAt                    = e.CreatedAt,
            ModifiedAt                   = e.ModifiedAt,
        };
    }
}

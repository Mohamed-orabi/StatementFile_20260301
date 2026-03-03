using System;
using System.Collections.Generic;
using System.IO;
using StatementFile.Application.UseCases.StatementGeneration;

namespace StatementFile.Infrastructure.Formatters.RawData
{
    // ── All raw-data (DAT/CSV/TXT) format adapters ──────────────────────────────

    /// <summary>
    /// AIBK raw data files:
    ///   STMT_HDR.DAT  – header records (comma-separated)
    ///   STMT_DTL.DAT  – detail/transaction records
    ///   STMT_Delivery.CSV – delivery manifest (account, address, phone, citizen ID)
    ///   .zip + .MD5 packaging
    ///
    /// Sort mode: FillStatementDataSet_SortCardPriority.
    /// Includes: installment processing, passport/identity data, product lookup.
    /// On-Hold rows excluded (HOLSTMT = 'Y').
    /// </summary>
    public sealed class RawDataAibkAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "RAWDATA_AIBK";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatRawDataAIBK().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    /// <summary>
    /// EGB raw data files (same structure as AIBK: HDR.DAT + DTL.DAT + Delivery.CSV).
    /// Applies card-number formatting via basText.formatCardNumber().
    /// </summary>
    public sealed class RawDataEgbAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "RAWDATA_EGB";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatRawDataEGB().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    /// <summary>
    /// AAIB raw data files:
    ///   _Main.txt  – pipe-delimited master/header records
    ///   _Trns.txt  – pipe-delimited transaction records
    ///   Includes OVERDUEDAYS column (special query variant).
    ///   Installment repayment parsing included.
    /// Jira: AAIB-9308, AAIB-12395.
    /// </summary>
    public sealed class RawDataAaibAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "RAWDATA_AAIB";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatRawData_AAIB().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    /// <summary>
    /// AIBK alternate format:
    ///   STM_Header_yyMM.txt  – header records
    ///   STM_Detail_yyMM.txt  – detail records
    ///   Uses "yyMM" date format in file names (not "yyyyMM").
    /// </summary>
    public sealed class RawDataAibkAltAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "RAWDATA_AIBK_ALT";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatRawData_AIBK().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    /// <summary>
    /// ALXB retail raw data (VISA cards excluded):
    ///   STMT_HDR.DAT + STMT_DTL.DAT + STMT_Delivery.CSV
    ///   Uses FillStatementDataSet_Exclude_VisaCards().
    ///   Installment docno stored as long to prevent int overflow.
    /// Jira: ALXB-5971.
    /// </summary>
    public sealed class RawDataAlxbAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "RAWDATA_ALXB";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatRawData_ALXB().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    /// <summary>
    /// ALXB corporate raw data (VISA cards excluded):
    ///   STMT_CORP_HDR.DAT + STMT_CORP_DTL.DAT + STMT_CORP_SMR.DAT
    ///   Pipe-separated. Includes company summary section.
    ///   Dynamic account name: cardaccountno vs accountno.
    /// Jira: ALXB-5971.
    /// </summary>
    public sealed class RawDataAlxbCorpAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "RAWDATA_ALXB_CORP";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatRawDataCorp_ALXB().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    /// <summary>
    /// BRKA raw data:
    ///   _Main.txt + _Trns.txt (pipe-separated)
    ///   Reward programme section: EarnedBonus, RedeemedBonus, ExpiredBonus,
    ///   BonusAdjustment, CreditContracts, OverDueDays, Card Expiry Date.
    /// </summary>
    public sealed class RawDataBrkaAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "RAWDATA_BRKA";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatRawData_BRKA().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    /// <summary>
    /// UNB raw data:
    ///   _Main.txt + _Trns.txt (pipe-separated)
    ///   No account-level loop (supplementary card removal).
    ///   Includes CardPrimary and PrimaryCardNo fields.
    /// </summary>
    public sealed class RawDataUnbAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "RAWDATA_UNB";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatRawData_UNB().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    /// <summary>
    /// VCBK raw data:
    ///   _Main.txt + _Trns.txt (pipe-separated)
    ///   Includes Email address and TotalDueAmount in header.
    ///   Detects MasterCard vs VISA via statement-type field.
    ///   Transaction date format: "dd/MMM/yyyy".
    /// </summary>
    public sealed class RawDataVcbkAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "RAWDATA_VCBK";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatRawData_VCBK().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }
}

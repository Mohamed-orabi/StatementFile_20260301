using System;
using System.Collections.Generic;
using System.IO;
using StatementFile.Application.UseCases.StatementGeneration;

namespace StatementFile.Infrastructure.Formatters.TextLabel
{
    // ── All text-label (fixed-width printer) format adapters ─────────────────────

    /// <summary>
    /// FCMB fixed-width text label for physical laser printer.
    ///
    /// Page flag logic:
    ///   F 0 = single-page statement
    ///   F 1 = first page of multi-page statement
    ///   F 2 = middle page
    ///   F 3 = last page
    ///
    /// Special sections:
    ///   - DAF (Deferred Annual Fee) calculation and display
    ///   - Reward programme table (RewardOpenBalance, EarnedBonus,
    ///     RedeemedBonus, ClosingBalance)
    ///
    /// Column headers: "Primary Card No | Credit Limit | Available Limit |
    ///   Minimum Due | Due Date | Over Due Amount"
    /// Transaction header: "T Date | P Date | Reference No | Description | Amount"
    ///
    /// Jira: FCMB-1790 (split Purchases and Cash columns).
    /// </summary>
    public sealed class TextLabelFcmbAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "TEXTLABEL_FCMB";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatTxtLbl_FCMB().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    /// <summary>
    /// Suez Canal Bank fixed-width text label.
    /// Same page-flag logic as FCMB.
    /// </summary>
    public sealed class TextLabelSuezAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "TEXTLABEL_SUEZ";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            // clsStatTxtLbl _Suez (note space in filename)
            new clsStatTxtLbl_Suez().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    /// <summary>FBN Debit text label.</summary>
    public sealed class TextLabelDbFbnAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "TEXTLABEL_FBN_DB";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatTxtLblDbFBN().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    /// <summary>AIB Debit text label.</summary>
    public sealed class TextLabelDbAibAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "TEXTLABEL_AIB_DB";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatTxtLblDb_AIB().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    /// <summary>BCA Debit text label.</summary>
    public sealed class TextLabelDbBcaAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "TEXTLABEL_BCA_DB";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatTxtLblDb_BCA().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    /// <summary>ICBG Debit text label.</summary>
    public sealed class TextLabelDbIcbgAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "TEXTLABEL_ICBG_DB";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatTxtLblDb_ICBG().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    /// <summary>EDBE text label.</summary>
    public sealed class TextLabelEdbeAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "TEXTLABEL_EDBE";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatTxtLblEDBE().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    /// <summary>FABG text label.</summary>
    public sealed class TextLabelFabgAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "TEXTLABEL_FABG";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatTxtLbl_FABG().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    // ── Plain-text (control-character) formatters ─────────────────────────────────

    /// <summary>
    /// EDBE plain-text statement with form-feed (\u000C) and carriage-return
    /// (\u000D) control characters for physical printing.
    /// End-of-statement marker: "--- END OF STATEMENT ---"
    /// isEmailStat flag controls email vs printed output path.
    /// </summary>
    public sealed class TextEdbeAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "TEXT_EDBE";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatTxt_EDBE().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }

    // ── XML formatter ─────────────────────────────────────────────────────────────

    /// <summary>
    /// IDBE VIP XML statement.
    /// Produces DataSet.WriteXml(XmlWriteMode.WriteSchema) output + ZIP + MD5.
    /// VIP filter: FillStatementDataSet(branchCode, "vip").
    /// Branch 16 only.
    /// </summary>
    public sealed class XmlIdbeAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "XML_IDBE";

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatXML_IDBE().Statement(
                fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate,
                cmd.StatementTypeSuffix, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName),
                startedAt);
        }
    }
}

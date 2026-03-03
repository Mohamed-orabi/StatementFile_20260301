using System;
using System.Data;
using System.IO;
using System.Text;
using StatementFile.Application.UseCases.StatementGeneration;
using StatementFile.Domain.Enums;

namespace StatementFile.Infrastructure.Formatters.TextLabel
{
    // ── Fixed-width text-label formatters (one class per bank) ───────────────────
    // All extend NativeTextLabelFormatterBase — no legacy class references.
    // ─────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// FCMB fixed-width text label for physical laser printer.
    ///
    /// Special features:
    ///   - DAF (Deferred Annual Fee) calculation and display
    ///   - Reward programme table: OpenBalance | EarnedBonus | RedeemedBonus | ClosingBalance
    ///   - Column header: Primary Card No | Credit Limit | Available Limit |
    ///                    Minimum Due | Due Date | Over Due Amount
    ///   - Jira FCMB-1790: separate Purchases and Cash columns
    /// </summary>
    public sealed class TextLabelFcmbFormatter : NativeTextLabelFormatterBase
    {
        public override string FormatterKey => "TEXTLABEL_FCMB";

        protected override void PrintHeader(GenerateStatementCommand cmd)
        {
            string custName  = AlignLeft(Str(MasterRow, MCustomerName), 50);
            string pageStr   = AlignRight($"Page {PageNo} of {TotalAccPages}", 25);
            Output.WriteLine(CurrentPageFlag + " " + custName + pageStr);

            string addr1     = AlignLeft(Str(MasterRow, MCustomerAddr1), 50);
            DateTime stmtDt  = ToDateTime(MasterRow, MStatementDateTo);
            Output.WriteLine("    " + addr1 + AlignRight(stmtDt.ToString("dd/MM/yyyy"), 25));

            string addr2     = AlignLeft(Str(MasterRow, MCustomerAddr2), 50);
            string cardFmt   = FormatCardNo(CurMainCard);
            Output.WriteLine("    " + addr2 + AlignRight(cardFmt, 25));

            // Column headers
            decimal limit    = ToDecimal(MasterRow, MAccountLim);
            decimal avail    = ToDecimal(MasterRow, MAvailableLim);
            decimal minDue   = ToDecimal(MasterRow, MMinDue);
            DateTime dueDate = ToDateTime(MasterRow, MStatementDueDate);
            decimal overdue  = ToDecimal(MasterRow, MTotalOverdue);

            Output.WriteLine(
                "    " +
                AlignLeft("Primary Card No", 20) +
                AlignRight($"{limit:N2}", 14) + "  " +
                AlignRight($"{avail:N2}", 14) + "  " +
                AlignRight($"{minDue:N2}", 14) + "  " +
                AlignRight(dueDate.ToString("dd/MM/yyyy"), 12) + "  " +
                AlignRight($"{overdue:N2}", 14));
        }

        protected override void PrintCardFooter(GenerateStatementCommand cmd, StatementDataContext ctx)
        {
            base.PrintCardFooter(cmd, ctx);

            // Reward section
            if (cmd.HasMode(ProcessingMode.Reward) && ctx.RewardDataSet?.Tables[0] != null)
            {
                string accountNo = CurAccountNo;
                var rr = ctx.RewardDataSet.Tables[0].Select($"ACCOUNTNO = '{EscSql(accountNo)}'");
                if (rr.Length > 0)
                {
                    decimal openBal   = ToDecimal(rr[0], "openingbonus");
                    decimal earned    = ToDecimal(rr[0], "earnedbonus");
                    decimal redeemed  = ToDecimal(rr[0], "redeemedbonus");
                    decimal closeBal  = ToDecimal(rr[0], "closingbonus");

                    Output.WriteLine(
                        AlignRight($"{openBal:N0}", 20) +
                        AlignRight($"{earned:N0}", 20) +
                        AlignRight($"{redeemed:N0}", 20) +
                        AlignRight($"{closeBal:N0}", 20));
                }
            }
        }

        private static string FormatCardNo(string cardNo)
        {
            if (cardNo?.Length == 16)
                return $"{cardNo[..4]}-{cardNo[4..8]}-{cardNo[8..12]}-{cardNo[12..]}";
            return cardNo ?? string.Empty;
        }
    }

    /// <summary>Suez Canal Bank fixed-width text label (same page-flag logic as FCMB).</summary>
    public sealed class TextLabelSuezFormatter : NativeTextLabelFormatterBase
    {
        public override string FormatterKey => "TEXTLABEL_SUEZ";
    }

    /// <summary>FBN Debit text label.</summary>
    public sealed class TextLabelDbFbnFormatter : NativeTextLabelFormatterBase
    {
        public override string FormatterKey => "TEXTLABEL_FBN_DB";

        protected override void PrintHeader(GenerateStatementCommand cmd)
        {
            string custName  = AlignLeft(Str(MasterRow, MCustomerName), 50);
            string pageStr   = AlignRight($"Page {PageNo} of {TotalAccPages}", 25);
            Output.WriteLine(CurrentPageFlag + " " + custName + pageStr);

            string addr1     = AlignLeft(Str(MasterRow, MCustomerAddr1), 50);
            DateTime stmtDt  = ToDateTime(MasterRow, MStatementDateTo);
            Output.WriteLine("    " + addr1 + AlignRight(stmtDt.ToString("dd/MM/yyyy"), 25));

            string addr2     = AlignLeft(Str(MasterRow, MCustomerAddr2), 50);
            Output.WriteLine("    " + addr2 + AlignRight(CurMainCard, 25));

            // Debit card header: Available Balance | Statement Date | Due Date
            decimal avail    = ToDecimal(MasterRow, MAvailableLim);
            DateTime dueDate = ToDateTime(MasterRow, MStatementDueDate);
            Output.WriteLine(
                "    " +
                AlignLeft("Available Balance", 22) +
                AlignRight($"{avail:N2}", 16) + "  " +
                AlignRight(dueDate.ToString("dd/MM/yyyy"), 14));
        }
    }

    /// <summary>AIB Debit text label.</summary>
    public sealed class TextLabelDbAibFormatter : NativeTextLabelFormatterBase
    {
        public override string FormatterKey => "TEXTLABEL_AIB_DB";

        protected override void PrintHeader(GenerateStatementCommand cmd)
        {
            string custName = AlignLeft(Str(MasterRow, MCustomerName), 50);
            Output.WriteLine(CurrentPageFlag + " " + custName +
                AlignRight($"Page {PageNo} of {TotalAccPages}", 25));

            string addr1    = AlignLeft(Str(MasterRow, MCustomerAddr1), 50);
            DateTime stmtDt = ToDateTime(MasterRow, MStatementDateTo);
            Output.WriteLine("    " + addr1 + AlignRight(stmtDt.ToString("dd/MM/yyyy"), 25));

            string addr2    = AlignLeft(Str(MasterRow, MCustomerAddr2), 50);
            Output.WriteLine("    " + addr2 + AlignRight(CurAccountNo, 25));

            // Available balance and due date
            decimal avail    = ToDecimal(MasterRow, MAvailableLim);
            DateTime dueDate = ToDateTime(MasterRow, MStatementDueDate);
            Output.WriteLine(
                "    " +
                AlignLeft("Available Balance", 22) +
                AlignRight($"{avail:N2}", 16) + "  " +
                AlignRight(dueDate.ToString("dd/MM/yyyy"), 14));
        }
    }

    /// <summary>BCA Debit text label.</summary>
    public sealed class TextLabelDbBcaFormatter : NativeTextLabelFormatterBase
    {
        public override string FormatterKey => "TEXTLABEL_BCA_DB";
    }

    /// <summary>ICBG Debit text label.</summary>
    public sealed class TextLabelDbIcbgFormatter : NativeTextLabelFormatterBase
    {
        public override string FormatterKey => "TEXTLABEL_ICBG_DB";
    }

    /// <summary>EDBE text label.</summary>
    public sealed class TextLabelEdbeFormatter : NativeTextLabelFormatterBase
    {
        public override string FormatterKey => "TEXTLABEL_EDBE";
    }

    /// <summary>FABG text label.</summary>
    public sealed class TextLabelFabgFormatter : NativeTextLabelFormatterBase
    {
        public override string FormatterKey => "TEXTLABEL_FABG";
    }

    // ── Plain-text (control-character) formatters ─────────────────────────────────

    /// <summary>
    /// EDBE plain-text statement with form-feed (\f) and carriage-return (\r)
    /// control characters for physical printing.
    /// End-of-statement marker: "--- END OF STATEMENT ---".
    /// </summary>
    public sealed class TextEdbeFormatter : NativeTextLabelFormatterBase
    {
        public override string FormatterKey => "TEXT_EDBE";

        protected override void PrintHeader(GenerateStatementCommand cmd)
        {
            // Form-feed before each page except the first
            if (PageNo > 1)
                Output.Write('\f');

            string custName = AlignLeft(Str(MasterRow, MCustomerName), 50);
            Output.WriteLine(CurrentPageFlag + " " + custName +
                AlignRight($"Page {PageNo} of {TotalAccPages}", 25));

            DateTime stmtDt = ToDateTime(MasterRow, MStatementDateTo);
            string addr1    = AlignLeft(Str(MasterRow, MCustomerAddr1), 50);
            Output.WriteLine("    " + addr1 + AlignRight(stmtDt.ToString("dd/MM/yyyy"), 25));

            string addr2    = AlignLeft(Str(MasterRow, MCustomerAddr2), 50);
            Output.WriteLine("    " + addr2 + AlignRight(CurAccountNo, 25));
        }

        protected override void PrintCardFooter(GenerateStatementCommand cmd, StatementDataContext ctx)
        {
            base.PrintCardFooter(cmd, ctx);
            Output.WriteLine("--- END OF STATEMENT ---");
            Output.Write('\r');
        }
    }

}

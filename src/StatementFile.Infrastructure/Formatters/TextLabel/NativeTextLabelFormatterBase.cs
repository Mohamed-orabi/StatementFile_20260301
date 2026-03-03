using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using StatementFile.Application.Interfaces;
using StatementFile.Application.UseCases.StatementGeneration;

namespace StatementFile.Infrastructure.Formatters.TextLabel
{
    /// <summary>
    /// Native base class for all fixed-width text-label (physical printer) formatters.
    /// Replaces the legacy clsStatTxtLbl hierarchy completely.
    ///
    /// Page flag system (preserved exactly from clsStatTxtLbl.cs):
    ///   F 0 = single-page statement
    ///   F 1 = first page of a multi-page statement
    ///   F 2 = middle page
    ///   F 3 = last page
    ///
    /// End-of-page marker = 80 spaces (matching strEndOfPage in legacy).
    ///
    /// Column header printed by each bank's subclass; the base class handles
    /// the iteration, page-break logic and file I/O.
    /// </summary>
    public abstract class NativeTextLabelFormatterBase : IStatementFormatterService
    {
        public abstract string FormatterKey { get; }

        // ── Column constants (tStatementMasterTable) ──────────────────────────────
        protected const string MCardNo             = "cardno";
        protected const string MAccountNo          = "accountno";
        protected const string MCardPrimary        = "cardprimary";
        protected const string MPrimaryCardNo      = "primarycardno";
        protected const string MClientId           = "idclient";
        protected const string MCustomerName       = "customername";
        protected const string MCustomerAddr1      = "customeraddress1";
        protected const string MCustomerAddr2      = "customeraddress2";
        protected const string MCustomerAddr3      = "customeraddress3";
        protected const string MCustomerRegion     = "customerregion";
        protected const string MCustomerCity       = "customercity";
        protected const string MContractType       = "contracttype";
        protected const string MCardBranchPart     = "cardbranchpart";
        protected const string MCardBranchPartName = "cardbranchpartname";
        protected const string MStatementDateTo    = "statementdateto";
        protected const string MAccountLim         = "accountlim";
        protected const string MAvailableLim       = "accountavailablelim";
        protected const string MMinDue             = "mindueamount";
        protected const string MClosingBalance     = "closingbalance";
        protected const string MStatementDueDate   = "statementduedate";
        protected const string MTotalOverdue       = "totaloverdueamount";
        protected const string MOpeningBalance     = "openingbalance";
        protected const string MTotalPayments      = "totalpayments";
        protected const string MTotalPurchases     = "totalpurchases";
        protected const string MTotalCash          = "totalcashwithdrawal";
        protected const string MTotalCharges       = "totalcharges";
        protected const string MTotalInterest      = "totalinterest";
        protected const string MExternalNo         = "externalno";
        protected const string MCardState         = "cardstate";
        protected const string MAccountCurrency    = "accountcurrency";
        protected const string MCardProduct        = "cardproduct";

        // ── Column constants (tStatementDetailTable) ──────────────────────────────
        protected const string DCardNo            = "cardno";
        protected const string DAccountNo         = "accountno";
        protected const string DPostingDate       = "postingdate";
        protected const string DDocNo             = "docno";
        protected const string DTransDate         = "transdate";
        protected const string DRefNo             = "refereneno";
        protected const string DTranDesc          = "trandescription";
        protected const string DMerchant          = "merchant";
        protected const string DOrigTranAmount    = "origtranamount";
        protected const string DOrigTranCurrency  = "origtrancurrency";
        protected const string DBillTranAmount    = "billtranamount";
        protected const string DBillTranAmountSign = "billtranamountsign";

        // End-of-page / end-of-account markers (80 spaces, matching legacy)
        protected static readonly string EndOfPage    = new string(' ', 80);
        protected static readonly string EndOfAccount = new string(' ', 80);

        protected virtual int MaxDetailInPage     => 20;
        protected virtual int MaxDetailInLastPage => 27;

        // ── Per-account state (safe to use from PrintHeader/Detail/Footer) ────────
        protected DataRow   MasterRow      { get; private set; }
        protected DataRow   DetailRow      { get; private set; }
        protected DataRow[] AccountDetails { get; private set; }
        protected StreamWriter Output      { get; private set; }
        protected string CurrentPageFlag   { get; private set; } = "F 0";
        protected int    PageNo            { get; private set; } = 1;
        protected int    TotalAccPages     { get; private set; } = 1;
        protected int    TotAccRows        { get; private set; }
        protected string CurAccountNo      { get; private set; } = string.Empty;
        protected string CurMainCard       { get; private set; } = string.Empty;

        // ── IStatementFormatterService ────────────────────────────────────────────

        public IEnumerable<string> Format(
            StatementDataContext     ctx,
            string                   outputDirectory,
            GenerateStatementCommand cmd)
        {
            Directory.CreateDirectory(outputDirectory);

            string baseName   = Path.Combine(outputDirectory,
                $"{cmd.BankName}_{cmd.StatementDate:yyyyMM}");
            string outputFile = baseName + ".txt";
            var outputFiles   = new List<string>();

            var masterTable = ctx.MasterDataSet?.Tables["tStatementMasterTable"];
            if (masterTable == null || masterTable.Rows.Count == 0)
                return outputFiles;

            var detailTable = ctx.MasterDataSet.Tables.Count > 1
                ? ctx.MasterDataSet.Tables["tStatementDetailTable"]
                : ctx.DetailDataSet?.Tables["tStatementDetailTable"];

            int totalStatements = 0;

            using (Output = new StreamWriter(outputFile, false, Encoding.UTF8))
            {
                string prevAccountNo = string.Empty;

                foreach (DataRow mRow in masterTable.Rows)
                {
                    MasterRow = mRow;
                    string acctNo = Str(mRow, MAccountNo);
                    if (acctNo == prevAccountNo) continue;
                    prevAccountNo = acctNo;
                    CurAccountNo  = acctNo;

                    // Load detail rows for this account
                    AccountDetails = detailTable != null
                        ? detailTable.Select($"ACCOUNTNO = '{EscSql(acctNo)}'")
                        : Array.Empty<DataRow>();

                    // Find primary card
                    CurMainCard = Str(mRow, MCardNo);
                    DataRow[] mainRows = masterTable.Select($"ACCOUNTNO = '{EscSql(acctNo)}'");
                    foreach (DataRow mr in mainRows)
                    {
                        if (Str(mr, MCardPrimary) == "Y")
                        {
                            CurMainCard = Str(mr, MCardNo);
                            break;
                        }
                    }

                    // Count valid rows
                    TotAccRows = 0;
                    foreach (DataRow dr in AccountDetails)
                        if (!IsOnHold(dr)) TotAccRows++;

                    decimal closing = ToDecimal(mRow, MClosingBalance);
                    if (TotAccRows < 1 && closing == 0m) continue;

                    // Compute total pages
                    TotalAccPages = TotAccRows <= MaxDetailInPage
                        ? 1
                        : 1 + (int)Math.Ceiling((double)(TotAccRows - MaxDetailInLastPage) / MaxDetailInPage);
                    if (TotalAccPages < 1) TotalAccPages = 1;

                    // Initialise page state
                    PageNo         = 1;
                    int curPageRec = 0;
                    CurrentPageFlag = TotalAccPages == 1 ? "F 0" : "F 1";

                    PrintHeader(cmd);
                    totalStatements++;

                    foreach (DataRow dr in AccountDetails)
                    {
                        if (IsOnHold(dr)) continue;
                        DetailRow = dr;

                        if (curPageRec >= MaxDetailInPage)
                        {
                            Output.WriteLine(EndOfPage);
                            PageNo++;
                            curPageRec      = 0;
                            CurrentPageFlag = PageNo == TotalAccPages ? "F 3" : "F 2";
                            PrintHeader(cmd);
                        }

                        PrintDetail(cmd);
                        curPageRec++;
                    }

                    PrintCardFooter(cmd, ctx);
                    Output.WriteLine(EndOfAccount);
                }
            }

            outputFiles.Add(outputFile);

            // Summary
            string summaryFile = baseName + "_Summary.txt";
            using (var sw = new StreamWriter(summaryFile, false, Encoding.UTF8))
            {
                sw.WriteLine($"{cmd.BankFullName} - Text Label Statement");
                sw.WriteLine("__________________________");
                sw.WriteLine($"No of Statements   {totalStatements}");
            }
            outputFiles.Add(summaryFile);

            return outputFiles;
        }

        // ── Virtual print methods – override in bank-specific subclasses ──────────

        /// <summary>
        /// Prints the page header. Called once per page.
        /// CurrentPageFlag, PageNo, TotalAccPages, CurAccountNo, CurMainCard are set.
        /// </summary>
        protected virtual void PrintHeader(GenerateStatementCommand cmd)
        {
            string custName  = AlignLeft(Str(MasterRow, MCustomerName), 50);
            string pageStr   = AlignRight($"Page {PageNo} of {TotalAccPages}", 25);
            Output.WriteLine(CurrentPageFlag + " " + custName + pageStr);

            string addr1     = AlignLeft(Str(MasterRow, MCustomerAddr1), 50);
            DateTime stmtDt  = ToDateTime(MasterRow, MStatementDateTo);
            Output.WriteLine("    " + addr1 + AlignRight(stmtDt.ToString("dd/MM/yyyy"), 25));

            string addr2     = AlignLeft(Str(MasterRow, MCustomerAddr2), 50);
            Output.WriteLine("    " + addr2 + AlignRight(CurAccountNo, 25));

            string city      = AlignLeft(
                Str(MasterRow, MCustomerRegion) + " " + Str(MasterRow, MCustomerCity), 50);
            Output.WriteLine("    " + city);
        }

        /// <summary>Prints one transaction line. DetailRow is set for the current row.</summary>
        protected virtual void PrintDetail(GenerateStatementCommand cmd)
        {
            DateTime trnsDate = ToDateTime(DetailRow, DTransDate);
            DateTime postDate = ToDateTime(DetailRow, DPostingDate);
            if (trnsDate > postDate) trnsDate = postDate;

            string desc      = Str(DetailRow, DMerchant);
            if (string.IsNullOrEmpty(desc)) desc = Str(DetailRow, DTranDesc);

            string origCurr  = Str(DetailRow, DOrigTranCurrency);
            string acctCurr  = Str(MasterRow, MAccountCurrency);
            decimal origAmt  = ToDecimal(DetailRow, DOrigTranAmount);
            decimal billAmt  = ToDecimal(DetailRow, DBillTranAmount);
            string sign      = Str(DetailRow, DBillTranAmountSign);
            bool isCr        = sign == "CR";

            string foreignStr = (acctCurr != origCurr && !string.IsNullOrWhiteSpace(origCurr) && origCurr != "XXX")
                ? AlignRight($"{origAmt:##,###,##0.00} {origCurr}", 16)
                : new string(' ', 16);

            string billStr = AlignRight($"{billAmt:##,###,##0.00}", 16);
            string crdb    = isCr ? "CR" : "  ";
            string refStr  = AlignLeft(TrimStr(Str(DetailRow, DRefNo), 24), 24);
            string descStr = AlignLeft(TrimStr(desc, 40), 40);

            Output.WriteLine(
                $"  {trnsDate:dd/MM}  {postDate:dd/MM}  {refStr}  {descStr} {foreignStr} {billStr} {crdb}");
        }

        /// <summary>Prints the closing balance / card footer.</summary>
        protected virtual void PrintCardFooter(GenerateStatementCommand cmd, StatementDataContext ctx)
        {
            decimal closing = ToDecimal(MasterRow, MClosingBalance);
            string sign     = closing >= 0 ? "DB" : "CR";
            string balStr   = AlignRight($"{Math.Abs(closing):##,###,##0.00}", 16);
            string label    = AlignLeft("Current Balance", 67);
            Output.WriteLine(new string(' ', 43) + label + balStr + " " + sign);
        }

        // ── Text alignment helpers ────────────────────────────────────────────────

        protected static string AlignLeft(string s, int width)
        {
            s = s ?? string.Empty;
            return s.Length >= width ? s.Substring(0, width) : s.PadRight(width);
        }

        protected static string AlignRight(string s, int width)
        {
            s = s ?? string.Empty;
            return s.Length >= width ? s.Substring(s.Length - width) : s.PadLeft(width);
        }

        protected static string AlignMiddle(string s, int width)
        {
            s = s ?? string.Empty;
            if (s.Length >= width) return s.Substring(0, width);
            int total = width - s.Length;
            int left  = total / 2;
            return new string(' ', left) + s + new string(' ', total - left);
        }

        protected static string TrimStr(string s, int maxLen)
            => s?.Length > maxLen ? s.Substring(0, maxLen) : (s ?? string.Empty);

        protected static string Repeat(char c, int count)
            => new string(c, Math.Max(0, count));

        // ── Data helpers ──────────────────────────────────────────────────────────

        protected static bool IsOnHold(DataRow dr)
            => dr[DPostingDate] == DBNull.Value && dr[DDocNo] == DBNull.Value;

        protected static string Str(DataRow row, string col)
        {
            try { return row.IsNull(col) ? string.Empty : (row[col]?.ToString()?.Trim() ?? string.Empty); }
            catch { return string.Empty; }
        }

        protected static decimal ToDecimal(DataRow row, string col)
        {
            try { return row.IsNull(col) ? 0m : Convert.ToDecimal(row[col]); }
            catch { return 0m; }
        }

        protected static DateTime ToDateTime(DataRow row, string col)
        {
            try
            {
                if (row.IsNull(col)) return DateTime.MinValue;
                return row[col] is DateTime dt ? dt : Convert.ToDateTime(row[col]);
            }
            catch { return DateTime.MinValue; }
        }

        protected static string EscSql(string s) => (s ?? string.Empty).Replace("'", "''");
    }
}

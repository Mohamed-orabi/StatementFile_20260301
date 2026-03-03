using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using StatementFile.Application.Interfaces;
using StatementFile.Application.UseCases.StatementGeneration;

namespace StatementFile.Infrastructure.Formatters.RawData
{
    /// <summary>
    /// Native base class for all raw-data formatters.
    /// Replaces the legacy clsStatRawData hierarchy completely.
    ///
    /// Default output: pipe-delimited _Main.txt + _Trns.txt (matching the legacy
    /// clsStatRawData.printHeader / printDetail layout).
    /// An MD5 checksum file is generated alongside the data files.
    ///
    /// AIBK/EGB variants use STMT_HDR.DAT + STMT_DTL.DAT naming; those
    /// subclasses override GetMainFilePath() / GetTransFilePath().
    /// </summary>
    public abstract class NativeRawDataFormatterBase : IStatementFormatterService
    {
        public abstract string FormatterKey { get; }

        // ── Column constants (tStatementMasterTable) ──────────────────────────────
        protected const string MCardNo             = "cardno";
        protected const string MAccountNo          = "accountno";
        protected const string MCardAccountNo      = "cardaccountno";
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
        protected const string MStatementNo        = "statementno";
        protected const string MCardState         = "cardstate";
        protected const string MAccountCurrency    = "accountcurrency";
        protected const string MCardProduct        = "cardproduct";

        // ── Column constants (tStatementDetailTable) ──────────────────────────────
        protected const string DCardNo            = "cardno";
        protected const string DStatementNo       = "statementno";
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

        protected virtual string FieldSeparator  => "|";
        protected virtual string TransDateFormat => "dd/MM";

        // corporate mode uses cardaccountno
        protected virtual string AccountNoField  => MAccountNo;

        // ── IStatementFormatterService ────────────────────────────────────────────

        public IEnumerable<string> Format(
            StatementDataContext     ctx,
            string                   outputDirectory,
            GenerateStatementCommand cmd)
        {
            Directory.CreateDirectory(outputDirectory);

            string baseName  = Path.Combine(outputDirectory,
                $"{cmd.BankName}_{cmd.StatementDate:yyyyMM}");

            string mainFile  = GetMainFilePath(baseName, cmd);
            string transFile = GetTransFilePath(baseName, cmd);
            var outputFiles  = new List<string>();

            var masterTable = ctx.MasterDataSet?.Tables["tStatementMasterTable"];
            if (masterTable == null)
                return outputFiles;

            var detailTable = ctx.MasterDataSet.Tables.Count > 1
                ? ctx.MasterDataSet.Tables["tStatementDetailTable"]
                : ctx.DetailDataSet?.Tables["tStatementDetailTable"];

            int totalStatements   = 0;
            int totalTransactions = 0;

            using var swMain  = new StreamWriter(mainFile,  false, Encoding.UTF8);
            using var swTrans = new StreamWriter(transFile, false, Encoding.UTF8);

            WriteMainHeader(swMain,  cmd);
            WriteTransHeader(swTrans, cmd);

            string prevAccountNo = string.Empty;
            string curMainCard   = string.Empty;

            foreach (DataRow mRow in masterTable.Rows)
            {
                string accountNo = Str(mRow, AccountNoField);
                if (accountNo == prevAccountNo) continue; // one header per account

                prevAccountNo = accountNo;

                // Get all detail rows for this account
                DataRow[] accountDetails = detailTable != null
                    ? detailTable.Select($"ACCOUNTNO = '{EscSql(accountNo)}'")
                    : Array.Empty<DataRow>();

                // Determine primary card
                curMainCard = Str(mRow, MCardNo);
                DataRow[] mainRows = masterTable.Select($"ACCOUNTNO = '{EscSql(accountNo)}'");
                foreach (DataRow mr in mainRows)
                {
                    if (Str(mr, MCardPrimary) == "Y")
                    {
                        curMainCard = Str(mr, MCardNo);
                        break;
                    }
                }

                // Count real transactions
                int realTrans = 0;
                foreach (DataRow dr in accountDetails)
                    if (!IsOnHold(dr)) realTrans++;

                decimal closingBal = ToDecimal(mRow, MClosingBalance);
                if (realTrans < 1 && closingBal == 0m) continue;

                WriteMainRow(swMain, mRow, accountDetails, curMainCard, accountNo, ctx, cmd);
                totalStatements++;

                foreach (DataRow dr in accountDetails)
                {
                    if (IsOnHold(dr)) continue;
                    WriteTransRow(swTrans, mRow, dr, curMainCard, ctx, cmd);
                    totalTransactions++;
                }
            }

            outputFiles.Add(mainFile);
            outputFiles.Add(transFile);

            // Summary
            string summaryFile = baseName + "_Summary.txt";
            using (var sw = new StreamWriter(summaryFile, false, Encoding.UTF8))
            {
                sw.WriteLine($"{cmd.BankFullName} - Statement (Raw Data)");
                sw.WriteLine("__________________________");
                sw.WriteLine($"No of Statements   {totalStatements}");
                sw.WriteLine($"No of Transactions {totalTransactions}");
            }
            outputFiles.Add(summaryFile);

            // MD5 checksum
            string md5File = baseName + ".MD5";
            WriteMd5File(md5File, mainFile, transFile);
            outputFiles.Add(md5File);

            return outputFiles;
        }

        // ── Overridable file-name providers ──────────────────────────────────────

        protected virtual string GetMainFilePath(string baseName, GenerateStatementCommand cmd)
            => baseName + "_Main.txt";

        protected virtual string GetTransFilePath(string baseName, GenerateStatementCommand cmd)
            => baseName + "_Trns.txt";

        // ── Overridable header / row writers ─────────────────────────────────────

        protected virtual void WriteMainHeader(TextWriter sw, GenerateStatementCommand cmd)
        {
            string s = FieldSeparator;
            sw.WriteLine(
                $"CardHolder Name{s}Address1{s}Address2{s}Address3{s}Address4{s}" +
                $"Card Type{s}Account{s}Branch{s}Statement Date{s}Card No.{s}" +
                $"Credit Limit{s}Available Credit{s}Min. Payment Due{s}New Balance{s}" +
                $"Payment Due Date{s}Past Due Amount{s}prev.balance{s}payment & CR{s}" +
                $"Purch.Cash&Dr{s}Finance Charge{s}External NO{s}");
        }

        protected virtual void WriteTransHeader(TextWriter sw, GenerateStatementCommand cmd)
        {
            string s = FieldSeparator;
            sw.WriteLine(
                $"Card No.{s}Date of Trans{s}Date of Post{s}Reference{s}" +
                $"Description{s}Purchase Currency & Amount{s}Amount{s}");
        }

        protected virtual void WriteMainRow(
            TextWriter sw, DataRow mRow, DataRow[] detailRows,
            string primaryCardNo, string accountNo,
            StatementDataContext ctx, GenerateStatementCommand cmd)
        {
            string s       = FieldSeparator;
            string addr1   = Str(mRow, MCustomerAddr1);
            string addr2   = Str(mRow, MCustomerAddr2);
            string addr3   = Str(mRow, MCustomerAddr3);
            string region  = Str(mRow, MCustomerRegion);
            string city    = Str(mRow, MCustomerCity);
            string extNo   = Str(mRow, MExternalNo);
            if (string.IsNullOrWhiteSpace(extNo)) extNo = accountNo;

            DateTime stmtDate = ToDateTime(mRow, MStatementDateTo);
            DateTime dueDate  = ToDateTime(mRow, MStatementDueDate);
            decimal opening   = ToDecimal(mRow, MOpeningBalance);
            decimal payments  = ToDecimal(mRow, MTotalPayments);
            decimal purchases = ToDecimal(mRow, MTotalPurchases) + ToDecimal(mRow, MTotalCash);
            decimal charges   = ToDecimal(mRow, MTotalCharges)   + ToDecimal(mRow, MTotalInterest);
            decimal closing   = ToDecimal(mRow, MClosingBalance);

            sw.WriteLine(
                Str(mRow, MCustomerName) + s +
                addr1 + s +
                addr2 + s +
                addr3 + s +
                region + "  " + city + s +
                Str(mRow, MContractType) + s +
                accountNo + s +
                Str(mRow, MCardBranchPart) + "  " + Str(mRow, MCardBranchPartName) + s +
                stmtDate.ToString("dd/MM/yy") + s +
                FormatCardNo(primaryCardNo) + s +
                Str(mRow, MAccountLim) + s +
                Str(mRow, MAvailableLim) + s +
                Str(mRow, MMinDue) + s +
                Str(mRow, MClosingBalance) + CrDb(closing) + s +
                dueDate.ToString("dd/MM/yy") + s +
                Str(mRow, MTotalOverdue) + s +
                FormatNum(opening) + " " + CrDb(opening) + s +
                FormatNum(payments)  + DbCr(payments) + s +
                FormatNum(purchases) + s +
                FormatNum(charges) + s +
                extNo + s
            );
        }

        protected virtual void WriteTransRow(
            TextWriter sw, DataRow mRow, DataRow dr,
            string primaryCardNo, StatementDataContext ctx, GenerateStatementCommand cmd)
        {
            string s         = FieldSeparator;
            DateTime trnsDate = ToDateTime(dr, DTransDate);
            DateTime postDate = ToDateTime(dr, DPostingDate);
            if (trnsDate > postDate) trnsDate = postDate;

            string desc = Str(dr, DMerchant);
            if (string.IsNullOrEmpty(desc)) desc = Str(dr, DTranDesc);

            string origCurr  = Str(dr, DOrigTranCurrency);
            string acctCurr  = Str(mRow, MAccountCurrency);
            decimal origAmt  = ToDecimal(dr, DOrigTranAmount);
            decimal billAmt  = ToDecimal(dr, DBillTranAmount);
            string sign      = Str(dr, DBillTranAmountSign);
            bool isCr        = sign == "CR";

            string foreignAmt = (acctCurr != origCurr && !string.IsNullOrWhiteSpace(origCurr) && origCurr != "XXX")
                ? $"{origAmt:##,###,##0.00} {origCurr}"
                : new string(' ', 16);

            sw.WriteLine(
                primaryCardNo + s +
                trnsDate.ToString(TransDateFormat) + s +
                postDate.ToString(TransDateFormat) + s +
                TrimStr(Str(dr, DRefNo), 25) + s +
                TrimStr(desc, 40) + s +
                foreignAmt + s +
                FormatNum(billAmt) + (isCr ? "CR" : string.Empty) + s
            );
        }

        // ── MD5 ───────────────────────────────────────────────────────────────────

        protected static void WriteMd5File(string md5File, params string[] files)
        {
            using var sw = new StreamWriter(md5File, false, Encoding.UTF8);
            foreach (string f in files)
            {
                if (!File.Exists(f)) continue;
                sw.WriteLine($"{Path.GetFileName(f)}  >>  {ComputeMd5(f)}");
            }
        }

        private static string ComputeMd5(string filePath)
        {
            using var md5 = MD5.Create();
            using var fs  = File.OpenRead(filePath);
            return BitConverter.ToString(md5.ComputeHash(fs)).Replace("-", string.Empty).ToUpperInvariant();
        }

        // ── Formatting helpers ────────────────────────────────────────────────────

        protected static string FormatCardNo(string cardNo)
        {
            if (cardNo?.Length == 16)
                return $"{cardNo[..4]}-{cardNo[4..8]}-{cardNo[8..12]}-{cardNo[12..]}";
            return cardNo ?? string.Empty;
        }

        protected static string FormatNum(decimal val)
        {
            if (val < 0) return $"({Math.Abs(val):##,###,##0.00})";
            return $"{val:##,###,##0.00}";
        }

        protected static string CrDb(decimal val) => val < 0 ? "CR" : string.Empty;
        protected static string DbCr(decimal val) => val < 0 ? "CR" : string.Empty;

        protected static string TrimStr(string s, int maxLen)
            => s?.Length > maxLen ? s.Substring(0, maxLen) : (s ?? string.Empty);

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

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using StatementFile.Application.UseCases.StatementGeneration;
using StatementFile.Domain.Enums;
using StatementFile.Domain.Interfaces.Services;

namespace StatementFile.Infrastructure.Formatters.Html
{
    /// <summary>
    /// Native base class for all HTML e-statement formatters.
    /// Replaces the legacy clsStatHtml hierarchy completely.
    /// Works solely from the pre-loaded StatementDataContext — no Oracle connections,
    /// no legacy class references.
    ///
    /// One HTML file per account is created in <paramref name="outputDirectory"/>.
    /// Email tracking files (Emails.txt / WithoutEmails.txt) and a Summary.txt are
    /// always written alongside the statement files.
    ///
    /// Subclasses customise branding by overriding the virtual branding methods and
    /// may override AppendAccountSummary / AppendTransactionTable / AppendCardFooter
    /// for bank-specific HTML layout variations.
    /// </summary>
    public abstract class NativeHtmlFormatterBase : IStatementFormatterService
    {
        public abstract string FormatterKey { get; }

        // ── Column name constants (tStatementMasterTable) ─────────────────────────
        protected const string MCardNo              = "cardno";
        protected const string MAccountNo           = "accountno";
        protected const string MCardAccountNo       = "cardaccountno";
        protected const string MCardPrimary         = "cardprimary";
        protected const string MPrimaryCardNo       = "primarycardno";
        protected const string MClientId            = "idclient";
        protected const string MCustomerName        = "customername";
        protected const string MCustomerAddr1       = "customeraddress1";
        protected const string MCustomerAddr2       = "customeraddress2";
        protected const string MCustomerAddr3       = "customeraddress3";
        protected const string MCustomerRegion      = "customerregion";
        protected const string MCustomerCity        = "customercity";
        protected const string MContractType        = "contracttype";
        protected const string MCardBranchPart      = "cardbranchpart";
        protected const string MCardBranchPartName  = "cardbranchpartname";
        protected const string MStatementDateTo     = "statementdateto";
        protected const string MAccountLim          = "accountlim";
        protected const string MAvailableLim        = "accountavailablelim";
        protected const string MMinDue              = "mindueamount";
        protected const string MClosingBalance      = "closingbalance";
        protected const string MStatementDueDate    = "statementduedate";
        protected const string MTotalOverdue        = "totaloverdueamount";
        protected const string MOpeningBalance      = "openingbalance";
        protected const string MTotalPayments       = "totalpayments";
        protected const string MTotalPurchases      = "totalpurchases";
        protected const string MTotalCash           = "totalcashwithdrawal";
        protected const string MTotalCharges        = "totalcharges";
        protected const string MTotalInterest       = "totalinterest";
        protected const string MExternalNo          = "externalno";
        protected const string MStatementNo         = "statementno";
        protected const string MCardState           = "cardstate";
        protected const string MAccountCurrency     = "accountcurrency";
        protected const string MCardProduct         = "cardproduct";

        // ── Column name constants (tStatementDetailTable) ─────────────────────────
        protected const string DCardNo             = "cardno";
        protected const string DStatementNo        = "statementno";
        protected const string DAccountNo          = "accountno";
        protected const string DPostingDate        = "postingdate";
        protected const string DDocNo              = "docno";
        protected const string DTransDate          = "transdate";
        protected const string DRefNo              = "refereneno";
        protected const string DTranDesc           = "trandescription";
        protected const string DMerchant           = "merchant";
        protected const string DOrigTranAmount     = "origtranamount";
        protected const string DOrigTranCurrency   = "origtrancurrency";
        protected const string DBillTranAmount     = "billtranamount";
        protected const string DBillTranAmountSign = "billtranamountsign";

        /// <summary>
        /// When true the account number is read from cardaccountno instead of accountno
        /// (Corporate mode: cmd.UseCorporateAccountNumber).
        /// </summary>
        protected virtual bool UseCorporateAccountNo => false;

        private string AccountNoField => UseCorporateAccountNo ? MCardAccountNo : MAccountNo;

        // ── IStatementFormatterService ────────────────────────────────────────────

        public IEnumerable<string> Format(
            StatementDataContext     ctx,
            string                   outputDirectory,
            GenerateStatementCommand cmd)
        {
            Directory.CreateDirectory(outputDirectory);

            string prefix       = $"{cmd.BankName}_{cmd.StatementTypeSuffix}_{cmd.StatementDate:yyyyMM}";
            string emailsFile   = Path.Combine(outputDirectory, prefix + "_Emails.txt");
            string noEmailsFile = Path.Combine(outputDirectory, prefix + "_WithoutEmails.txt");
            string summaryFile  = Path.Combine(outputDirectory, prefix + "_Summary.txt");

            var outputFiles = new List<string>();

            int totalStatements   = 0;
            int totalTransactions = 0;
            int withEmailCount    = 0;
            int withoutEmailCount = 0;

            using var swEmails   = new StreamWriter(emailsFile,   false, Encoding.UTF8);
            using var swNoEmails = new StreamWriter(noEmailsFile,  false, Encoding.UTF8);
            using var swSummary  = new StreamWriter(summaryFile,   false, Encoding.UTF8);

            swEmails  .WriteLine("AccountNumber|ClientID|Email|MobilePhone|Date Time");
            swNoEmails.WriteLine("AccountNumber|ClientID|Email|Mobile Phone");

            var masterTable = ctx.MasterDataSet?.Tables["tStatementMasterTable"];
            if (masterTable == null || masterTable.Rows.Count == 0)
            {
                swSummary.WriteLine("No data to process.");
                AddTrackingFiles(outputFiles, emailsFile, noEmailsFile, summaryFile);
                return outputFiles;
            }

            // Detail may be a second table in the master DataSet, or a separate DataSet
            var detailTable = ctx.MasterDataSet.Tables.Count > 1
                ? ctx.MasterDataSet.Tables["tStatementDetailTable"]
                : ctx.DetailDataSet?.Tables["tStatementDetailTable"];

            string prevAccountNo = string.Empty;

            foreach (DataRow mRow in masterTable.Rows)
            {
                string cardNo    = Str(mRow, MCardNo);
                if (cardNo.Length != 16) continue;

                string accountNo = Str(mRow, AccountNoField);
                if (accountNo == prevAccountNo) continue; // One statement per account

                prevAccountNo = accountNo;

                // Get all detail rows for this account
                DataRow[] accountDetails = detailTable != null
                    ? detailTable.Select($"ACCOUNTNO = '{EscSql(accountNo)}'")
                    : Array.Empty<DataRow>();

                // Count real (non-on-hold) transactions
                int realTrans = 0;
                foreach (DataRow dr in accountDetails)
                    if (!IsOnHold(dr)) realTrans++;

                decimal closingBalance = ToDecimal(mRow, MClosingBalance);
                if (realTrans < 1 && closingBalance == 0m) continue;

                totalStatements++;
                totalTransactions += realTrans;

                // Primary card number
                string primaryCardNo = cardNo;
                if (Str(mRow, MCardPrimary) == "N")
                    primaryCardNo = Str(mRow, MPrimaryCardNo);

                // Client email
                string clientId  = Str(mRow, MClientId);
                string emailAddr = string.Empty;
                string mobileNo  = string.Empty;
                if (ctx.EmailDataSet?.Tables["Emails"] != null)
                {
                    foreach (DataRow er in ctx.EmailDataSet.Tables["Emails"]
                             .Select($"idclient = {clientId}"))
                    {
                        string e = Str(er, 1);
                        string m = Str(er, 2);
                        if (!string.IsNullOrWhiteSpace(e)) emailAddr = e;
                        if (!string.IsNullOrWhiteSpace(m)) mobileNo  = m;
                    }
                }

                // Generate HTML
                var sb = new StringBuilder(8192);
                BuildStatement(sb, mRow, accountDetails, primaryCardNo, accountNo, ctx, cmd);

                string htmlFile = Path.Combine(outputDirectory, $"{primaryCardNo}.html");
                File.WriteAllText(htmlFile, sb.ToString(), Encoding.UTF8);
                outputFiles.Add(htmlFile);

                // Email tracking
                string now = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                if (!string.IsNullOrWhiteSpace(emailAddr))
                {
                    swEmails.WriteLine($"{accountNo}|{clientId}|{emailAddr}|{mobileNo}|{now}");
                    withEmailCount++;
                }
                else
                {
                    swNoEmails.WriteLine($"{accountNo}|{clientId}||{mobileNo}|Without Email");
                    withoutEmailCount++;
                }
            }

            // Summary
            swSummary.WriteLine($"{cmd.BankFullName} - Statement");
            swSummary.WriteLine("__________________________");
            swSummary.WriteLine(string.Empty);
            swSummary.WriteLine($"No of Statements   {totalStatements}");
            swSummary.WriteLine($"No of Transactions {totalTransactions}");
            swSummary.WriteLine($"With Email         {withEmailCount}");
            swSummary.WriteLine($"Without Email      {withoutEmailCount}");

            AddTrackingFiles(outputFiles, emailsFile, noEmailsFile, summaryFile);
            return outputFiles;
        }

        private static void AddTrackingFiles(
            List<string> files,
            string emailsFile, string noEmailsFile, string summaryFile)
        {
            files.Add(emailsFile);
            files.Add(noEmailsFile);
            files.Add(summaryFile);
        }

        // ── HTML construction ─────────────────────────────────────────────────────

        private void BuildStatement(
            StringBuilder sb,
            DataRow        masterRow,
            DataRow[]      detailRows,
            string         primaryCardNo,
            string         accountNo,
            StatementDataContext     ctx,
            GenerateStatementCommand cmd)
        {
            AppendHtmlOpen(sb, masterRow, primaryCardNo, accountNo, ctx, cmd);
            AppendAccountSummary(sb, masterRow, ctx, cmd);
            AppendTransactionTable(sb, masterRow, detailRows, ctx, cmd);
            AppendCardFooter(sb, masterRow, ctx, cmd);
            AppendHtmlClose(sb, ctx, cmd);
        }

        protected virtual void AppendHtmlOpen(
            StringBuilder sb, DataRow masterRow,
            string primaryCardNo, string accountNo,
            StatementDataContext ctx, GenerateStatementCommand cmd)
        {
            string color1    = GetPrimaryColor();
            string color2    = GetSecondaryColor();
            string bankName  = EscHtml(cmd.BankFullName);
            string website   = EscHtml(cmd.BankWebsiteUrl ?? string.Empty);
            DateTime stmtDt  = ToDateTime(masterRow, MStatementDateTo);

            sb.Append($@"<!DOCTYPE html>
<html lang=""en"">
<head>
<meta charset=""UTF-8""/>
<meta name=""viewport"" content=""width=device-width,initial-scale=1.0""/>
<title>{bankName} - Account Statement</title>
<style>
  body{{font-family:Arial,sans-serif;font-size:11px;margin:0;padding:10px;color:#333;}}
  .hdr{{background:{color1};color:#fff;padding:10px 20px;display:flex;justify-content:space-between;align-items:center;}}
  .hdr .title{{font-size:16px;font-weight:bold;}}
  .hdr .dt{{font-size:12px;}}
  .sec-hdr{{background:{color2};color:#fff;padding:4px 10px;font-weight:bold;font-size:11px;margin-top:10px;}}
  table{{width:100%;border-collapse:collapse;margin:0;}}
  th{{background:{color1};color:#fff;padding:4px 8px;text-align:left;font-size:10px;}}
  td{{padding:3px 8px;border-bottom:1px solid #eee;font-size:10px;vertical-align:top;}}
  .amt{{text-align:right;}}
  .cr{{color:green;}}
  .total{{font-weight:bold;background:#f5f5f5;}}
  .ftr{{color:#888;font-size:9px;margin-top:20px;text-align:center;border-top:1px solid #ccc;padding-top:5px;}}
</style>
</head>
<body>
<div class=""hdr"">
  <div class=""title"">{bankName} &mdash; Account Statement</div>
  <div class=""dt"">{stmtDt:MMMM yyyy}</div>
</div>
");
        }

        protected virtual void AppendAccountSummary(
            StringBuilder sb, DataRow masterRow,
            StatementDataContext ctx, GenerateStatementCommand cmd)
        {
            string accountNo  = Str(masterRow, AccountNoField);
            string custName   = Str(masterRow, MCustomerName);
            string addr1      = Str(masterRow, MCustomerAddr1);
            string addr2      = Str(masterRow, MCustomerAddr2);
            string addr3      = Str(masterRow, MCustomerAddr3);
            string region     = Str(masterRow, MCustomerRegion);
            string city       = Str(masterRow, MCustomerCity);
            string branch     = Str(masterRow, MCardBranchPart) + " " + Str(masterRow, MCardBranchPartName);
            string currency   = Str(masterRow, MAccountCurrency);
            string externalNo = Str(masterRow, MExternalNo);
            if (string.IsNullOrWhiteSpace(externalNo)) externalNo = accountNo;

            DateTime stmtDate = ToDateTime(masterRow, MStatementDateTo);
            DateTime dueDate  = ToDateTime(masterRow, MStatementDueDate);
            decimal limit     = ToDecimal(masterRow, MAccountLim);
            decimal avail     = ToDecimal(masterRow, MAvailableLim);
            decimal minDue    = ToDecimal(masterRow, MMinDue);
            decimal closing   = ToDecimal(masterRow, MClosingBalance);
            decimal prev      = ToDecimal(masterRow, MOpeningBalance);
            decimal payments  = ToDecimal(masterRow, MTotalPayments);
            decimal purchases = ToDecimal(masterRow, MTotalPurchases) + ToDecimal(masterRow, MTotalCash);
            decimal charges   = ToDecimal(masterRow, MTotalCharges)   + ToDecimal(masterRow, MTotalInterest);
            decimal overdue   = ToDecimal(masterRow, MTotalOverdue);

            bool isPrepaid = cmd.HasMode(ProcessingMode.Prepaid);

            sb.Append($@"
<div class=""sec-hdr"">Customer Details</div>
<table>
  <tr><td><b>Name</b></td><td>{EscHtml(custName)}</td><td><b>Account No</b></td><td>{EscHtml(externalNo)}</td></tr>
  <tr><td><b>Address</b></td><td>{EscHtml(addr1)}{(string.IsNullOrEmpty(addr2) ? "" : ", " + EscHtml(addr2))}</td><td><b>Branch</b></td><td>{EscHtml(branch.Trim())}</td></tr>
  <tr><td></td><td>{EscHtml(addr3)}{(string.IsNullOrEmpty(city) ? "" : ", " + EscHtml(city))}</td><td><b>Statement Date</b></td><td>{stmtDate:dd/MM/yyyy}</td></tr>
</table>
<div class=""sec-hdr"">Account Summary ({EscHtml(currency)})</div>
<table>
  <tr><td><b>{(isPrepaid ? "Available Balance" : "Credit Limit")}</b></td><td class=""amt"">{(isPrepaid ? avail : limit):N2}</td><td><b>Payment Due Date</b></td><td>{dueDate:dd/MM/yyyy}</td></tr>
  <tr><td><b>Available Credit</b></td><td class=""amt"">{avail:N2}</td><td><b>Minimum Payment Due</b></td><td class=""amt"">{minDue:N2}</td></tr>
  <tr><td><b>Previous Balance</b></td><td class=""amt"">{FmtSignedAmt(prev)}</td><td><b>Past Due Amount</b></td><td class=""amt"">{overdue:N2}</td></tr>
  <tr><td><b>Payments &amp; Credits</b></td><td class=""amt cr"">{payments:N2} CR</td><td></td><td></td></tr>
  <tr><td><b>Purchases &amp; Cash Advances</b></td><td class=""amt"">{purchases:N2}</td><td></td><td></td></tr>
  <tr><td><b>Charges &amp; Interest</b></td><td class=""amt"">{charges:N2}</td><td></td><td></td></tr>
  <tr class=""total""><td><b>New Balance</b></td><td class=""amt"">{FmtSignedAmt(closing)}</td><td></td><td></td></tr>
</table>
");
        }

        protected virtual void AppendTransactionTable(
            StringBuilder sb, DataRow masterRow,
            DataRow[] detailRows, StatementDataContext ctx, GenerateStatementCommand cmd)
        {
            string currency = Str(masterRow, MAccountCurrency);

            sb.Append(@"
<div class=""sec-hdr"">Transactions</div>
<table>
  <thead>
    <tr>
      <th>Trans Date</th><th>Post Date</th><th>Reference No</th>
      <th>Description</th><th>Foreign Amount</th><th>Amount</th><th>CR/DB</th>
    </tr>
  </thead>
  <tbody>
");
            foreach (DataRow dr in detailRows)
            {
                if (IsOnHold(dr)) continue;

                DateTime transDate = ToDateTime(dr, DTransDate);
                DateTime postDate  = ToDateTime(dr, DPostingDate);
                if (transDate > postDate) transDate = postDate;

                string desc      = Str(dr, DMerchant);
                if (string.IsNullOrEmpty(desc)) desc = Str(dr, DTranDesc);

                string origCurr  = Str(dr, DOrigTranCurrency);
                decimal origAmt  = ToDecimal(dr, DOrigTranAmount);
                decimal billAmt  = ToDecimal(dr, DBillTranAmount);
                string sign      = Str(dr, DBillTranAmountSign);
                bool isCr        = sign == "CR";

                string foreignAmt = (currency != origCurr && !string.IsNullOrWhiteSpace(origCurr) && origCurr != "XXX")
                    ? $"{origAmt:##,###,##0.00} {origCurr}"
                    : string.Empty;

                string crdb      = isCr ? "CR" : string.Empty;
                string amtClass  = isCr ? "amt cr" : "amt";

                sb.Append($@"    <tr>
      <td>{transDate:dd/MM/yyyy}</td>
      <td>{postDate:dd/MM/yyyy}</td>
      <td>{EscHtml(TrimStr(Str(dr, DRefNo), 25))}</td>
      <td>{EscHtml(TrimStr(desc, 40))}</td>
      <td class=""amt"">{EscHtml(foreignAmt)}</td>
      <td class=""{amtClass}"">{billAmt:N2}</td>
      <td>{crdb}</td>
    </tr>
");
            }
            sb.AppendLine("  </tbody>\n</table>");
        }

        protected virtual void AppendCardFooter(
            StringBuilder sb, DataRow masterRow,
            StatementDataContext ctx, GenerateStatementCommand cmd)
        {
            if (cmd.HasMode(ProcessingMode.Reward) && ctx.RewardDataSet != null)
                AppendRewardSection(sb, masterRow, ctx);
        }

        protected virtual void AppendRewardSection(
            StringBuilder sb, DataRow masterRow, StatementDataContext ctx)
        {
            var rt = ctx.RewardDataSet?.Tables[0];
            if (rt == null) return;

            string accountNo = Str(masterRow, AccountNoField);
            var rr = rt.Select($"ACCOUNTNO = '{EscSql(accountNo)}'");
            if (rr.Length == 0) return;

            DataRow r = rr[0];
            sb.Append($@"
<div class=""sec-hdr"">Reward Programme</div>
<table>
  <thead>
    <tr><th>Opening Balance</th><th>Earned Bonus</th><th>Redeemed Bonus</th><th>Expired Bonus</th><th>Closing Balance</th></tr>
  </thead>
  <tbody>
    <tr>
      <td class=""amt"">{SafeDec(r, "openingbonus"):N0}</td>
      <td class=""amt"">{SafeDec(r, "earnedbonus"):N0}</td>
      <td class=""amt"">{SafeDec(r, "redeemedbonus"):N0}</td>
      <td class=""amt"">{SafeDec(r, "expiredbonus"):N0}</td>
      <td class=""amt"">{SafeDec(r, "closingbonus"):N0}</td>
    </tr>
  </tbody>
</table>
");
        }

        protected virtual void AppendHtmlClose(
            StringBuilder sb, StatementDataContext ctx, GenerateStatementCommand cmd)
        {
            string website = EscHtml(cmd.BankWebsiteUrl ?? string.Empty);
            sb.Append($@"
<div class=""ftr"">
  <p>This is a computer-generated statement. No signature is required.</p>
  {(string.IsNullOrEmpty(website) ? "" : $"<p><a href=\"http://{website}\">{website}</a></p>")}
</div>
</body>
</html>
");
        }

        // ── Branding ── override in bank-specific subclasses ──────────────────────

        protected virtual string GetPrimaryColor()   => "#003366";
        protected virtual string GetSecondaryColor() => "#4472C4";

        // ── Data helpers ──────────────────────────────────────────────────────────

        protected static bool IsOnHold(DataRow dr)
            => dr[DPostingDate] == DBNull.Value && dr[DDocNo] == DBNull.Value;

        protected static string Str(DataRow row, string col)
        {
            try { return row.IsNull(col) ? string.Empty : (row[col]?.ToString()?.Trim() ?? string.Empty); }
            catch { return string.Empty; }
        }

        protected static string Str(DataRow row, int idx)
        {
            try { return row.IsNull(idx) ? string.Empty : (row[idx]?.ToString()?.Trim() ?? string.Empty); }
            catch { return string.Empty; }
        }

        protected static decimal ToDecimal(DataRow row, string col)
        {
            try { return row.IsNull(col) ? 0m : Convert.ToDecimal(row[col]); }
            catch { return 0m; }
        }

        protected static decimal SafeDec(DataRow row, string col)
        {
            try { return ToDecimal(row, col); } catch { return 0m; }
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

        protected static string FmtSignedAmt(decimal val)
            => val >= 0 ? $"{val:N2} DB" : $"{Math.Abs(val):N2} CR";

        protected static string TrimStr(string s, int maxLen)
            => s?.Length > maxLen ? s.Substring(0, maxLen) : (s ?? string.Empty);

        protected static string EscHtml(string s)
            => (s ?? string.Empty)
               .Replace("&", "&amp;")
               .Replace("<", "&lt;")
               .Replace(">", "&gt;")
               .Replace("\"", "&quot;");

        protected static string EscSql(string s)
            => (s ?? string.Empty).Replace("'", "''");
    }
}

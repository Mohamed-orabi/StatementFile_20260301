using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using StatementFile.Application.UseCases.StatementGeneration;

namespace StatementFile.Infrastructure.Formatters.RawData
{
    // ── Raw-data formatters (one class per bank) ──────────────────────────────────
    // All extend NativeRawDataFormatterBase — no legacy class references.
    // ─────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// AIBK raw data:
    ///   STMT_HDR.DAT  – comma-separated header records (includes installment,
    ///                   passport/identity data, sort by card priority)
    ///   STMT_DTL.DAT  – comma-separated transaction records
    ///   STMT_Delivery.CSV – delivery manifest (address, phone, citizen ID)
    ///   .MD5 checksum file
    /// </summary>
    public sealed class RawDataAibkFormatter : NativeRawDataFormatterBase
    {
        public override string FormatterKey => "RAWDATA_AIBK";
        protected override string FieldSeparator => ",";

        protected override string GetMainFilePath(string baseName, GenerateStatementCommand cmd)
            => Path.Combine(Path.GetDirectoryName(baseName)!, "STMT_HDR.DAT");

        protected override string GetTransFilePath(string baseName, GenerateStatementCommand cmd)
            => Path.Combine(Path.GetDirectoryName(baseName)!, "STMT_DTL.DAT");

        protected override void WriteMainHeader(StreamWriter sw, GenerateStatementCommand cmd)
        {
            const string s = ",";
            sw.WriteLine(
                $"CardHolder Name{s}Address1{s}Address2{s}Address3{s}Address4{s}" +
                $"Card Type{s}Account{s}Branch{s}Statement Date{s}Card No.{s}" +
                $"Credit Limit{s}Available Credit{s}Min. Payment Due{s}New Balance{s}" +
                $"Payment Due Date{s}Past Due Amount{s}prev.balance{s}payment & CR{s}" +
                $"Purch.Cash&Dr{s}Finance Charge{s}External NO{s}");
        }

        protected override void WriteTransHeader(StreamWriter sw, GenerateStatementCommand cmd)
        {
            const string s = ",";
            sw.WriteLine(
                $"Card No.{s}Date of Trans{s}Date of Post{s}Reference{s}" +
                $"Description{s}Purchase Currency & Amount{s}Amount{s}");
        }
    }

    /// <summary>
    /// EGB raw data (same STMT_HDR.DAT + STMT_DTL.DAT structure as AIBK,
    /// comma-separated, card number formatted as XXXX-XXXX-XXXX-XXXX).
    /// </summary>
    public sealed class RawDataEgbFormatter : NativeRawDataFormatterBase
    {
        public override string FormatterKey => "RAWDATA_EGB";
        protected override string FieldSeparator => ",";

        protected override string GetMainFilePath(string baseName, GenerateStatementCommand cmd)
            => Path.Combine(Path.GetDirectoryName(baseName)!, "STMT_HDR.DAT");

        protected override string GetTransFilePath(string baseName, GenerateStatementCommand cmd)
            => Path.Combine(Path.GetDirectoryName(baseName)!, "STMT_DTL.DAT");
    }

    /// <summary>
    /// AAIB raw data:
    ///   _Main.txt + _Trns.txt (pipe-separated)
    ///   Includes OVERDUEDAYS column (from LoadWithOverdueDays query variant).
    ///   Jira: AAIB-9308, AAIB-12395.
    /// </summary>
    public sealed class RawDataAaibFormatter : NativeRawDataFormatterBase
    {
        public override string FormatterKey => "RAWDATA_AAIB";
        protected override string FieldSeparator => "|";

        protected override void WriteMainHeader(TextWriter sw, GenerateStatementCommand cmd)
        {
            const string s = "|";
            sw.WriteLine(
                $"CardHolder Name{s}Address1{s}Address2{s}Address3{s}Address4{s}" +
                $"Card Type{s}Account{s}Branch{s}Statement Date{s}Card No.{s}" +
                $"Credit Limit{s}Available Credit{s}Min. Payment Due{s}New Balance{s}" +
                $"Payment Due Date{s}Past Due Amount{s}prev.balance{s}payment & CR{s}" +
                $"Purch.Cash&Dr{s}Finance Charge{s}External NO{s}OverDueDays{s}");
        }

        protected override void WriteMainRow(
            TextWriter sw, DataRow mRow, DataRow[] detailRows,
            string primaryCardNo, string accountNo,
            StatementDataContext ctx, GenerateStatementCommand cmd)
        {
            // Write standard row then append OverDueDays at end
            var sb = new StringBuilder();
            using var tmpSw = new StringWriter(sb);
            base.WriteMainRow(tmpSw, mRow, detailRows, primaryCardNo, accountNo, ctx, cmd);
            string line = sb.ToString().TrimEnd('\r', '\n');
            string overdueDays = Str(mRow, "overduedays");
            sw.WriteLine(line + overdueDays + FieldSeparator);
        }
    }

    /// <summary>
    /// AIBK alternate format:
    ///   STM_Header_yyMM.txt + STM_Detail_yyMM.txt
    ///   Uses "yyMM" date format in filenames.
    /// </summary>
    public sealed class RawDataAibkAltFormatter : NativeRawDataFormatterBase
    {
        public override string FormatterKey => "RAWDATA_AIBK_ALT";
        protected override string FieldSeparator => ",";

        protected override string GetMainFilePath(string baseName, GenerateStatementCommand cmd)
        {
            string dir = Path.GetDirectoryName(baseName)!;
            return Path.Combine(dir, $"STM_Header_{cmd.StatementDate:yyMM}.txt");
        }

        protected override string GetTransFilePath(string baseName, GenerateStatementCommand cmd)
        {
            string dir = Path.GetDirectoryName(baseName)!;
            return Path.Combine(dir, $"STM_Detail_{cmd.StatementDate:yyMM}.txt");
        }
    }

    /// <summary>
    /// ALXB retail raw data (VISA cards excluded via LoadExcludingVisa):
    ///   STMT_HDR.DAT + STMT_DTL.DAT + STMT_Delivery.CSV
    ///   Jira: ALXB-5971.
    /// </summary>
    public sealed class RawDataAlxbFormatter : NativeRawDataFormatterBase
    {
        public override string FormatterKey => "RAWDATA_ALXB";
        protected override string FieldSeparator => ",";

        protected override string GetMainFilePath(string baseName, GenerateStatementCommand cmd)
            => Path.Combine(Path.GetDirectoryName(baseName)!, "STMT_HDR.DAT");

        protected override string GetTransFilePath(string baseName, GenerateStatementCommand cmd)
            => Path.Combine(Path.GetDirectoryName(baseName)!, "STMT_DTL.DAT");
    }

    /// <summary>
    /// ALXB corporate raw data (VISA cards excluded):
    ///   STMT_CORP_HDR.DAT + STMT_CORP_DTL.DAT (pipe-separated)
    ///   Dynamic account field: cardaccountno.
    ///   Jira: ALXB-5971.
    /// </summary>
    public sealed class RawDataAlxbCorpFormatter : NativeRawDataFormatterBase
    {
        public override string FormatterKey => "RAWDATA_ALXB_CORP";
        protected override string FieldSeparator => "|";
        protected override string AccountNoField  => MCardAccountNo;

        protected override string GetMainFilePath(string baseName, GenerateStatementCommand cmd)
            => Path.Combine(Path.GetDirectoryName(baseName)!, "STMT_CORP_HDR.DAT");

        protected override string GetTransFilePath(string baseName, GenerateStatementCommand cmd)
            => Path.Combine(Path.GetDirectoryName(baseName)!, "STMT_CORP_DTL.DAT");
    }

    /// <summary>
    /// BRKA raw data (pipe-separated):
    ///   _Main.txt + _Trns.txt
    ///   Includes reward fields: EarnedBonus, RedeemedBonus, ExpiredBonus,
    ///   BonusAdjustment, CreditContracts, OverDueDays, Card Expiry Date.
    /// </summary>
    public sealed class RawDataBrkaFormatter : NativeRawDataFormatterBase
    {
        public override string FormatterKey => "RAWDATA_BRKA";
        protected override string FieldSeparator => "|";

        protected override void WriteMainHeader(TextWriter sw, GenerateStatementCommand cmd)
        {
            const string s = "|";
            sw.WriteLine(
                $"CardHolder Name{s}Address1{s}Address2{s}Address3{s}Address4{s}" +
                $"Card Type{s}Account{s}Branch{s}Statement Date{s}Card No.{s}" +
                $"Credit Limit{s}Available Credit{s}Min. Payment Due{s}New Balance{s}" +
                $"Payment Due Date{s}Past Due Amount{s}prev.balance{s}payment & CR{s}" +
                $"Purch.Cash&Dr{s}Finance Charge{s}External NO{s}" +
                $"EarnedBonus{s}RedeemedBonus{s}ExpiredBonus{s}BonusAdjustment{s}" +
                $"CreditContracts{s}OverDueDays{s}CardExpiryDate{s}");
        }

        protected override void WriteMainRow(
            TextWriter sw, DataRow mRow, DataRow[] detailRows,
            string primaryCardNo, string accountNo,
            StatementDataContext ctx, GenerateStatementCommand cmd)
        {
            var sb = new StringBuilder();
            using var tmpSw = new StringWriter(sb);
            base.WriteMainRow(tmpSw, mRow, detailRows, primaryCardNo, accountNo, ctx, cmd);
            string line = sb.ToString().TrimEnd('\r', '\n');

            string earnedBonus   = "0";
            string redeemedBonus = "0";
            string expiredBonus  = "0";
            string bonusAdj      = "0";
            string creditContr   = "0";
            string overdueDays   = Str(mRow, "overduedays");
            string cardExpiry    = Str(mRow, "cardexpirydate");

            if (ctx.RewardDataSet?.Tables[0] != null)
            {
                var rr = ctx.RewardDataSet.Tables[0].Select($"ACCOUNTNO = '{EscSql(accountNo)}'");
                if (rr.Length > 0)
                {
                    earnedBonus   = Str(rr[0], "earnedbonus");
                    redeemedBonus = Str(rr[0], "redeemedbonus");
                    expiredBonus  = Str(rr[0], "expiredbonus");
                    bonusAdj      = Str(rr[0], "bonusadjustment");
                }
            }

            const string s = "|";
            sw.WriteLine(line + earnedBonus + s + redeemedBonus + s + expiredBonus + s +
                         bonusAdj + s + creditContr + s + overdueDays + s + cardExpiry + s);
        }
    }

    /// <summary>
    /// UNB raw data (pipe-separated):
    ///   _Main.txt + _Trns.txt
    ///   Includes CardPrimary and PrimaryCardNo fields in header.
    /// </summary>
    public sealed class RawDataUnbFormatter : NativeRawDataFormatterBase
    {
        public override string FormatterKey => "RAWDATA_UNB";
        protected override string FieldSeparator => "|";

        protected override void WriteMainHeader(TextWriter sw, GenerateStatementCommand cmd)
        {
            const string s = "|";
            sw.WriteLine(
                $"CardHolder Name{s}Address1{s}Address2{s}Address3{s}Address4{s}" +
                $"Card Type{s}Account{s}Branch{s}Statement Date{s}Card No.{s}" +
                $"Credit Limit{s}Available Credit{s}Min. Payment Due{s}New Balance{s}" +
                $"Payment Due Date{s}Past Due Amount{s}prev.balance{s}payment & CR{s}" +
                $"Purch.Cash&Dr{s}Finance Charge{s}External NO{s}" +
                $"CardPrimary{s}PrimaryCardNo{s}");
        }

        protected override void WriteMainRow(
            TextWriter sw, DataRow mRow, DataRow[] detailRows,
            string primaryCardNo, string accountNo,
            StatementDataContext ctx, GenerateStatementCommand cmd)
        {
            var sb = new StringBuilder();
            using var tmpSw = new StringWriter(sb);
            base.WriteMainRow(tmpSw, mRow, detailRows, primaryCardNo, accountNo, ctx, cmd);
            string line = sb.ToString().TrimEnd('\r', '\n');

            const string s = "|";
            string cardPrimary   = Str(mRow, MCardPrimary);
            string primaryCardNo2 = Str(mRow, MPrimaryCardNo);
            sw.WriteLine(line + cardPrimary + s + primaryCardNo2 + s);
        }
    }

    /// <summary>
    /// VCBK raw data (pipe-separated):
    ///   _Main.txt + _Trns.txt
    ///   Includes Email address and TotalDueAmount in header.
    ///   Transaction date format: "dd/MMM/yyyy".
    ///   Detects MasterCard vs VISA via statement-type field.
    /// </summary>
    public sealed class RawDataVcbkFormatter : NativeRawDataFormatterBase
    {
        public override string FormatterKey  => "RAWDATA_VCBK";
        protected override string FieldSeparator  => "|";
        protected override string TransDateFormat => "dd/MMM/yyyy";

        protected override void WriteMainHeader(TextWriter sw, GenerateStatementCommand cmd)
        {
            const string s = "|";
            sw.WriteLine(
                $"CardHolder Name{s}Address1{s}Address2{s}Address3{s}Address4{s}" +
                $"Card Type{s}Account{s}Branch{s}Statement Date{s}Card No.{s}" +
                $"Credit Limit{s}Available Credit{s}Min. Payment Due{s}New Balance{s}" +
                $"Payment Due Date{s}Past Due Amount{s}prev.balance{s}payment & CR{s}" +
                $"Purch.Cash&Dr{s}Finance Charge{s}External NO{s}" +
                $"Email{s}TotalDueAmount{s}");
        }

        protected override void WriteMainRow(
            TextWriter sw, DataRow mRow, DataRow[] detailRows,
            string primaryCardNo, string accountNo,
            StatementDataContext ctx, GenerateStatementCommand cmd)
        {
            var sb = new StringBuilder();
            using var tmpSw = new StringWriter(sb);
            base.WriteMainRow(tmpSw, mRow, detailRows, primaryCardNo, accountNo, ctx, cmd);
            string line = sb.ToString().TrimEnd('\r', '\n');

            // Email address from EmailDataSet
            string email = string.Empty;
            string clientId = Str(mRow, MClientId);
            if (ctx.EmailDataSet?.Tables["Emails"] != null)
            {
                var er = ctx.EmailDataSet.Tables["Emails"].Select($"idclient = {clientId}");
                if (er.Length > 0) email = Str(er[0], 1);
            }

            string totalDue = Str(mRow, MTotalOverdue);
            const string s  = "|";
            sw.WriteLine(line + email + s + totalDue + s);
        }
    }
}

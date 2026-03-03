using System;
using System.Collections.Generic;
using System.IO;
using StatementFile.Application.UseCases.StatementGeneration;

namespace StatementFile.Infrastructure.Formatters.Html
{
    // ── Helper base for all HTML adapters that call the same Statement() signature ──

    /// <summary>
    /// Shared implementation for all HTML adapters whose legacy class
    /// exposes Statement(pStrFileName, pBankName, pBankCode, pStrFile, pDate,
    ///                   pStmntType, pAppendData) or the 8-argument variant
    ///                   with pReportName.
    /// Each concrete class supplies its FormatterKey and instantiates its
    /// specific legacy class.
    /// </summary>
    internal abstract class HtmlStatAdapter7 : LegacyFormatterAdapterBase
    {
        protected abstract void InvokeLegacy(
            string fileBasePath, string bankName, int branchCode,
            string suffix, DateTime date, bool appendMode);

        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fileBasePath, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            InvokeLegacy(fileBasePath, cmd.BankName, cmd.BranchCode,
                         cmd.StatementTypeSuffix, cmd.StatementDate, cmd.AppendMode);
            return CollectOutputFiles(
                Path.Combine(fileBasePath,
                    cmd.StatementDate.ToString("yyyyMM") + cmd.BankName + "_" + cmd.StatementTypeSuffix),
                startedAt);
        }
    }

    // ──────────────────────────────────────────────────────────────────────────────
    // One concrete adapter per bank HTML class.
    // FormatterKey convention: "HTML_{BankCode}[_{Variant}]"
    // ──────────────────────────────────────────────────────────────────────────────

    /// <summary>Access Bank ABP credit e-statement.</summary>
    public sealed class HtmlAbpAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_ABP";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlABP().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>Access Bank ABP supplementary card variant.</summary>
    public sealed class HtmlAbpSupAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_ABP_SUP";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlABP_Sup().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>AIBK credit/debit e-statement (with National ID &amp; Birth Year).</summary>
    public sealed class HtmlAibkAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_AIBK";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlAIBK().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>AIBK Valu product (credit/prepaid) e-statement.</summary>
    public sealed class HtmlAibkValuAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_AIBK_VALU";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlAIBKValu().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>Alexandria Bank (ALXB) credit e-statement (with social media links).</summary>
    public sealed class HtmlAlxbAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_ALXB";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlALXB().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>Alexandria Bank corporate prepaid e-statement.</summary>
    public sealed class HtmlAlxbCpAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_ALXB_CP";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlALXB_CP().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>BAI Angola credit e-statement (Portuguese language).</summary>
    public sealed class HtmlBaiCreditAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_BAI_CREDIT";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlBAICredit().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>BAI Angola prepaid e-statement (Kamba branding, Kwanza amounts).</summary>
    public sealed class HtmlBaiPrepaidAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_BAI_PREPAID";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlBAIPrepaid().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>BDCA (Bank of Credit and Investment Cairo) credit e-statement.</summary>
    public sealed class HtmlBdcaAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_BDCA";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlBDCA().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>BPC credit e-statement.</summary>
    public sealed class HtmlBpcAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_BPC";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlBPC().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>BPC prepaid e-statement.</summary>
    public sealed class HtmlBpcPrepaidAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_BPC_PREPAID";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlBPCPrepaid().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>CMB (Cairo Commercial Bank) credit e-statement with reward programme.</summary>
    public sealed class HtmlCmbAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_CMB";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlCMB().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>DBN credit e-statement.</summary>
    public sealed class HtmlDbnAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_DBN";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlDBN().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>Fidelity Bank Group (FBPG) credit e-statement.</summary>
    public sealed class HtmlFbpgAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_FBP";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlFBPG().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>GTBank Nigeria (GTBK) credit e-statement.</summary>
    public sealed class HtmlGtbkAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_GTBK";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlGTBK().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>GTBank Nigeria (GTBK) debit e-statement.</summary>
    public sealed class HtmlGtbkDebitAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_GTBK_DB";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlGTBKDB().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>GTBank Nigeria (GTBN) credit e-statement.</summary>
    public sealed class HtmlGtbnAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_GTBN";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlGTBN().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>GTBank Uganda prepaid e-statement.</summary>
    public sealed class HtmlGtbuPrepaidAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_GTBU_PREPAID";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlGTBUPrepaid().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>First Bank Nigeria (FBN) generic credit e-statement.</summary>
    public sealed class HtmlFbnAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_FBN";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlGnrlFBN().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>First Bank Nigeria corporate e-statement.</summary>
    public sealed class HtmlFbnCorpAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_FBN_CORP";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlGnrlFBNCompany().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>First Bank Nigeria debit e-statement.</summary>
    public sealed class HtmlFbnDebitAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_FBN_DB";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlGnrlFBNDebit().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>First Bank Nigeria supplementary card e-statement.</summary>
    public sealed class HtmlFbnSupAdapter : LegacyFormatterAdapterBase
    {
        public override string FormatterKey => "HTML_FBN_SUP";
        protected override IEnumerable<string> FormatCore(
            StatementDataContext ctx, string fp, GenerateStatementCommand cmd)
        {
            var startedAt = DateTime.Now;
            new clsStatHtmlGnrlFBN_Sup().Statement(fp, cmd.BankName, cmd.BranchCode,
                cmd.StatementTypeSuffix, cmd.StatementDate, cmd.StatementTypeSuffix,
                cmd.AppendMode, string.Empty);
            return CollectOutputFiles(
                Path.Combine(fp, cmd.StatementDate.ToString("yyyyMM") + cmd.BankName + "_" + cmd.StatementTypeSuffix),
                startedAt);
        }
    }

    /// <summary>HBLN credit e-statement.</summary>
    public sealed class HtmlHblnAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_HBLN";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlHBLN().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>HBLN prepaid e-statement.</summary>
    public sealed class HtmlHblnPrepaidAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_HBLN_PREPAID";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlHBLNPre().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>ICBG (International Commercial Bank Ghana) credit e-statement.</summary>
    public sealed class HtmlIcbgAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_ICBG";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlICBG().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>IMB prepaid e-statement (generic).</summary>
    public sealed class HtmlImbPrepaidAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_IMB_PREPAID";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlIMBPrepaid().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>NBS credit e-statement.</summary>
    public sealed class HtmlNbsAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_NBS";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlNBS().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>RBGH (Regional Bank Ghana) credit e-statement.</summary>
    public sealed class HtmlRbghAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_RBGH";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlRBGHCredit().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>Stanbic Bank Nigeria (SBN) credit e-statement.</summary>
    public sealed class HtmlSbnAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_SBN";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlSBN().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>Stanbic Bank Nigeria updated template.</summary>
    public sealed class HtmlSbnNewAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_SBN_NEW";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlSBN_New().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>Stanbic Bank Nigeria with digital-signature support.</summary>
    public sealed class HtmlSbnSignatureAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_SBN_SIG";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlSBN_New_Signature().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>Polaris Bank / Skye Bank (SBP) credit e-statement.</summary>
    public sealed class HtmlSbpAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_SBP";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlSBP().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>Polaris Bank debit e-statement.</summary>
    public sealed class HtmlSbpDebitAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_SBP_DB";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlSBPDebit().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>Polaris Bank prepaid e-statement.</summary>
    public sealed class HtmlSbpPrepaidAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_SBP_PREPAID";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlSBPPrepaid().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>Stanbic IBTC Nigeria (SIBN) credit e-statement.</summary>
    public sealed class HtmlSibnAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_SIBN";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlSIBN().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>UBA Ghana prepaid e-statement.</summary>
    public sealed class HtmlUbaGPrepaidAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_UBAG_PREPAID";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlUBAGPrepaid().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>UNBN credit e-statement.</summary>
    public sealed class HtmlUnbnAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_UNBN";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlUNBN().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>Wema Bank (WEMA) credit e-statement.</summary>
    public sealed class HtmlWemaAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_WEMA";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlWEMA().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>Wema Bank debit e-statement.</summary>
    public sealed class HtmlWemaDebitAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_WEMA_DB";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlWEMADebit().Statement(fp, bank, code, sfx, dt, sfx, append);
    }

    /// <summary>AAIB HTML e-statement.</summary>
    public sealed class HtmlAaibAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_AAIB";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmAAIB().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>QNB PDF-format e-statement (Crystal Reports output).</summary>
    public sealed class PdfQnbAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "PDF_QNB";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlQNBPDF().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>BDCA PDF-format e-statement.</summary>
    public sealed class PdfBdcaAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "PDF_BDCA";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlBDCAPDF().Statement(fp, bank, code, sfx, dt, sfx, append, string.Empty);
    }

    /// <summary>IMB generic prepaid e-statement (Gnr01 variant).</summary>
    public sealed class HtmlGnrImbPrepaidAdapter : HtmlStatAdapter7
    {
        public override string FormatterKey => "HTML_GNR_IMB_PREPAID";
        protected override void InvokeLegacy(string fp, string bank, int code,
            string sfx, DateTime dt, bool append)
            => new clsStatHtmlGnrIMBPrepaid().Statement(fp, bank, code, sfx, dt, sfx, append);
    }
}

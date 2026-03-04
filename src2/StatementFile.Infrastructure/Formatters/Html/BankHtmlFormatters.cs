using StatementFile.Application.UseCases.StatementGeneration;
using StatementFile.Domain.Enums;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace StatementFile.Infrastructure.Formatters.Html
{
    // ── HTML e-statement formatters (one class per bank/variant) ─────────────────
    // All extend NativeHtmlFormatterBase — no legacy class references.
    // Bank identity is set via FormatterKey; branding via GetPrimaryColor() overrides.
    // ─────────────────────────────────────────────────────────────────────────────

    /// <summary>UBA – United Bank for Africa credit e-statement.</summary>
    public sealed class HtmlUbaFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_UBA";
        protected override string GetPrimaryColor()   => "#D2232A"; // UBA red
        protected override string GetSecondaryColor() => "#891014";
    }

    /// <summary>Access Bank ABP credit e-statement.</summary>
    public sealed class HtmlAbpFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_ABP";
        protected override string GetPrimaryColor()   => "#E8620B"; // Access Bank orange
        protected override string GetSecondaryColor() => "#B34A00";
    }

    /// <summary>Access Bank ABP supplementary card variant.</summary>
    public sealed class HtmlAbpSupFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_ABP_SUP";
        protected override bool UseCorporateAccountNo => false;
        protected override string GetPrimaryColor()   => "#E8620B";
        protected override string GetSecondaryColor() => "#B34A00";
    }

    /// <summary>AIBK – Arab International Bank Kuwait, credit/debit e-statement.</summary>
    public sealed class HtmlAibkFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_AIBK";
        protected override string GetPrimaryColor()   => "#003A6C";
        protected override string GetSecondaryColor() => "#0066B2";
    }

    /// <summary>AIBK Valu product e-statement.</summary>
    public sealed class HtmlAibkValuFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_AIBK_VALU";
        protected override string GetPrimaryColor()   => "#003A6C";
        protected override string GetSecondaryColor() => "#0066B2";
    }

    /// <summary>ALXB – Alexandria Bank credit e-statement.</summary>
    public sealed class HtmlAlxbFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_ALXB";
        protected override string GetPrimaryColor()   => "#004A8F";
        protected override string GetSecondaryColor() => "#0071BC";
    }

    /// <summary>ALXB corporate prepaid e-statement.</summary>
    public sealed class HtmlAlxbCpFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_ALXB_CP";
        protected override string GetPrimaryColor()   => "#004A8F";
        protected override string GetSecondaryColor() => "#0071BC";
    }

    /// <summary>BAI Angola credit e-statement (Portuguese).</summary>
    public sealed class HtmlBaiCreditFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_BAI_CREDIT";
        protected override string GetPrimaryColor()   => "#C4141B";
        protected override string GetSecondaryColor() => "#8B0000";
    }

    /// <summary>BAI Angola prepaid e-statement.</summary>
    public sealed class HtmlBaiPrepaidFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_BAI_PREPAID";
        protected override string GetPrimaryColor()   => "#C4141B";
        protected override string GetSecondaryColor() => "#8B0000";
    }

    /// <summary>BDCA – Bank of Credit and Investment Cairo credit e-statement.</summary>
    public sealed class HtmlBdcaFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_BDCA";
        protected override string GetPrimaryColor()   => "#00436B";
        protected override string GetSecondaryColor() => "#005F96";
    }

    /// <summary>BPC credit e-statement.</summary>
    public sealed class HtmlBpcFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_BPC";
        protected override string GetPrimaryColor()   => "#CC0000";
        protected override string GetSecondaryColor() => "#900000";
    }

    /// <summary>BPC prepaid e-statement.</summary>
    public sealed class HtmlBpcPrepaidFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_BPC_PREPAID";
        protected override string GetPrimaryColor()   => "#CC0000";
        protected override string GetSecondaryColor() => "#900000";
    }

    /// <summary>CMB – Cairo Commercial Bank credit e-statement with reward.</summary>
    public sealed class HtmlCmbFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_CMB";
        protected override string GetPrimaryColor()   => "#1A5276";
        protected override string GetSecondaryColor() => "#2874A6";
    }

    /// <summary>DBN credit e-statement.</summary>
    public sealed class HtmlDbnFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_DBN";
        protected override string GetPrimaryColor()   => "#006400";
        protected override string GetSecondaryColor() => "#228B22";
    }

    /// <summary>FBPG – Fidelity Bank credit e-statement.</summary>
    public sealed class HtmlFbpgFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_FBP";
        protected override string GetPrimaryColor()   => "#006400";
        protected override string GetSecondaryColor() => "#228B22";
    }

    /// <summary>GTBK – GTBank Nigeria credit e-statement.</summary>
    public sealed class HtmlGtbkFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_GTBK";
        protected override string GetPrimaryColor()   => "#F37021";
        protected override string GetSecondaryColor() => "#C45A10";
    }

    /// <summary>GTBK debit e-statement.</summary>
    public sealed class HtmlGtbkDebitFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_GTBK_DB";
        protected override string GetPrimaryColor()   => "#F37021";
        protected override string GetSecondaryColor() => "#C45A10";
    }

    /// <summary>GTBN – GTBank Nigeria (branch) credit e-statement.</summary>
    public sealed class HtmlGtbnFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_GTBN";
        protected override string GetPrimaryColor()   => "#F37021";
        protected override string GetSecondaryColor() => "#C45A10";
    }

    /// <summary>GTBU Uganda prepaid e-statement.</summary>
    public sealed class HtmlGtbuPrepaidFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_GTBU_PREPAID";
        protected override string GetPrimaryColor()   => "#F37021";
        protected override string GetSecondaryColor() => "#C45A10";
    }

    /// <summary>FBN – First Bank Nigeria credit e-statement.</summary>
    public sealed class HtmlFbnFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_FBN";
        protected override string GetPrimaryColor()   => "#004B87";
        protected override string GetSecondaryColor() => "#0072C6";
    }

    /// <summary>FBN corporate e-statement.</summary>
    public sealed class HtmlFbnCorpFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_FBN_CORP";
        protected override bool UseCorporateAccountNo => true;
        protected override string GetPrimaryColor()   => "#004B87";
        protected override string GetSecondaryColor() => "#0072C6";
    }

    /// <summary>FBN debit e-statement.</summary>
    public sealed class HtmlFbnDebitFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_FBN_DB";
        protected override string GetPrimaryColor()   => "#004B87";
        protected override string GetSecondaryColor() => "#0072C6";
    }

    /// <summary>FBN supplementary card e-statement.</summary>
    public sealed class HtmlFbnSupFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_FBN_SUP";
        protected override string GetPrimaryColor()   => "#004B87";
        protected override string GetSecondaryColor() => "#0072C6";
    }

    /// <summary>HBLN credit e-statement.</summary>
    public sealed class HtmlHblnFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_HBLN";
        protected override string GetPrimaryColor()   => "#003399";
        protected override string GetSecondaryColor() => "#3366CC";
    }

    /// <summary>HBLN prepaid e-statement.</summary>
    public sealed class HtmlHblnPrepaidFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_HBLN_PREPAID";
        protected override string GetPrimaryColor()   => "#003399";
        protected override string GetSecondaryColor() => "#3366CC";
    }

    /// <summary>ICBG – International Commercial Bank Ghana credit e-statement.</summary>
    public sealed class HtmlIcbgFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_ICBG";
        protected override string GetPrimaryColor()   => "#004080";
        protected override string GetSecondaryColor() => "#0059B3";
    }

    /// <summary>IMB prepaid e-statement.</summary>
    public sealed class HtmlImbPrepaidFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_IMB_PREPAID";
        protected override string GetPrimaryColor()   => "#006633";
        protected override string GetSecondaryColor() => "#009944";
    }

    /// <summary>NBS credit e-statement.</summary>
    public sealed class HtmlNbsFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_NBS";
        protected override string GetPrimaryColor()   => "#800000";
        protected override string GetSecondaryColor() => "#B22222";
    }

    /// <summary>RBGH – Regional Bank Ghana credit e-statement.</summary>
    public sealed class HtmlRbghFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_RBGH";
        protected override string GetPrimaryColor()   => "#1F5C99";
        protected override string GetSecondaryColor() => "#2E8BC0";
    }

    /// <summary>SBN – Stanbic Bank Nigeria credit e-statement.</summary>
    public sealed class HtmlSbnFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_SBN";
        protected override string GetPrimaryColor()   => "#009FE3";
        protected override string GetSecondaryColor() => "#0077AA";
    }

    /// <summary>SBN updated template.</summary>
    public sealed class HtmlSbnNewFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_SBN_NEW";
        protected override string GetPrimaryColor()   => "#009FE3";
        protected override string GetSecondaryColor() => "#0077AA";
    }

    /// <summary>SBN with digital signature.</summary>
    public sealed class HtmlSbnSignatureFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_SBN_SIG";
        protected override string GetPrimaryColor()   => "#009FE3";
        protected override string GetSecondaryColor() => "#0077AA";
    }

    /// <summary>SBP – Polaris/Skye Bank credit e-statement.</summary>
    public sealed class HtmlSbpFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_SBP";
        protected override string GetPrimaryColor()   => "#761B8D";
        protected override string GetSecondaryColor() => "#5A0E6D";
    }

    /// <summary>SBP debit e-statement.</summary>
    public sealed class HtmlSbpDebitFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_SBP_DB";
        protected override string GetPrimaryColor()   => "#761B8D";
        protected override string GetSecondaryColor() => "#5A0E6D";
    }

    /// <summary>SBP prepaid e-statement.</summary>
    public sealed class HtmlSbpPrepaidFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_SBP_PREPAID";
        protected override string GetPrimaryColor()   => "#761B8D";
        protected override string GetSecondaryColor() => "#5A0E6D";
    }

    /// <summary>SIBN – Stanbic IBTC Nigeria credit e-statement.</summary>
    public sealed class HtmlSibnFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_SIBN";
        protected override string GetPrimaryColor()   => "#009FE3";
        protected override string GetSecondaryColor() => "#0077AA";
    }

    /// <summary>UBAG – UBA Ghana prepaid e-statement.</summary>
    public sealed class HtmlUbaGPrepaidFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_UBAG_PREPAID";
        protected override string GetPrimaryColor()   => "#D2232A";
        protected override string GetSecondaryColor() => "#891014";
    }

    /// <summary>UNBN credit e-statement.</summary>
    public sealed class HtmlUnbnFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_UNBN";
        protected override string GetPrimaryColor()   => "#003366";
        protected override string GetSecondaryColor() => "#004C99";
    }

    /// <summary>WEMA Bank credit e-statement.</summary>
    public sealed class HtmlWemaFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_WEMA";
        protected override string GetPrimaryColor()   => "#6A0DAD";
        protected override string GetSecondaryColor() => "#4B0082";
    }

    /// <summary>WEMA Bank debit e-statement.</summary>
    public sealed class HtmlWemaDebitFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_WEMA_DB";
        protected override string GetPrimaryColor()   => "#6A0DAD";
        protected override string GetSecondaryColor() => "#4B0082";
    }

    /// <summary>AAIB HTML e-statement.</summary>
    public sealed class HtmlAaibFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_AAIB";
        protected override string GetPrimaryColor()   => "#005580";
        protected override string GetSecondaryColor() => "#007ACC";
    }

    /// <summary>IMB generic prepaid e-statement (Gnr01 variant).</summary>
    public sealed class HtmlGnrImbPrepaidFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "HTML_GNR_IMB_PREPAID";
        protected override string GetPrimaryColor()   => "#006633";
        protected override string GetSecondaryColor() => "#009944";
    }

    // ── PDF-format formatters ─────────────────────────────────────────────────────
    // QNB and BDCA were previously Crystal Reports PDF formatters.
    // The native implementation produces a structured HTML file; PDF conversion
    // is handled downstream by the OS print pipeline or a dedicated PDF service.

    /// <summary>QNB PDF-format e-statement (rendered as structured HTML).</summary>
    public sealed class PdfQnbFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "PDF_QNB";
        protected override string GetPrimaryColor()   => "#6B1E1E";
        protected override string GetSecondaryColor() => "#8B2525";
    }

    /// <summary>BDCA PDF-format e-statement (rendered as structured HTML).</summary>
    public sealed class PdfBdcaFormatter : NativeHtmlFormatterBase
    {
        public override string FormatterKey  => "PDF_BDCA";
        protected override string GetPrimaryColor()   => "#00436B";
        protected override string GetSecondaryColor() => "#005F96";
    }
}

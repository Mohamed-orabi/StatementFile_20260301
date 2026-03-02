namespace StatementFile.Domain.Enums
{
    /// <summary>
    /// The physical file format produced by a statement generation run.
    /// Each bank product is bound to exactly one output type.
    /// All types are represented in the Banks/ folder of the legacy codebase.
    /// </summary>
    public enum StatementOutputType
    {
        /// <summary>
        /// HTML e-statement sent as email body + optional PDF attachment.
        /// Base class: clsStatHtml / clsBasStatement.
        /// Banks: UBA, ABP, BAI, GTBK, WEMA, SBP, FBN, FCMB, HBLN, ICBG, ALXB, AIBK,
        ///        CMB, DBN, NBS, SIBN, BDCA, UNBN, SBN, RBGH, IMB, GTBU, FBP, BPC, AAIB, etc.
        /// Produces: {bankName}_{type}_Emails.txt + WithoutEmails.txt + per-card HTML/PDF.
        /// </summary>
        Html = 0,

        /// <summary>
        /// Fixed-width binary DAT / CSV files packaged as ZIP + MD5.
        /// Base class: clsStatRawDataAIBK / clsBasStatement.
        /// Banks: AIBK, EGB, ALXB (retail), ALXB (corporate), AAIB, BRKA, UNB, VCBK.
        /// Produces: STMT_HDR.DAT + STMT_DTL.DAT + STMT_Delivery.CSV + .zip + .MD5.
        /// </summary>
        RawData = 1,

        /// <summary>
        /// Fixed-width text for a physical laser printer with page-flag markers.
        /// Page flags: F 0 (single), F 1 (first), F 2 (middle), F 3 (last).
        /// Base class: clsStatTxt.
        /// Banks: FCMB, Suez, FBN (Debit label), AIB (Debit), BCA (Debit),
        ///        ICBG (Debit), EDBE (label), FABG.
        /// Produces: single .txt file with control characters.
        /// </summary>
        TextLabel = 2,

        /// <summary>
        /// Plain fixed-format text with form-feed (\u000C) and carriage-return (\u000D)
        /// control characters for physical printing.
        /// Base class: clsBasStatement.
        /// Banks: EDBE.
        /// </summary>
        Text = 3,

        /// <summary>
        /// DataSet serialised to XML (WriteXml with WriteSchema) + ZIP + MD5.
        /// Base class: clsBasStatement.
        /// Banks: IDBE (VIP cards, branch 16).
        /// </summary>
        Xml = 4,

        /// <summary>
        /// Crystal Reports PDF export or PdfSharp-encrypted PDF.
        /// Base class: clsBasStatement.
        /// Banks: QNB (Crystal Reports PDF), BDCA PDF variant, UBA password-protected.
        /// Produces: .pdf file (optionally encrypted).
        /// </summary>
        Pdf = 5,
    }
}

using System;

namespace StatementFile.Domain.Enums
{
    /// <summary>
    /// Flags that control which optional processing branches are active for a statement run.
    /// Multiple modes can be combined with bitwise OR.
    /// The active set is set per bank-product configuration entry.
    /// </summary>
    [Flags]
    public enum ProcessingMode
    {
        /// <summary>Standard credit/debit statement, no special processing.</summary>
        Standard = 0,

        /// <summary>
        /// Include the Reward Programme section (EarnedBonus, RedeemedBonus, ExpiredBonus,
        /// BonusAdjustment, ExpiredBonusNextMonth).
        /// Oracle condition: contracttype = 'New Reward Contract' (configurable per bank).
        /// Legacy: isRewardVal flag in clsStatHtml.
        /// </summary>
        Reward = 1 << 0,

        /// <summary>
        /// Corporate/company-level statement: one summary section per company,
        /// then individual card sub-sections.
        /// Legacy: createCorporateVal + company DataRelation.
        /// </summary>
        Corporate = 1 << 1,

        /// <summary>
        /// Supplementary card handling: primary card is the aggregate root,
        /// supplementary cards are nested children.
        /// Legacy: isPrimaryOnly / totCrdNoInAcc logic.
        /// </summary>
        Supplementary = 1 << 2,

        /// <summary>
        /// VIP filtering: only cards with cardvip = 'Y' or a bank-specific VIP flag.
        /// Legacy: FillStatementDataSet(pBankCode, "vip") call in clsStatXML_IDBE.
        /// </summary>
        Vip = 1 << 3,

        /// <summary>
        /// Installment programme: detect and parse installment transactions
        /// (Cash Installment, Balance Transfer Installment, Purchase Installment,
        /// Early Settlement).
        /// Legacy: isInstallmentVal, getInstallment() calls.
        /// </summary>
        Installment = 1 << 4,

        /// <summary>
        /// Collect external identity data: National ID, Passport Number, Birth Year.
        /// Passed to getClientPassportNo() and loaded into DSPassportNos / DSBirthYandID.
        /// Legacy: clsStatRawDataAIBK, clsStatHtmlAIBK, clsStatHtmlBDCA.
        /// </summary>
        ExternalId = 1 << 5,

        /// <summary>
        /// Include OverDueDays field using the specialised query
        /// FillStatementDataSet_WithOverDueDays().
        /// Legacy: clsStatRawData_AAIB (Jira AAIB-9308).
        /// </summary>
        WithOverdueDays = 1 << 6,

        /// <summary>
        /// Exclude VISA-branded cards (use FillStatementDataSet_Exclude_VisaCards()).
        /// Legacy: clsStatRawDataCorp_ALXB, clsStatRawData_ALXB.
        /// </summary>
        ExcludeVisa = 1 << 7,

        /// <summary>
        /// Sort cards by priority within the master DataSet
        /// (FillStatementDataSet_SortCardPriority).
        /// Legacy: clsStatRawDataAIBK, clsStatRawData_AIBK.
        /// </summary>
        SortCardPriority = 1 << 8,

        /// <summary>
        /// Output an OMR (Optical Mark Recognition) sidecar file for physical sorting.
        /// Legacy: clsOMR integration in clsStatementAAIB.
        /// </summary>
        Omr = 1 << 9,

        /// <summary>
        /// Password-protect the output PDF using PdfSharp encryption.
        /// Password derived from customer data (e.g. date of birth, last 4 digits of card).
        /// Legacy: PdfSharp.Pdf.Security usage in clsStatHtmlUBA / clsStatHtmlABP.
        /// </summary>
        PasswordProtectedPdf = 1 << 10,

        /// <summary>
        /// Apply Arabic address fixing before rendering.
        /// Detects "???" corruption prefix and strips it.
        /// Legacy: clsMaintainData.fixArbicAddress().
        /// </summary>
        FixArabicAddress = 1 << 11,

        /// <summary>
        /// Delete on-hold transaction rows (HOLSTMT = 'Y') before generating.
        /// Legacy: deleteOnHoldTrans() in clsMaintainData.
        /// </summary>
        DeleteOnHold = 1 << 12,

        /// <summary>
        /// Prepaid card mode: shows Available Balance instead of Credit Limit.
        /// Legacy: isPrepaidVal in clsStatHtml.
        /// </summary>
        Prepaid = 1 << 13,

        /// <summary>
        /// Include an Excel export alongside the primary output format.
        /// Legacy: clsStatExcel rawExcel in clsStatementAUB.
        /// </summary>
        ExcelExport = 1 << 14,
    }
}

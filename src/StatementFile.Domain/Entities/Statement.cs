using System;

namespace StatementFile.Domain.Entities
{
    /// <summary>
    /// Represents a bank card statement master record sourced from TSTATEMENTMASTERTABLE.
    /// This entity is the aggregate root for a customer's billing period.
    /// </summary>
    public class Statement
    {
        // --- Identity & Routing ---
        public int    Branch             { get; private set; }
        public string StatementNo        { get; private set; }
        public string StatementNumber    { get; private set; }
        public string CardNo             { get; private set; }
        public string Mbr                { get; private set; }
        public string CardProduct        { get; private set; }
        public bool   CardPrimary        { get; private set; }
        public string PrimaryCardNo      { get; private set; }
        public string AccountNo          { get; private set; }
        public string ExternalNo         { get; private set; }
        public string AccountType        { get; private set; }
        public string AccountStatus      { get; private set; }
        public string AccountCurrency    { get; private set; }

        // --- Statement Period ---
        public DateTime StatementDateFrom { get; private set; }
        public DateTime StatementDateTo   { get; private set; }
        public DateTime StatementDueDate  { get; private set; }

        // --- Card State ---
        public string CardState          { get; private set; }
        public string CardStatus         { get; private set; }
        public string CardExpiryDate     { get; private set; }
        public string CardVip            { get; private set; }
        public string CardPaymentMethod  { get; private set; }
        public string HolStmt            { get; private set; }
        public string Barcode            { get; private set; }

        // --- Customer & Address ---
        public string CustomerName       { get; private set; }
        public string CustomerTitle      { get; private set; }
        public string CustomerAddress1   { get; private set; }
        public string CustomerAddress2   { get; private set; }
        public string CustomerAddress3   { get; private set; }
        public string CustomerRegion     { get; private set; }
        public string CustomerCity       { get; private set; }
        public string CustomerZipCode    { get; private set; }
        public string CustomerCountry    { get; private set; }

        // --- Card Address ---
        public string CardAddress1       { get; private set; }
        public string CardAddress2       { get; private set; }
        public string CardAddress3       { get; private set; }
        public string CardAddressBarcode { get; private set; }
        public string CardAccountNo      { get; private set; }
        public string CardClientId       { get; private set; }
        public string CardClientName     { get; private set; }
        public string CardTitle          { get; private set; }
        public string CardBranchPart     { get; private set; }
        public string CardBranchPartName { get; private set; }

        // --- Financial Limits ---
        public decimal AccountLim          { get; private set; }
        public decimal AccountAvailableLim { get; private set; }
        public decimal CardLimit           { get; private set; }
        public decimal CardAvailableLimit  { get; private set; }
        public decimal ContractLimit       { get; private set; }

        // --- Due & Balance Amounts ---
        public decimal TotalOverdueAmount  { get; private set; }
        public decimal TotalDueAmount      { get; private set; }
        public decimal MinDueAmount        { get; private set; }
        public decimal TotalDebits         { get; private set; }
        public decimal TotalCredits        { get; private set; }
        public decimal TotalPayments       { get; private set; }
        public decimal TotalPurchases      { get; private set; }
        public decimal TotalCashWithdrawal { get; private set; }
        public decimal TotalCharges        { get; private set; }
        public decimal TotalInterest       { get; private set; }
        public decimal OpeningBalance      { get; private set; }
        public decimal ClosingBalance      { get; private set; }
        public decimal CardDafAmount       { get; private set; }
        public decimal DafPercentage       { get; private set; }

        // --- Contract ---
        public string  ContractType        { get; private set; }
        public string  ContractNo          { get; private set; }
        public string  ContractState       { get; private set; }
        public string  CreditContracts     { get; private set; }
        public int     ClientId            { get; private set; }
        public string  ContactPersonName   { get; private set; }
        public string  Dept                { get; private set; }

        // --- Statement Messages ---
        public string StatementMessageLine1 { get; private set; }
        public string StatementMessageLine2 { get; private set; }
        public string StatementMessageLine3 { get; private set; }

        // --- Reward / Bonus ---
        public decimal EarnedBonus           { get; private set; }
        public decimal RedeemedBonus         { get; private set; }
        public decimal ExpiredBonus          { get; private set; }
        public decimal BonusAdjustment       { get; private set; }
        public decimal ExpiredBonusNextMonth { get; private set; }
        public DateTime? ExpiredBonusDate    { get; private set; }

        // --- Installment ---
        public decimal IntSpentLimit         { get; private set; }
        public decimal MinPayPercentage      { get; private set; }
        public int     MonthsCount           { get; private set; }
        public string  PackageName           { get; private set; }
        public decimal InstallmentUsedLimit  { get; private set; }

        // --- Misc ---
        public int     OverdueDays           { get; private set; }
        public string  UserActField1         { get; private set; }
        public string  AccountZipCode        { get; private set; }

        // Private constructor enforces creation through factory method.
        private Statement() { }

        public static Statement Create(
            int branch, string statementNo, string statementNumber,
            string cardNo, string mbr, string cardProduct,
            bool cardPrimary, string primaryCardNo, string accountNo,
            string externalNo, string accountType, string accountStatus,
            string accountCurrency, DateTime statementDateFrom, DateTime statementDateTo,
            DateTime statementDueDate, string cardState, string cardStatus,
            string cardExpiryDate, string cardVip, string cardPaymentMethod,
            string holStmt, string barcode,
            string customerName, string customerTitle,
            string customerAddress1, string customerAddress2, string customerAddress3,
            string customerRegion, string customerCity, string customerZipCode, string customerCountry,
            string cardAddress1, string cardAddress2, string cardAddress3,
            string cardAddressBarcode, string cardAccountNo, string cardClientId,
            string cardClientName, string cardTitle,
            string cardBranchPart, string cardBranchPartName,
            decimal accountLim, decimal accountAvailableLim,
            decimal cardLimit, decimal cardAvailableLimit, decimal contractLimit,
            decimal totalOverdueAmount, decimal totalDueAmount, decimal minDueAmount,
            decimal totalDebits, decimal totalCredits, decimal totalPayments,
            decimal totalPurchases, decimal totalCashWithdrawal, decimal totalCharges,
            decimal totalInterest, decimal openingBalance, decimal closingBalance,
            decimal cardDafAmount, decimal dafPercentage,
            string contractType, string contractNo, string contractState, string creditContracts,
            int clientId, string contactPersonName, string dept,
            string statementMessageLine1, string statementMessageLine2, string statementMessageLine3,
            decimal earnedBonus, decimal redeemedBonus, decimal expiredBonus,
            decimal bonusAdjustment, decimal expiredBonusNextMonth, DateTime? expiredBonusDate,
            decimal intSpentLimit, decimal minPayPercentage, int monthsCount,
            string packageName, decimal installmentUsedLimit,
            int overdueDays, string userActField1, string accountZipCode)
        {
            return new Statement
            {
                Branch = branch,
                StatementNo = statementNo,
                StatementNumber = statementNumber,
                CardNo = cardNo,
                Mbr = mbr,
                CardProduct = cardProduct,
                CardPrimary = cardPrimary,
                PrimaryCardNo = primaryCardNo,
                AccountNo = accountNo,
                ExternalNo = externalNo,
                AccountType = accountType,
                AccountStatus = accountStatus,
                AccountCurrency = accountCurrency,
                StatementDateFrom = statementDateFrom,
                StatementDateTo = statementDateTo,
                StatementDueDate = statementDueDate,
                CardState = cardState,
                CardStatus = cardStatus,
                CardExpiryDate = cardExpiryDate,
                CardVip = cardVip,
                CardPaymentMethod = cardPaymentMethod,
                HolStmt = holStmt,
                Barcode = barcode,
                CustomerName = customerName,
                CustomerTitle = customerTitle,
                CustomerAddress1 = customerAddress1,
                CustomerAddress2 = customerAddress2,
                CustomerAddress3 = customerAddress3,
                CustomerRegion = customerRegion,
                CustomerCity = customerCity,
                CustomerZipCode = customerZipCode,
                CustomerCountry = customerCountry,
                CardAddress1 = cardAddress1,
                CardAddress2 = cardAddress2,
                CardAddress3 = cardAddress3,
                CardAddressBarcode = cardAddressBarcode,
                CardAccountNo = cardAccountNo,
                CardClientId = cardClientId,
                CardClientName = cardClientName,
                CardTitle = cardTitle,
                CardBranchPart = cardBranchPart,
                CardBranchPartName = cardBranchPartName,
                AccountLim = accountLim,
                AccountAvailableLim = accountAvailableLim,
                CardLimit = cardLimit,
                CardAvailableLimit = cardAvailableLimit,
                ContractLimit = contractLimit,
                TotalOverdueAmount = totalOverdueAmount,
                TotalDueAmount = totalDueAmount,
                MinDueAmount = minDueAmount,
                TotalDebits = totalDebits,
                TotalCredits = totalCredits,
                TotalPayments = totalPayments,
                TotalPurchases = totalPurchases,
                TotalCashWithdrawal = totalCashWithdrawal,
                TotalCharges = totalCharges,
                TotalInterest = totalInterest,
                OpeningBalance = openingBalance,
                ClosingBalance = closingBalance,
                CardDafAmount = cardDafAmount,
                DafPercentage = dafPercentage,
                ContractType = contractType,
                ContractNo = contractNo,
                ContractState = contractState,
                CreditContracts = creditContracts,
                ClientId = clientId,
                ContactPersonName = contactPersonName,
                Dept = dept,
                StatementMessageLine1 = statementMessageLine1,
                StatementMessageLine2 = statementMessageLine2,
                StatementMessageLine3 = statementMessageLine3,
                EarnedBonus = earnedBonus,
                RedeemedBonus = redeemedBonus,
                ExpiredBonus = expiredBonus,
                BonusAdjustment = bonusAdjustment,
                ExpiredBonusNextMonth = expiredBonusNextMonth,
                ExpiredBonusDate = expiredBonusDate,
                IntSpentLimit = intSpentLimit,
                MinPayPercentage = minPayPercentage,
                MonthsCount = monthsCount,
                PackageName = packageName,
                InstallmentUsedLimit = installmentUsedLimit,
                OverdueDays = overdueDays,
                UserActField1 = userActField1,
                AccountZipCode = accountZipCode,
            };
        }
    }
}

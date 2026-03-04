using System;

namespace StatementFile.Application.DTOs
{
    /// <summary>
    /// Read-only projection of a Statement entity used by Use Cases and the UI.
    /// </summary>
    public sealed class StatementDto
    {
        public int     Branch             { get; set; }
        public string  StatementNo        { get; set; }
        public string  StatementNumber    { get; set; }
        public string  CardNo             { get; set; }
        public string  Mbr                { get; set; }
        public string  CardProduct        { get; set; }
        public bool    CardPrimary        { get; set; }
        public string  PrimaryCardNo      { get; set; }
        public string  AccountNo          { get; set; }
        public string  ExternalNo         { get; set; }
        public string  AccountType        { get; set; }
        public string  AccountStatus      { get; set; }
        public string  AccountCurrency    { get; set; }

        public DateTime StatementDateFrom { get; set; }
        public DateTime StatementDateTo   { get; set; }
        public DateTime StatementDueDate  { get; set; }

        public string  CardState          { get; set; }
        public string  CardStatus         { get; set; }
        public string  CardExpiryDate     { get; set; }
        public string  CardVip            { get; set; }
        public string  CardPaymentMethod  { get; set; }

        public string  CustomerName       { get; set; }
        public string  CustomerTitle      { get; set; }
        public string  CustomerAddress1   { get; set; }
        public string  CustomerAddress2   { get; set; }
        public string  CustomerAddress3   { get; set; }
        public string  CustomerCity       { get; set; }
        public string  CustomerCountry    { get; set; }

        public decimal AccountLim          { get; set; }
        public decimal AccountAvailableLim { get; set; }
        public decimal CardLimit           { get; set; }
        public decimal CardAvailableLimit  { get; set; }

        public decimal TotalOverdueAmount  { get; set; }
        public decimal TotalDueAmount      { get; set; }
        public decimal MinDueAmount        { get; set; }
        public decimal OpeningBalance      { get; set; }
        public decimal ClosingBalance      { get; set; }

        public string  ContractType        { get; set; }
        public string  ContractNo          { get; set; }

        public decimal EarnedBonus         { get; set; }
        public decimal RedeemedBonus       { get; set; }
        public decimal ExpiredBonus        { get; set; }

        public string  StatementMessageLine1 { get; set; }
        public string  StatementMessageLine2 { get; set; }
        public string  StatementMessageLine3 { get; set; }
    }
}

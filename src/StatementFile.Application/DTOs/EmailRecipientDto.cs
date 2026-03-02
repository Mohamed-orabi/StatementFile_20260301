namespace StatementFile.Application.DTOs
{
    /// <summary>
    /// Carries a resolved email recipient for a specific bank/statement run,
    /// derived from mailConfiguration.json.
    /// </summary>
    public sealed class EmailRecipientDto
    {
        public string Email      { get; set; }
        public string Name       { get; set; }
        public bool?  Valid      { get; set; }  // null = always include; true/false = bank-specific filtering
        public string RecipientType { get; set; } // "To", "CC", "BCC"
    }
}

using System;

namespace StatementFile.Application.DTOs
{
    /// <summary>
    /// API request body for POST /api/statements/generate.
    /// Identifies which bank-product configuration to run and on which date.
    /// </summary>
    public sealed class GenerateStatementRequest
    {
        /// <summary>ID of the BankProductConfig row to use.</summary>
        public int      ConfigId        { get; set; }

        /// <summary>The billing cycle / statement date to generate for.</summary>
        public DateTime StatementDate   { get; set; }

        /// <summary>
        /// Override output directory.
        /// When null/empty the API uses the server-configured default path.
        /// </summary>
        public string   OutputDirectory { get; set; }

        /// <summary>
        /// Override the recipient email address (useful for testing).
        /// When null the address stored in the config row is used.
        /// </summary>
        public string   EmailOverride   { get; set; }

        /// <summary>Append data to existing files instead of overwriting.</summary>
        public bool     AppendData      { get; set; }
    }

    /// <summary>
    /// Request body for POST /api/statements/generate-bulk.
    /// Wraps multiple generation requests that are executed in sequence.
    /// </summary>
    public sealed class GenerateBulkRequest
    {
        public GenerateStatementRequest[] Items { get; set; }
    }
}

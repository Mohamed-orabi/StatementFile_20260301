using System;

namespace StatementFile.Application.UseCases.StatementGeneration
{
    /// <summary>
    /// Request body for POST /api/statement-generation/run.
    /// Identifies which bank/product configuration to run and the statement date.
    /// The API loads the full configuration from the database using ConfigId.
    /// </summary>
    public sealed class StatementRunRequest
    {
        public int      ConfigId      { get; set; }
        public DateTime StatementDate { get; set; }
    }
}

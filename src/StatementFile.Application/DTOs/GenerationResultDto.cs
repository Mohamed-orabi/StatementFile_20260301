namespace StatementFile.Application.DTOs
{
    /// <summary>
    /// API response returned by the statement generation and merchant endpoints.
    /// Carries the key metrics from <c>GenerateStatementResult</c> together with
    /// a simple success/error structure so clients don't need to inspect HTTP
    /// status codes for business-logic failures.
    /// </summary>
    public sealed class GenerationResultDto
    {
        public bool   Success          { get; set; }
        public string Error            { get; set; }

        public int    FilesGenerated   { get; set; }
        public int    EmailsSent       { get; set; }
        public int    StatementsCount  { get; set; }
        public int    TransactionCount { get; set; }
        public string OutputDirectory  { get; set; }
    }
}

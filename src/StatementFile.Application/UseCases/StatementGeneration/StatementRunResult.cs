namespace StatementFile.Application.UseCases.StatementGeneration
{
    /// <summary>
    /// Response body for POST /api/statement-generation/run.
    /// Mirrors <see cref="GenerateStatementResult"/> plus a Success/Error envelope.
    /// </summary>
    public sealed class StatementRunResult
    {
        public bool   Success         { get; set; }
        public string Error           { get; set; }
        public string OutputDirectory { get; set; }
        public int    FilesGenerated  { get; set; }
        public int    EmailsSent      { get; set; }
        public int    NoEmailCount    { get; set; }
        public int    StatementsCount { get; set; }
        public int    TransactionCount { get; set; }
    }
}

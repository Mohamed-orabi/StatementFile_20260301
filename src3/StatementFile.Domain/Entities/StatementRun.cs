using System;

namespace StatementFile.Domain.Entities
{
    /// <summary>
    /// Audit record written to STAT_STATEMENT_RUN for every generation attempt.
    /// Tracks what ran, when, and whether it succeeded.
    /// </summary>
    public sealed class StatementRun
    {
        public int      Id              { get; private set; }
        public int      ConfigId        { get; private set; }
        public DateTime StatementDate   { get; private set; }
        public DateTime StartedAt       { get; private set; }
        public DateTime? FinishedAt     { get; private set; }
        public bool     IsSuccess       { get; private set; }
        public string   ErrorMessage    { get; private set; }
        public int      FilesGenerated  { get; private set; }
        public int      EmailsSent      { get; private set; }
        public int      StatementsCount { get; private set; }
        public string   OutputDirectory { get; private set; }

        private StatementRun() { }

        public static StatementRun Start(int configId, DateTime statementDate, string outputDirectory)
        {
            return new StatementRun
            {
                ConfigId        = configId,
                StatementDate   = statementDate,
                StartedAt       = DateTime.UtcNow,
                OutputDirectory = outputDirectory,
            };
        }

        public void Complete(int filesGenerated, int emailsSent, int statementsCount)
        {
            IsSuccess       = true;
            FilesGenerated  = filesGenerated;
            EmailsSent      = emailsSent;
            StatementsCount = statementsCount;
            FinishedAt      = DateTime.UtcNow;
        }

        public void Fail(string errorMessage)
        {
            IsSuccess    = false;
            ErrorMessage = errorMessage;
            FinishedAt   = DateTime.UtcNow;
        }
    }
}

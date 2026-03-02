namespace StatementFile.Domain.Enums
{
    /// <summary>
    /// Tracks the lifecycle state of a statement generation job.
    /// </summary>
    public enum ProcessingStatus
    {
        Pending    = 0,
        Running    = 1,
        Completed  = 2,
        Failed     = 3,
        Skipped    = 4,
    }
}

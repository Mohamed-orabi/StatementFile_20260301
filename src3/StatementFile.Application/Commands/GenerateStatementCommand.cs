using System;
using StatementFile.Domain.Entities;

namespace StatementFile.Application.Commands
{
    /// <summary>
    /// Internal command object passed to <see cref="GenerateStatementHandler"/>.
    /// Carries the resolved config entity plus runtime parameters.
    /// </summary>
    public sealed class GenerateStatementCommand
    {
        public BankProductConfig Config          { get; init; }
        public DateTime          StatementDate   { get; init; }
        public string            OutputDirectory { get; init; }
        public string            EmailOverride   { get; init; }
        public bool              AppendData      { get; init; }
        public string            ConnectionString{ get; init; }
        public string            ReportTemplate  { get; init; }
    }
}

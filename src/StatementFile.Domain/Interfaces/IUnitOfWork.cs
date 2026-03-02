using System;
using StatementFile.Domain.Interfaces.Repositories;

namespace StatementFile.Domain.Interfaces
{
    /// <summary>
    /// Coordinates all Oracle repository operations within a single connection/transaction scope.
    /// </summary>
    public interface IUnitOfWork : IDisposable
    {
        IStatementRepository     Statements  { get; }
        IBankRepository          Banks       { get; }

        /// <summary>Commits the current Oracle transaction.</summary>
        void Commit();

        /// <summary>Rolls back the current Oracle transaction.</summary>
        void Rollback();
    }
}

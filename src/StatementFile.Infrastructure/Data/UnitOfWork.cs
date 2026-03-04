using System;
using Microsoft.Data.SqlClient;
using StatementFile.Application.Interfaces;
using StatementFile.Domain.Interfaces;
using StatementFile.Domain.Interfaces.Repositories;
using StatementFile.Infrastructure.Data.Repositories;

namespace StatementFile.Infrastructure.Data
{
    /// <summary>
    /// SQL Server–backed implementation of <see cref="IUnitOfWork"/>.
    /// Shares a single SqlConnection (and optional transaction) across all
    /// repositories created in this scope, then disposes the connection on Dispose().
    /// </summary>
    public sealed class UnitOfWork : IUnitOfWork
    {
        private readonly SqlConnection      _conn;
        private          SqlTransaction     _transaction;
        private          bool               _disposed;

        private IStatementRepository _statementRepo;
        private IBankRepository      _bankRepo;

        private readonly IConfigurationService _config;
        private readonly SessionContext        _session;

        public UnitOfWork(SqlConnectionFactory factory, IConfigurationService config, SessionContext session)
        {
            _config  = config   ?? throw new ArgumentNullException(nameof(config));
            _session = session  ?? throw new ArgumentNullException(nameof(session));
            _conn    = factory.CreateOpenConnection();
        }

        public IStatementRepository Statements =>
            _statementRepo ??= new StatementRepository(_conn, _config, _session);

        public IBankRepository Banks =>
            _bankRepo ??= new BankRepository(_conn, _config);

        public void Commit()
        {
            _transaction?.Commit();
        }

        public void Rollback()
        {
            _transaction?.Rollback();
        }

        public void Dispose()
        {
            if (_disposed) return;
            _transaction?.Dispose();
            _conn?.Dispose();
            _disposed = true;
        }
    }
}

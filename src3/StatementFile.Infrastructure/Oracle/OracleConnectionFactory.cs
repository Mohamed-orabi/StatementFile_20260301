using System;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Infrastructure.Oracle
{
    /// <summary>
    /// Creates OracleConnection instances from the configured connection string.
    /// Registered as a singleton in DI.
    /// </summary>
    public sealed class OracleConnectionFactory : IDbConnectionFactory
    {
        public string ConnectionString { get; }

        public OracleConnectionFactory(string connectionString)
        {
            if (string.IsNullOrWhiteSpace(connectionString))
                throw new ArgumentException("Oracle connection string is required.", nameof(connectionString));
            ConnectionString = connectionString;
        }

        public IDbConnection CreateConnection()
        {
            var conn = new OracleConnection(ConnectionString);
            conn.Open();
            return conn;
        }
    }
}

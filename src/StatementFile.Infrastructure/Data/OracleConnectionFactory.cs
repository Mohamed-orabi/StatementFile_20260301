using System;
using Oracle.ManagedDataAccess.Client;
using StatementFile.Application.Interfaces;

namespace StatementFile.Infrastructure.Data
{
    /// <summary>
    /// Creates and opens Oracle connections using the connection string
    /// supplied by <see cref="IConfigurationService"/>.
    /// Centralises all ODP.NET connection concerns in one place.
    /// </summary>
    public sealed class OracleConnectionFactory
    {
        private readonly IConfigurationService _config;

        public OracleConnectionFactory(IConfigurationService config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
        }

        /// <summary>Returns a new, open Oracle connection.</summary>
        public OracleConnection CreateOpenConnection()
        {
            var conn = new OracleConnection(_config.GetOracleConnectionString());
            conn.Open();
            return conn;
        }
    }
}

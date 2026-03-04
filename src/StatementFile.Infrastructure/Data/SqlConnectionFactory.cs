using System;
using Microsoft.Data.SqlClient;
using StatementFile.Application.Interfaces;

namespace StatementFile.Infrastructure.Data
{
    /// <summary>
    /// Creates and opens SQL Server connections using the connection string
    /// supplied by <see cref="IConfigurationService"/>.
    /// Centralises all SqlClient connection concerns in one place.
    /// </summary>
    public sealed class SqlConnectionFactory
    {
        private readonly IConfigurationService _config;

        public SqlConnectionFactory(IConfigurationService config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
        }

        /// <summary>Returns a new, open SQL Server connection.</summary>
        public SqlConnection CreateOpenConnection()
        {
            var conn = new SqlConnection(_config.GetSqlConnectionString());
            conn.Open();
            return conn;
        }
    }
}

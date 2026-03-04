using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using StatementFile.Application.Interfaces;
using StatementFile.Domain.Entities;
using StatementFile.Domain.Interfaces.Repositories;

namespace StatementFile.Infrastructure.Data.Repositories
{
    public sealed class BankRepository : IBankRepository
    {
        private readonly SqlConnection         _conn;
        private readonly IConfigurationService _config;

        public BankRepository(SqlConnection conn, IConfigurationService config)
        {
            _conn   = conn   ?? throw new ArgumentNullException(nameof(conn));
            _config = config ?? throw new ArgumentNullException(nameof(config));
        }

        public Bank GetByBranchCode(int branchCode)
        {
            string schema = _config.GetMainSchema();
            string sql    = $"SELECT branch, branchname, branchpart, ident FROM {schema}TBRANCH WHERE branch = @code";

            using (var cmd = new SqlCommand(sql, _conn))
            {
                cmd.Parameters.Add(new SqlParameter("@code", SqlDbType.Int) { Value = branchCode });
                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                        return MapRow(reader);
                }
            }
            return null;
        }

        public IEnumerable<Bank> GetAll()
        {
            string schema = _config.GetMainSchema();
            string sql    = $"SELECT branch, branchname, branchpart, ident FROM {schema}TBRANCH ORDER BY branch";
            var    result = new List<Bank>();

            using (var cmd = new SqlCommand(sql, _conn))
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                    result.Add(MapRow(reader));
            }
            return result;
        }

        private static Bank MapRow(IDataReader r)
        {
            return Bank.Create(
                branchCode: r.GetInt32(0),
                branchName: r.IsDBNull(1) ? string.Empty : r.GetString(1),
                branchPart: r.IsDBNull(2) ? string.Empty : r.GetString(2),
                ident:      r.IsDBNull(3) ? string.Empty : r.GetString(3));
        }
    }
}

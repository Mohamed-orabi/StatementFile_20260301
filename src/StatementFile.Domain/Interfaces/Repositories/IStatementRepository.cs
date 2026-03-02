using System.Collections.Generic;
using System.Data;
using StatementFile.Domain.Entities;

namespace StatementFile.Domain.Interfaces.Repositories
{
    /// <summary>
    /// Contract for reading and writing statement master/detail data.
    /// Implementations live in the Infrastructure layer (Oracle ODP.NET).
    /// </summary>
    public interface IStatementRepository
    {
        /// <summary>
        /// Loads a raw master DataSet for the given branch from Oracle.
        /// The schema / table name is resolved at runtime from session values.
        /// </summary>
        DataSet LoadMasterDataSet(int branchCode, string orderBy, string additionalCondition = null);

        /// <summary>
        /// Loads a raw detail DataSet for the given branch from Oracle.
        /// </summary>
        DataSet LoadDetailDataSet(int branchCode, string additionalCondition = null);

        /// <summary>
        /// Loads email address records for electronic-statement delivery.
        /// </summary>
        DataSet LoadEmailDataSet(int branchCode);

        /// <summary>
        /// Executes a PL/SQL batch block (BEGIN ... END;) against Oracle.
        /// Used by bulk-maintenance operations.
        /// </summary>
        int ExecuteBatch(string plSqlBlock);

        /// <summary>
        /// Executes a single DML statement and returns the number of rows affected.
        /// </summary>
        int ExecuteAction(string sql);
    }
}

using System.Data;

namespace StatementFile.Domain.Interfaces
{
    /// <summary>Abstracts Oracle connection creation so tests can substitute a fake.</summary>
    public interface IDbConnectionFactory
    {
        IDbConnection CreateConnection();
        string ConnectionString { get; }
    }
}

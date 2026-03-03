using System.Threading.Tasks;

namespace StatementFile.Presentation.Services
{
    /// <summary>
    /// Async HTTP client abstraction for the authentication API.
    /// Pings the API to verify Oracle connectivity before allowing the user in.
    /// </summary>
    public interface IAuthApiClient
    {
        /// <summary>
        /// Returns true if the API can reach the Oracle database.
        /// Throws <see cref="System.Net.Http.HttpRequestException"/> on network failure.
        /// </summary>
        Task<bool> PingAsync();
    }
}

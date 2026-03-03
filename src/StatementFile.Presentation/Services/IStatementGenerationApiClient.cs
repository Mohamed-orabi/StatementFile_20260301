using System.Threading.Tasks;
using StatementFile.Application.UseCases.StatementGeneration;

namespace StatementFile.Presentation.Services
{
    /// <summary>
    /// Async HTTP client abstraction for the statement generation API.
    /// Blazor pages call this to trigger generation without touching Infrastructure directly.
    /// </summary>
    public interface IStatementGenerationApiClient
    {
        /// <summary>
        /// Runs statement generation for the given bank/product configuration and date.
        /// The API loads the full configuration from the database using ConfigId.
        /// </summary>
        Task<StatementRunResult> RunAsync(StatementRunRequest request);
    }
}

using System.Threading.Tasks;
using StatementFile.Application.UseCases.MerchantOnboarding;

namespace StatementFile.Presentation.Services
{
    /// <summary>
    /// Async HTTP client abstraction for the merchant statement API.
    /// Blazor pages call this to process merchant XML files without touching Infrastructure directly.
    /// </summary>
    public interface IMerchantStatementApiClient
    {
        /// <summary>Processes a merchant XML file sent as base64-encoded content.</summary>
        Task<MerchantProcessResult> ProcessAsync(MerchantProcessRequest request);
    }
}

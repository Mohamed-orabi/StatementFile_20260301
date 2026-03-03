using System;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
using StatementFile.Application.UseCases.MerchantOnboarding;

namespace StatementFile.Presentation.Services
{
    /// <summary>
    /// HttpClient implementation of <see cref="IMerchantStatementApiClient"/>.
    /// Calls POST /api/merchant-statement/process on the StatementFile.Api project.
    /// </summary>
    public sealed class MerchantStatementApiClient : IMerchantStatementApiClient
    {
        private readonly HttpClient _http;

        public MerchantStatementApiClient(HttpClient http)
        {
            _http = http ?? throw new ArgumentNullException(nameof(http));
        }

        public async Task<MerchantProcessResult> ProcessAsync(MerchantProcessRequest request)
        {
            var response = await _http.PostAsJsonAsync("api/merchant-statement/process", request);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<MerchantProcessResult>();
        }
    }
}

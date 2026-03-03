using System;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
using StatementFile.Application.UseCases.StatementGeneration;

namespace StatementFile.Presentation.Services
{
    /// <summary>
    /// HttpClient implementation of <see cref="IStatementGenerationApiClient"/>.
    /// Calls POST /api/statement-generation/run on the StatementFile.Api project.
    /// </summary>
    public sealed class StatementGenerationApiClient : IStatementGenerationApiClient
    {
        private readonly HttpClient _http;

        public StatementGenerationApiClient(HttpClient http)
        {
            _http = http ?? throw new ArgumentNullException(nameof(http));
        }

        public async Task<StatementRunResult> RunAsync(StatementRunRequest request)
        {
            var response = await _http.PostAsJsonAsync("api/statement-generation/run", request);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<StatementRunResult>();
        }
    }
}

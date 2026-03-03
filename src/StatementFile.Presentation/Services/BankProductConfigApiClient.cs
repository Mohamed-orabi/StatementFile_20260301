using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
using StatementFile.Application.UseCases.BankConfiguration;

namespace StatementFile.Presentation.Services
{
    /// <summary>
    /// HttpClient implementation of <see cref="IBankProductConfigApiClient"/>.
    ///
    /// Calls the embedded REST API exposed by <c>BankConfigController</c>.
    /// The base address is configured in <c>appsettings.json → ApiBaseUrl</c>
    /// and registered in Program.cs via <c>AddHttpClient</c>.
    ///
    /// All methods are async to avoid blocking Blazor Server's synchronisation context.
    /// </summary>
    public sealed class BankProductConfigApiClient : IBankProductConfigApiClient
    {
        private readonly HttpClient _http;

        public BankProductConfigApiClient(HttpClient http)
        {
            _http = http ?? throw new ArgumentNullException(nameof(http));
        }

        public async Task<IReadOnlyList<BankProductConfigDto>> GetAllAsync()
        {
            var result = await _http.GetFromJsonAsync<List<BankProductConfigDto>>(
                "api/bank-config");
            return result ?? new List<BankProductConfigDto>();
        }

        public async Task<IReadOnlyList<BankProductConfigDto>> GetActiveAsync()
        {
            var result = await _http.GetFromJsonAsync<List<BankProductConfigDto>>(
                "api/bank-config?activeOnly=true");
            return result ?? new List<BankProductConfigDto>();
        }

        public async Task<BankProductConfigDto> GetByIdAsync(int id)
        {
            var response = await _http.GetAsync($"api/bank-config/{id}");
            if (response.StatusCode == HttpStatusCode.NotFound)
                return null;
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<BankProductConfigDto>();
        }

        public async Task<int> CreateAsync(SaveBankProductConfigRequest request)
        {
            var response = await _http.PostAsJsonAsync("api/bank-config", request);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<int>();
        }

        public async Task UpdateAsync(int id, SaveBankProductConfigRequest request)
        {
            var response = await _http.PutAsJsonAsync($"api/bank-config/{id}", request);
            response.EnsureSuccessStatusCode();
        }

        public async Task DeleteAsync(int id)
        {
            var response = await _http.DeleteAsync($"api/bank-config/{id}");
            response.EnsureSuccessStatusCode();
        }
    }
}

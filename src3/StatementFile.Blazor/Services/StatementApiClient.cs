using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
using StatementFile.Application.DTOs;

namespace StatementFile.Blazor.Services
{
    /// <summary>
    /// Typed HttpClient that wraps all StatementFile API endpoints.
    /// Injected into Blazor pages via DI; registered in Program.cs.
    /// </summary>
    public sealed class StatementApiClient
    {
        private readonly HttpClient _http;

        public StatementApiClient(HttpClient http) =>
            _http = http ?? throw new ArgumentNullException(nameof(http));

        // ── Auth ──────────────────────────────────────────────────────────────────

        public async Task<bool> ValidateConnectionAsync()
        {
            try
            {
                var resp = await _http.PostAsync("api/auth/validate", null);
                return resp.IsSuccessStatusCode;
            }
            catch { return false; }
        }

        // ── Bank configs ──────────────────────────────────────────────────────────

        public Task<List<BankConfigDto>> GetAllConfigsAsync() =>
            _http.GetFromJsonAsync<List<BankConfigDto>>("api/bank-configs");

        public Task<List<BankConfigDto>> GetActiveConfigsAsync() =>
            _http.GetFromJsonAsync<List<BankConfigDto>>("api/bank-configs/active");

        public Task<BankConfigDto> GetConfigAsync(int id) =>
            _http.GetFromJsonAsync<BankConfigDto>($"api/bank-configs/{id}");

        public async Task<int> CreateConfigAsync(BankConfigDto dto)
        {
            var resp = await _http.PostAsJsonAsync("api/bank-configs", dto);
            resp.EnsureSuccessStatusCode();
            return await resp.Content.ReadFromJsonAsync<int>();
        }

        public async Task UpdateConfigAsync(int id, BankConfigDto dto)
        {
            var resp = await _http.PutAsJsonAsync($"api/bank-configs/{id}", dto);
            resp.EnsureSuccessStatusCode();
        }

        public async Task DeleteConfigAsync(int id)
        {
            var resp = await _http.DeleteAsync($"api/bank-configs/{id}");
            resp.EnsureSuccessStatusCode();
        }

        // ── Statement generation ──────────────────────────────────────────────────

        public async Task<GenerationResultDto> GenerateAsync(GenerateStatementRequest req)
        {
            var resp = await _http.PostAsJsonAsync("api/statements/generate", req);
            return await resp.Content.ReadFromJsonAsync<GenerationResultDto>();
        }

        public async Task<List<GenerationResultDto>> GenerateBulkAsync(GenerateBulkRequest req)
        {
            var resp = await _http.PostAsJsonAsync("api/statements/generate-bulk", req);
            resp.EnsureSuccessStatusCode();
            return await resp.Content.ReadFromJsonAsync<List<GenerationResultDto>>();
        }

        // ── Maintenance ───────────────────────────────────────────────────────────

        public async Task<bool> RunMaintenanceAsync(MaintenanceRequest req)
        {
            var resp = await _http.PostAsJsonAsync("api/maintenance/run", req);
            return resp.IsSuccessStatusCode;
        }

        // ── Formatter keys ────────────────────────────────────────────────────────

        public Task<List<string>> GetFormatterKeysAsync() =>
            _http.GetFromJsonAsync<List<string>>("api/formatter-keys");
    }
}

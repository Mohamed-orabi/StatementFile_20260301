using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
using StatementFile.Application.DTOs;

namespace StatementFile.Presentation.Services
{
    /// <summary>
    /// Typed HTTP client for the StatementFile REST API.
    ///
    /// Registered as a typed HttpClient in Program.cs so ASP.NET Core's
    /// HttpClientFactory manages connection pooling and lifetime.
    ///
    /// All Blazor pages inject this service instead of the legacy
    /// <c>CompositionRoot</c>, removing every direct Infrastructure reference
    /// from the Presentation layer.
    /// </summary>
    public sealed class StatementApiClient
    {
        private readonly HttpClient _http;

        public StatementApiClient(HttpClient http) => _http = http;

        // ── Auth ──────────────────────────────────────────────────────────────

        /// <summary>
        /// Tests the Oracle DB connection on the API side.
        /// Returns true when the API returns HTTP 200.
        /// </summary>
        public async Task<(bool Success, string Error)> ValidateConnectionAsync()
        {
            try
            {
                var resp = await _http.PostAsync("api/auth/validate", null);
                if (resp.IsSuccessStatusCode)
                    return (true, null);

                var body = await resp.Content.ReadAsStringAsync();
                return (false, body);
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        // ── Bank Configurations ───────────────────────────────────────────────

        /// <summary>Returns all active bank / product configurations.</summary>
        public async Task<List<BankConfigDto>> GetActiveBankConfigsAsync()
        {
            return await _http.GetFromJsonAsync<List<BankConfigDto>>("api/bank-configs/active")
                   ?? new List<BankConfigDto>();
        }

        /// <summary>Returns all bank / product configurations (active + inactive).</summary>
        public async Task<List<BankConfigDto>> GetAllBankConfigsAsync()
        {
            return await _http.GetFromJsonAsync<List<BankConfigDto>>("api/bank-configs")
                   ?? new List<BankConfigDto>();
        }

        /// <summary>Returns a single configuration by ID, or null if not found.</summary>
        public async Task<BankConfigDto> GetBankConfigAsync(int id)
        {
            var resp = await _http.GetAsync($"api/bank-configs/{id}");
            if (resp.StatusCode == System.Net.HttpStatusCode.NotFound) return null;
            resp.EnsureSuccessStatusCode();
            return await resp.Content.ReadFromJsonAsync<BankConfigDto>();
        }

        /// <summary>Creates a new configuration and returns the new ID.</summary>
        public async Task<(int Id, string Error)> CreateBankConfigAsync(BankConfigDto dto)
        {
            try
            {
                var resp = await _http.PostAsJsonAsync("api/bank-configs", dto);
                if (!resp.IsSuccessStatusCode)
                {
                    var err = await resp.Content.ReadAsStringAsync();
                    return (0, err);
                }
                int newId = await resp.Content.ReadFromJsonAsync<int>();
                return (newId, null);
            }
            catch (Exception ex)
            {
                return (0, ex.Message);
            }
        }

        /// <summary>Updates an existing configuration.</summary>
        public async Task<string> UpdateBankConfigAsync(int id, BankConfigDto dto)
        {
            try
            {
                var resp = await _http.PutAsJsonAsync($"api/bank-configs/{id}", dto);
                if (!resp.IsSuccessStatusCode)
                    return await resp.Content.ReadAsStringAsync();
                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        /// <summary>Deletes a configuration permanently.</summary>
        public async Task<string> DeleteBankConfigAsync(int id)
        {
            try
            {
                var resp = await _http.DeleteAsync($"api/bank-configs/{id}");
                if (!resp.IsSuccessStatusCode)
                    return await resp.Content.ReadAsStringAsync();
                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        // ── Statement Generation ──────────────────────────────────────────────

        /// <summary>
        /// Generates statements for a single bank / product.
        /// Runs bulk maintenance + full pipeline.
        /// </summary>
        public async Task<GenerationResultDto> GenerateStatementAsync(GenerateStatementApiRequest req)
        {
            try
            {
                var resp = await _http.PostAsJsonAsync("api/statements/generate", req);
                return await resp.Content.ReadFromJsonAsync<GenerationResultDto>()
                       ?? new GenerationResultDto { Success = false, Error = "Empty response from API." };
            }
            catch (Exception ex)
            {
                return new GenerationResultDto { Success = false, Error = ex.Message };
            }
        }

        /// <summary>
        /// Generates statements for multiple bank / product combinations in sequence.
        /// </summary>
        public async Task<List<GenerationResultDto>> GenerateBulkAsync(
            List<GenerateStatementApiRequest> requests)
        {
            try
            {
                var resp = await _http.PostAsJsonAsync("api/statements/generate-bulk", requests);
                return await resp.Content.ReadFromJsonAsync<List<GenerationResultDto>>()
                       ?? new List<GenerationResultDto>();
            }
            catch (Exception ex)
            {
                return new List<GenerationResultDto>
                {
                    new GenerationResultDto { Success = false, Error = ex.Message }
                };
            }
        }

        // ── Merchant Statement ────────────────────────────────────────────────

        /// <summary>Processes a merchant XML statement file.</summary>
        public async Task<GenerationResultDto> ProcessMerchantAsync(ProcessMerchantApiRequest req)
        {
            try
            {
                var resp = await _http.PostAsJsonAsync("api/merchants/process", req);
                return await resp.Content.ReadFromJsonAsync<GenerationResultDto>()
                       ?? new GenerationResultDto { Success = false, Error = "Empty response from API." };
            }
            catch (Exception ex)
            {
                return new GenerationResultDto { Success = false, Error = ex.Message };
            }
        }

        // ── Bulk Maintenance (standalone) ─────────────────────────────────────

        /// <summary>Runs bulk data maintenance independently of statement generation.</summary>
        public async Task<(bool Success, string Error)> RunMaintenanceAsync(int branchCode, long processingModes)
        {
            try
            {
                var req = new { BranchCode = branchCode, ProcessingModes = processingModes };
                var resp = await _http.PostAsJsonAsync("api/maintenance/run", req);
                return resp.IsSuccessStatusCode ? (true, null) : (false, await resp.Content.ReadAsStringAsync());
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }
    }
}

using System;
using System.Net.Http;
using System.Threading.Tasks;

namespace StatementFile.Presentation.Services
{
    /// <summary>
    /// HttpClient implementation of <see cref="IAuthApiClient"/>.
    /// Calls POST /api/auth/ping on the StatementFile.Api project.
    /// </summary>
    public sealed class AuthApiClient : IAuthApiClient
    {
        private readonly HttpClient _http;

        public AuthApiClient(HttpClient http)
        {
            _http = http ?? throw new ArgumentNullException(nameof(http));
        }

        public async Task<bool> PingAsync()
        {
            var response = await _http.PostAsync("api/auth/ping", content: null);
            return response.IsSuccessStatusCode;
        }
    }
}

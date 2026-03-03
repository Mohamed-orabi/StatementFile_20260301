using System.Collections.Generic;
using System.Threading.Tasks;
using StatementFile.Application.UseCases.BankConfiguration;

namespace StatementFile.Presentation.Services
{
    /// <summary>
    /// Async HTTP client abstraction consumed by Blazor pages.
    ///
    /// Decouples the Presentation layer from the Infrastructure repositories:
    /// Blazor components → IBankProductConfigApiClient (HTTP) → BankConfigController → IBankProductConfigRepository
    /// </summary>
    public interface IBankProductConfigApiClient
    {
        /// <summary>Returns all configurations (active and inactive).</summary>
        Task<IReadOnlyList<BankProductConfigDto>> GetAllAsync();

        /// <summary>Returns only active configurations ordered by SortOrder.</summary>
        Task<IReadOnlyList<BankProductConfigDto>> GetActiveAsync();

        /// <summary>Returns the configuration with the given ID, or null when not found.</summary>
        Task<BankProductConfigDto> GetByIdAsync(int id);

        /// <summary>Creates a new configuration and returns the new ID.</summary>
        Task<int> CreateAsync(SaveBankProductConfigRequest request);

        /// <summary>Updates mutable fields of an existing configuration.</summary>
        Task UpdateAsync(int id, SaveBankProductConfigRequest request);

        /// <summary>Deletes the configuration with the given ID.</summary>
        Task DeleteAsync(int id);
    }
}

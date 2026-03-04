using System;
using Microsoft.AspNetCore.Mvc;
using StatementFile.Infrastructure.Configuration;

namespace StatementFile.Api.Controllers
{
    /// <summary>
    /// Validates that the configured Oracle credentials can open a live connection.
    ///
    /// The API uses a pre-configured connection string (App.config / appsettings.json).
    /// This endpoint allows the Blazor front-end to confirm the database is reachable
    /// before showing the main UI, preserving the legacy login-check behaviour.
    /// </summary>
    [ApiController]
    [Route("api/auth")]
    public sealed class AuthController : ControllerBase
    {
        private readonly CompositionRoot _root;

        public AuthController(CompositionRoot root) => _root = root;

        /// <summary>
        /// Opens (and immediately disposes) a Unit-of-Work to verify DB connectivity.
        /// Returns 200 on success, 400 with an error message on failure.
        /// </summary>
        [HttpPost("validate")]
        public IActionResult Validate()
        {
            try
            {
                using var uow = _root.CreateUnitOfWork();
                return Ok(new { success = true, message = "Database connection successful." });
            }
            catch (Exception ex)
            {
                return BadRequest(new { success = false, message = ex.Message });
            }
        }
    }
}

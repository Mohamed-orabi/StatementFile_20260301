using System;
using Microsoft.AspNetCore.Mvc;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Api.Controllers
{
    /// <summary>
    /// Validates the Oracle database connection.
    ///
    /// Routes:
    ///   POST /api/auth/validate  → tests DB connectivity; returns 200 on success
    /// </summary>
    [ApiController]
    [Route("api/auth")]
    public sealed class AuthController : ControllerBase
    {
        private readonly IDbConnectionFactory _factory;

        public AuthController(IDbConnectionFactory factory) =>
            _factory = factory ?? throw new ArgumentNullException(nameof(factory));

        [HttpPost("validate")]
        public IActionResult Validate()
        {
            try
            {
                using var conn = _factory.CreateConnection();
                return Ok(new { success = true, message = "Database connection successful." });
            }
            catch (Exception ex)
            {
                return StatusCode(503, new { success = false, error = ex.Message });
            }
        }
    }
}

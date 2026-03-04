using System;
using Microsoft.AspNetCore.Mvc;
using StatementFile.Infrastructure.Data;

namespace StatementFile.Api.Controllers
{
    /// <summary>
    /// Validates the SQL Server database connection via EF Core.
    ///
    /// Routes:
    ///   POST /api/auth/validate  → tests DB connectivity; returns 200 on success
    /// </summary>
    [ApiController]
    [Route("api/auth")]
    public sealed class AuthController : ControllerBase
    {
        private readonly StatementFileDbContext _db;

        public AuthController(StatementFileDbContext db) =>
            _db = db ?? throw new ArgumentNullException(nameof(db));

        [HttpPost("validate")]
        public IActionResult Validate()
        {
            try
            {
                // CanConnect() tests the connection without running a migration.
                var ok = _db.Database.CanConnect();
                return ok
                    ? Ok(new { success = true,  message = "Database connection successful." })
                    : StatusCode(503, new { success = false, error = "Cannot connect to database." });
            }
            catch (Exception ex)
            {
                return StatusCode(503, new { success = false, error = ex.Message });
            }
        }
    }
}

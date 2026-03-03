using System;
using Microsoft.AspNetCore.Mvc;
using StatementFile.Infrastructure.Configuration;

namespace StatementFile.Api.Controllers
{
    /// <summary>
    /// REST API for application authentication.
    ///
    /// Routes
    ///   POST /api/auth/ping — verifies Oracle connectivity by opening a UnitOfWork
    ///
    /// Note: Oracle credentials come from App.config (not from the request).
    /// This endpoint confirms the database is reachable, not Oracle user identity.
    /// </summary>
    [ApiController]
    [Route("api/auth")]
    public sealed class AuthController : ControllerBase
    {
        private readonly CompositionRoot _root;

        public AuthController(CompositionRoot root)
        {
            _root = root ?? throw new ArgumentNullException(nameof(root));
        }

        // ── POST /api/auth/ping ───────────────────────────────────────────────

        [HttpPost("ping")]
        public IActionResult Ping()
        {
            try
            {
                using var uow = _root.CreateUnitOfWork();
                return Ok();
            }
            catch (Exception ex)
            {
                return StatusCode(503, ex.Message);
            }
        }
    }
}

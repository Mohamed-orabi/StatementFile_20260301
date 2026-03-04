using Microsoft.AspNetCore.Mvc;
using StatementFile.Infrastructure.Formatters;

namespace StatementFile.Api.Controllers
{
    /// <summary>
    /// Returns the list of formatter keys registered in the application.
    /// Useful for the front-end to validate a FormatterKey before creating
    /// or updating a bank configuration.
    ///
    /// Routes:
    ///   GET /api/formatter-keys
    /// </summary>
    [ApiController]
    [Route("api/formatter-keys")]
    public sealed class FormatterKeysController : ControllerBase
    {
        private readonly FormatterRegistry _registry;

        public FormatterKeysController(FormatterRegistry registry) => _registry = registry;

        [HttpGet]
        public IActionResult GetAll() => Ok(_registry.RegisteredKeys);
    }
}

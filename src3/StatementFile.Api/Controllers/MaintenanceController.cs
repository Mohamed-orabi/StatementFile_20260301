using System;
using Microsoft.AspNetCore.Mvc;
using StatementFile.Application.Commands;
using StatementFile.Application.DTOs;
using StatementFile.Domain.Enums;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Api.Controllers
{
    /// <summary>
    /// Exposes bulk data-maintenance as a standalone endpoint so operators can
    /// run maintenance independently of statement generation (e.g. to fix data
    /// issues between billing cycles).
    ///
    /// Routes:
    ///   POST /api/maintenance/run
    /// </summary>
    [ApiController]
    [Route("api/maintenance")]
    public sealed class MaintenanceController : ControllerBase
    {
        private readonly RunMaintenanceHandler _handler;
        private readonly IConfigurationService _config;

        public MaintenanceController(RunMaintenanceHandler handler, IConfigurationService config)
        {
            _handler = handler ?? throw new ArgumentNullException(nameof(handler));
            _config  = config  ?? throw new ArgumentNullException(nameof(config));
        }

        [HttpPost("run")]
        public IActionResult Run([FromBody] MaintenanceRequest req)
        {
            if (req == null)
                return BadRequest(new { error = "Request body is required." });

            var result = _handler.Handle(new RunMaintenanceCommand
            {
                BranchCode              = req.BranchCode,
                ProcessingModes         = (ProcessingMode)req.ProcessingModes,
                RunNullCardDelete       = req.RunNullCardDelete,
                ExcludeReward           = req.ExcludeReward,
                RunCardBranchMatch      = req.RunCardBranchMatch,
                RewardContractCondition = req.RewardContractCondition,
                InstallmentCondition    = req.InstallmentCondition,
                ConnectionString        = _config.GetOracleConnectionString(),
            });

            if (result.IsSuccess)
                return Ok(new { success = true });

            return StatusCode(500, new { success = false, error = result.Error });
        }
    }
}

using System;
using System.Collections.Generic;
using Microsoft.AspNetCore.Mvc;
using StatementFile.Application.Commands;
using StatementFile.Application.DTOs;
using StatementFile.Domain.Enums;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Api.Controllers
{
    /// <summary>
    /// Triggers statement generation for one or more bank / product configurations.
    ///
    /// Routes:
    ///   POST /api/statements/generate       → single generation run
    ///   POST /api/statements/generate-bulk  → multiple runs in sequence
    ///
    /// Each request resolves the <see cref="BankProductConfig"/> row by its ID,
    /// runs pre-generation maintenance, then invokes the correct formatter via
    /// <see cref="GenerateStatementHandler"/>.
    /// </summary>
    [ApiController]
    [Route("api/statements")]
    public sealed class StatementController : ControllerBase
    {
        private readonly IBankProductConfigRepository _repo;
        private readonly GenerateStatementHandler     _generator;
        private readonly RunMaintenanceHandler        _maintenance;
        private readonly IConfigurationService        _config;

        public StatementController(
            IBankProductConfigRepository repo,
            GenerateStatementHandler     generator,
            RunMaintenanceHandler        maintenance,
            IConfigurationService        config)
        {
            _repo        = repo        ?? throw new ArgumentNullException(nameof(repo));
            _generator   = generator   ?? throw new ArgumentNullException(nameof(generator));
            _maintenance = maintenance ?? throw new ArgumentNullException(nameof(maintenance));
            _config      = config      ?? throw new ArgumentNullException(nameof(config));
        }

        // ── Single generation ─────────────────────────────────────────────────────

        /// <summary>
        /// Runs pre-generation maintenance then generates statements for a single
        /// bank / product configuration identified by <c>ConfigId</c>.
        /// </summary>
        [HttpPost("generate")]
        public IActionResult Generate([FromBody] GenerateStatementRequest req)
        {
            if (req == null)
                return BadRequest(new { error = "Request body is required." });

            var config = _repo.GetById(req.ConfigId);
            if (config == null)
                return NotFound(new { error = $"No active configuration found for ID {req.ConfigId}." });

            if (!config.IsActive)
                return BadRequest(new { error = $"Configuration ID {req.ConfigId} is inactive." });

            var connStr   = _config.GetSqlConnectionString();
            var outputDir = string.IsNullOrWhiteSpace(req.OutputDirectory)
                ? _config.GetStatementOutputPath()
                : req.OutputDirectory;

            // 1. Maintenance
            var maintResult = _maintenance.Handle(new RunMaintenanceCommand
            {
                BranchCode              = config.BranchCode,
                ProcessingModes         = config.ProcessingModes,
                RunNullCardDelete       = config.RunNullCardDelete,
                ExcludeReward           = config.ExcludeReward,
                RunCardBranchMatch      = config.RunCardBranchMatch,
                RewardContractCondition = config.RewardContractCondition,
                InstallmentCondition    = config.InstallmentCondition,
                ConnectionString        = connStr,
            });

            if (!maintResult.IsSuccess)
                return StatusCode(500, new GenerationResultDto
                {
                    Success = false,
                    Error   = $"Maintenance failed: {maintResult.Error}",
                });

            // 2. Generation
            var genResult = _generator.Handle(new GenerateStatementCommand
            {
                Config          = config,
                StatementDate   = req.StatementDate,
                OutputDirectory = outputDir,
                EmailOverride   = req.EmailOverride,
                AppendData      = req.AppendData,
                ConnectionString= connStr,
                ReportTemplate  = _config.GetReportTemplatePath(),
            });

            if (!genResult.IsSuccess)
                return StatusCode(500, new GenerationResultDto
                {
                    Success = false,
                    Error   = genResult.Error,
                });

            return Ok(new GenerationResultDto
            {
                Success         = true,
                FilesGenerated  = genResult.FilesGenerated,
                EmailsSent      = genResult.EmailsSent,
                StatementsCount = genResult.StatementsCount,
                OutputDirectory = genResult.OutputDirectory,
            });
        }

        // ── Bulk generation ───────────────────────────────────────────────────────

        /// <summary>
        /// Executes multiple generation requests in sequence.
        /// Returns one <see cref="GenerationResultDto"/> per input item.
        /// </summary>
        [HttpPost("generate-bulk")]
        public IActionResult GenerateBulk([FromBody] GenerateBulkRequest req)
        {
            if (req?.Items == null || req.Items.Length == 0)
                return BadRequest(new { error = "No items in request." });

            var results = new List<GenerationResultDto>();

            foreach (var item in req.Items)
            {
                var singleResult = Generate(item) as ObjectResult;
                if (singleResult?.Value is GenerationResultDto dto)
                    results.Add(dto);
                else
                    results.Add(new GenerationResultDto
                    {
                        Success = false,
                        Error   = "Unexpected error processing bulk item.",
                    });
            }

            return Ok(results);
        }
    }
}

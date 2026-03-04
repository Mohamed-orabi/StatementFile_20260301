using System;
using Microsoft.AspNetCore.Mvc;
using StatementFile.Application.UseCases.BulkProcessing;
using StatementFile.Domain.Enums;
using StatementFile.Infrastructure.Configuration;

namespace StatementFile.Api.Controllers
{
    /// <summary>
    /// Exposes the bulk data-maintenance operations as a standalone endpoint.
    ///
    /// Maintenance typically runs automatically before each statement generation
    /// (see <see cref="StatementController"/>), but this endpoint allows operators
    /// to trigger it independently — for example, to fix data issues between
    /// generation cycles.
    /// </summary>
    [ApiController]
    [Route("api/maintenance")]
    public sealed class MaintenanceController : ControllerBase
    {
        private readonly CompositionRoot _root;

        public MaintenanceController(CompositionRoot root) => _root = root;

        /// <summary>Runs bulk data maintenance for a single branch.</summary>
        [HttpPost("run")]
        public IActionResult Run([FromBody] BulkMaintenanceApiRequest req)
        {
            if (req == null)
                return BadRequest(new { error = "Request body is required." });

            try
            {
                var processingModes = (ProcessingMode)req.ProcessingModes;

                var cmd = new RunBulkMaintenanceCommand(
                    branchCode:             req.BranchCode,
                    runNullCardDelete:      req.RunNullCardDelete,
                    excludeReward:          req.ExcludeReward,
                    excludeInstallment:     processingModes.HasFlag(ProcessingMode.Installment),
                    installmentCondition:   req.InstallmentCondition,
                    rewardContractCondition:req.RewardContractCondition ?? "'New Reward Contract'",
                    runCardBranchMatch:     req.RunCardBranchMatch,
                    runArabicAddressFix:    processingModes.HasFlag(ProcessingMode.FixArabicAddress),
                    runFixAddress:          processingModes.HasFlag(ProcessingMode.FixAddress),
                    runFixArabicAddressLang:processingModes.HasFlag(ProcessingMode.FixArabicAddressLang),
                    runOnHoldDelete:        processingModes.HasFlag(ProcessingMode.DeleteOnHold),
                    isRewardRun:            processingModes.HasFlag(ProcessingMode.Reward),
                    runMergeMarkUpFees:     processingModes.HasFlag(ProcessingMode.MergeMarkUpFees),
                    runRewardFix:           processingModes.HasFlag(ProcessingMode.Reward));

                var result = _root.BulkHandler.Handle(cmd);

                if (result.IsSuccess)
                    return Ok(new { success = true });

                return Ok(new { success = false, error = result.Error });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, error = ex.Message });
            }
        }
    }

    /// <summary>Request DTO for the bulk maintenance endpoint.</summary>
    public sealed class BulkMaintenanceApiRequest
    {
        public int    BranchCode              { get; set; }
        /// <summary>ProcessingMode flags as a long integer.</summary>
        public long   ProcessingModes         { get; set; }
        public bool   RunNullCardDelete       { get; set; } = true;
        public bool   ExcludeReward           { get; set; } = true;
        public bool   RunCardBranchMatch      { get; set; } = true;
        public string RewardContractCondition { get; set; } = "'New Reward Contract'";
        public string InstallmentCondition    { get; set; }
    }
}

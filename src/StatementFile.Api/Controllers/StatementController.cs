using System;
using System.Collections.Generic;
using Microsoft.AspNetCore.Mvc;
using StatementFile.Application.DTOs;
using StatementFile.Application.UseCases.BulkProcessing;
using StatementFile.Application.UseCases.StatementGeneration;
using StatementFile.Domain.Enums;
using StatementFile.Infrastructure.Configuration;

namespace StatementFile.Api.Controllers
{
    /// <summary>
    /// Triggers statement generation for one bank / product combination.
    ///
    /// Each call executes the full 10-step pipeline:
    ///   1. Bulk data maintenance (card-branch match, Arabic fix, on-hold delete…)
    ///   2. Statement generation (Oracle load → formatter → email / FTP delivery)
    ///
    /// The processing mode flags carried in the request determine which
    /// optional maintenance steps are activated, exactly as the legacy
    /// frmStatementFileExtn.cs switch-case did per product.
    /// </summary>
    [ApiController]
    [Route("api/statements")]
    public sealed class StatementController : ControllerBase
    {
        private readonly CompositionRoot _root;

        public StatementController(CompositionRoot root) => _root = root;

        /// <summary>
        /// Generates statements for a single bank / product.
        /// Runs bulk maintenance first, then the full generation pipeline.
        /// </summary>
        [HttpPost("generate")]
        public ActionResult<GenerationResultDto> Generate([FromBody] GenerateStatementApiRequest req)
        {
            if (req == null)
                return BadRequest(new GenerationResultDto { Success = false, Error = "Request body is required." });
            if (string.IsNullOrWhiteSpace(req.BankName))
                return BadRequest(new GenerationResultDto { Success = false, Error = "BankName is required." });
            if (string.IsNullOrWhiteSpace(req.FormatterKey))
                return BadRequest(new GenerationResultDto { Success = false, Error = "FormatterKey is required." });

            try
            {
                var processingModes = (ProcessingMode)req.ProcessingModes;

                // ── Step 1: Bulk maintenance ─────────────────────────────────
                var bulkCmd = new RunBulkMaintenanceCommand(
                    branchCode:          req.BranchCode,
                    runCardBranchMatch:  true,
                    runArabicAddressFix: processingModes.HasFlag(ProcessingMode.FixArabicAddress),
                    runFixAddress:       processingModes.HasFlag(ProcessingMode.FixAddress),
                    runFixArabicAddressLang: processingModes.HasFlag(ProcessingMode.FixArabicAddressLang),
                    runOnHoldDelete:     processingModes.HasFlag(ProcessingMode.DeleteOnHold),
                    isRewardRun:         processingModes.HasFlag(ProcessingMode.Reward),
                    runMergeMarkUpFees:  processingModes.HasFlag(ProcessingMode.MergeMarkUpFees),
                    runRewardFix:        processingModes.HasFlag(ProcessingMode.Reward),
                    rewardContractCondition: req.RewardContractCondition ?? "'New Reward Contract'",
                    installmentCondition:    req.InstallmentContractCondition,
                    excludeInstallment:      processingModes.HasFlag(ProcessingMode.Installment));

                var bulkResult = _root.BulkHandler.Handle(bulkCmd);
                // Maintenance failures are warnings only — generation continues.

                // ── Step 2: Resolve output root path ─────────────────────────
                string outputRoot = string.IsNullOrWhiteSpace(req.OutputRootPath)
                    ? _root.ConfigService.GetStatementOutputPath()
                    : req.OutputRootPath;

                // ── Step 3: Statement generation ─────────────────────────────
                var genCmd = new GenerateStatementCommand(
                    branchCode:                   req.BranchCode,
                    bankName:                     req.BankName,
                    bankFullName:                 req.BankFullName ?? req.BankName,
                    cardProduct:                  req.CardProduct ?? string.Empty,
                    outputType:                   (StatementOutputType)req.OutputType,
                    formatterKey:                 req.FormatterKey,
                    processingModes:              processingModes,
                    cardType:                     (CardType)req.CardType,
                    statementType:                (StatementType)req.StatementType,
                    statementDate:                req.StatementDate,
                    outputRootPath:               outputRoot,
                    appendMode:                   req.AppendMode,
                    rewardContractCondition:      req.RewardContractCondition      ?? "'New Reward Contract'",
                    installmentContractCondition: req.InstallmentContractCondition,
                    useCorporateAccountNumber:    req.UseCorporateAccountNumber,
                    statementTypeSuffix:          req.StatementTypeSuffix          ?? "CR",
                    reportTemplateName:           req.ReportTemplateName,
                    pdfPasswordScheme:            string.IsNullOrEmpty(req.PdfPasswordScheme) ? null : req.PdfPasswordScheme,
                    emailFromAddress:             req.EmailFromAddress             ?? "cardservices@emp-group.com",
                    bankWebsiteUrl:               req.BankWebsiteUrl               ?? "www.emp-group.com",
                    facebookUrl:                  req.FacebookUrl,
                    linkedInUrl:                  req.LinkedInUrl,
                    youTubeUrl:                   req.YouTubeUrl,
                    maxTransactionsPerPage:       req.MaxTransactionsPerPage  > 0 ? req.MaxTransactionsPerPage  : 20,
                    maxTransactionsLastPage:      req.MaxTransactionsLastPage > 0 ? req.MaxTransactionsLastPage : 27,
                    emailWaitPeriodMs:            req.EmailWaitPeriodMs       > 0 ? req.EmailWaitPeriodMs       : 7000,
                    fieldSeparator:               req.FieldSeparator          ?? ",");

                var handler = _root.CreateStatementHandler();
                var result  = handler.Handle(genCmd);

                if (result.IsSuccess)
                {
                    return Ok(new GenerationResultDto
                    {
                        Success          = true,
                        FilesGenerated   = result.Value.FilesGenerated,
                        EmailsSent       = result.Value.EmailsSent,
                        StatementsCount  = result.Value.StatementsCount,
                        TransactionCount = result.Value.TransactionCount,
                        OutputDirectory  = result.Value.OutputDirectory,
                    });
                }

                return Ok(new GenerationResultDto { Success = false, Error = result.Error });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new GenerationResultDto { Success = false, Error = ex.Message });
            }
        }

        /// <summary>
        /// Generates statements for multiple bank / product combinations in sequence.
        /// Returns a result per entry in the same order as the request list.
        /// </summary>
        [HttpPost("generate-bulk")]
        public ActionResult<List<GenerationResultDto>> GenerateBulk(
            [FromBody] List<GenerateStatementApiRequest> requests)
        {
            if (requests == null || requests.Count == 0)
                return BadRequest(new { error = "At least one request is required." });

            var results = new List<GenerationResultDto>(requests.Count);
            foreach (var req in requests)
            {
                // Reuse the single-generate logic by calling the repository directly.
                // We accept that each item runs sequentially (same as legacy BackgroundWorker).
                var singleResult = Generate(req);

                if (singleResult.Result is OkObjectResult ok && ok.Value is GenerationResultDto dto)
                    results.Add(dto);
                else if (singleResult.Result is ObjectResult objResult && objResult.Value is GenerationResultDto dtoErr)
                    results.Add(dtoErr);
                else
                    results.Add(new GenerationResultDto { Success = false, Error = "Unexpected error." });
            }

            return Ok(results);
        }
    }
}

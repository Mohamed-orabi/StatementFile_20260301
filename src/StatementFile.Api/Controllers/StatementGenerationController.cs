using System;
using Microsoft.AspNetCore.Mvc;
using StatementFile.Application.UseCases.StatementGeneration;
using StatementFile.Domain.Enums;
using StatementFile.Infrastructure.Configuration;

namespace StatementFile.Api.Controllers
{
    /// <summary>
    /// REST API for running statement generation.
    ///
    /// Routes
    ///   POST /api/statement-generation/run — generate statements for one bank/product config
    ///
    /// The request contains only ConfigId + StatementDate; the controller loads
    /// the full bank/product configuration from the database, builds the command,
    /// and delegates to GenerateStatementHandler (which runs pre-processing internally).
    /// </summary>
    [ApiController]
    [Route("api/statement-generation")]
    public sealed class StatementGenerationController : ControllerBase
    {
        private readonly CompositionRoot _root;

        public StatementGenerationController(CompositionRoot root)
        {
            _root = root ?? throw new ArgumentNullException(nameof(root));
        }

        // ── POST /api/statement-generation/run ────────────────────────────────

        [HttpPost("run")]
        public IActionResult Run([FromBody] StatementRunRequest req)
        {
            var cfg = _root.BankProductCfgRepo.GetById(req.ConfigId);
            if (cfg == null)
                return NotFound($"Bank/product configuration {req.ConfigId} not found.");

            string outputRootPath = _root.ConfigService.GetStatementOutputPath();

            var cmd = new GenerateStatementCommand(
                branchCode:                   cfg.BranchCode,
                bankName:                     cfg.BankName,
                bankFullName:                 cfg.BankFullName,
                cardProduct:                  cfg.CardProduct,
                outputType:                   cfg.OutputType,
                formatterKey:                 cfg.FormatterKey,
                processingModes:              cfg.ProcessingModes,
                cardType:                     cfg.CardType,
                statementType:                cfg.StatementType,
                statementDate:                req.StatementDate,
                outputRootPath:               outputRootPath,
                appendMode:                   cfg.AppendMode,
                rewardContractCondition:      cfg.RewardContractCondition,
                installmentContractCondition: cfg.InstallmentContractCondition,
                useCorporateAccountNumber:    cfg.UseCorporateAccountNumber,
                statementTypeSuffix:          cfg.StatementTypeSuffix,
                reportTemplateName:           cfg.ReportTemplateName,
                pdfPasswordScheme:            cfg.PdfPasswordScheme,
                emailFromAddress:             cfg.EmailFromAddress,
                bankWebsiteUrl:               cfg.BankWebsiteUrl,
                facebookUrl:                  cfg.FacebookUrl,
                linkedInUrl:                  cfg.LinkedInUrl,
                youTubeUrl:                   cfg.YouTubeUrl,
                maxTransactionsPerPage:       cfg.MaxTransactionsPerPage,
                maxTransactionsLastPage:      cfg.MaxTransactionsLastPage,
                emailWaitPeriodMs:            cfg.EmailWaitPeriodMs,
                fieldSeparator:               cfg.FieldSeparator);

            var handler = _root.CreateStatementHandler();
            var result  = handler.Handle(cmd);

            if (result.IsSuccess)
            {
                var v = result.Value;
                return Ok(new StatementRunResult
                {
                    Success          = true,
                    OutputDirectory  = v.OutputDirectory,
                    FilesGenerated   = v.FilesGenerated,
                    EmailsSent       = v.EmailsSent,
                    NoEmailCount     = v.NoEmailCount,
                    StatementsCount  = v.StatementsCount,
                    TransactionCount = v.TransactionCount,
                });
            }

            return Ok(new StatementRunResult
            {
                Success = false,
                Error   = result.Error,
            });
        }
    }
}

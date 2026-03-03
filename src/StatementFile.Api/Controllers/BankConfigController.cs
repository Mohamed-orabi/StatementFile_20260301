using System;
using System.Linq;
using Microsoft.AspNetCore.Mvc;
using StatementFile.Application.UseCases.BankConfiguration;
using StatementFile.Domain.Entities;
using StatementFile.Domain.Enums;
using StatementFile.Infrastructure.Configuration;

namespace StatementFile.Api.Controllers
{
    /// <summary>
    /// REST API for bank/product configuration records.
    ///
    /// Routes
    ///   GET    /api/bank-config                  — all rows (active+inactive)
    ///   GET    /api/bank-config?activeOnly=true   — active rows only, ordered by SortOrder
    ///   GET    /api/bank-config/{id}              — single row
    ///   POST   /api/bank-config                  — create
    ///   PUT    /api/bank-config/{id}              — update mutable fields
    ///   DELETE /api/bank-config/{id}              — delete
    /// </summary>
    [ApiController]
    [Route("api/bank-config")]
    public sealed class BankConfigController : ControllerBase
    {
        private readonly CompositionRoot _root;

        public BankConfigController(CompositionRoot root)
        {
            _root = root ?? throw new ArgumentNullException(nameof(root));
        }

        // ── GET /api/bank-config[?activeOnly=true] ────────────────────────────

        [HttpGet]
        public IActionResult GetAll([FromQuery] bool activeOnly = false)
        {
            var items = activeOnly
                ? _root.BankProductCfgRepo.GetActive()
                : _root.BankProductCfgRepo.GetAll();

            return Ok(items.Select(BankProductConfigDto.From));
        }

        // ── GET /api/bank-config/{id} ─────────────────────────────────────────

        [HttpGet("{id:int}")]
        public IActionResult GetById(int id)
        {
            var cfg = _root.BankProductCfgRepo.GetById(id);
            return cfg == null ? NotFound() : Ok(BankProductConfigDto.From(cfg));
        }

        // ── POST /api/bank-config ─────────────────────────────────────────────

        [HttpPost]
        public IActionResult Create([FromBody] SaveBankProductConfigRequest req)
        {
            if (string.IsNullOrWhiteSpace(req.BankName))
                return BadRequest("BankName is required.");
            if (string.IsNullOrWhiteSpace(req.FormatterKey))
                return BadRequest("FormatterKey is required.");
            if (req.ScheduledDay < 1 || req.ScheduledDay > 31)
                return BadRequest("ScheduledDay must be between 1 and 31.");

            var entity = BankProductConfig.Create(
                branchCode:                   req.BranchCode,
                bankName:                     req.BankName,
                bankFullName:                 req.BankFullName,
                cardProduct:                  req.CardProduct,
                outputType:                   (StatementOutputType)req.OutputType,
                formatterKey:                 req.FormatterKey,
                processingModes:              (ProcessingMode)req.ProcessingModes,
                cardType:                     (CardType)req.CardType,
                statementType:                (StatementType)req.StatementType,
                statementTypeSuffix:          req.StatementTypeSuffix,
                rewardContractCondition:      req.RewardContractCondition,
                installmentContractCondition: req.InstallmentContractCondition,
                appendMode:                   req.AppendMode,
                useCorporateAccountNumber:    req.UseCorporateAccountNumber,
                reportTemplateName:           req.ReportTemplateName,
                pdfPasswordScheme:            NullIfEmpty(req.PdfPasswordScheme),
                emailFromAddress:             req.EmailFromAddress,
                bankWebsiteUrl:               req.BankWebsiteUrl,
                facebookUrl:                  req.FacebookUrl,
                linkedInUrl:                  req.LinkedInUrl,
                youTubeUrl:                   req.YouTubeUrl,
                maxTransactionsPerPage:       req.MaxTransactionsPerPage,
                maxTransactionsLastPage:      req.MaxTransactionsLastPage,
                emailWaitPeriodMs:            req.EmailWaitPeriodMs,
                fieldSeparator:               req.FieldSeparator,
                scheduledDay:                 req.ScheduledDay,
                displayName:                  req.DisplayName,
                sortOrder:                    req.SortOrder,
                isActive:                     req.IsActive);

            int newId = _root.BankProductCfgRepo.Insert(entity);

            return CreatedAtAction(nameof(GetById), new { id = newId }, newId);
        }

        // ── PUT /api/bank-config/{id} ─────────────────────────────────────────

        [HttpPut("{id:int}")]
        public IActionResult Update(int id, [FromBody] SaveBankProductConfigRequest req)
        {
            var entity = _root.BankProductCfgRepo.GetById(id);
            if (entity == null) return NotFound();

            if (string.IsNullOrWhiteSpace(req.FormatterKey))
                return BadRequest("FormatterKey is required.");
            if (req.ScheduledDay < 1 || req.ScheduledDay > 31)
                return BadRequest("ScheduledDay must be between 1 and 31.");

            entity.Update(
                bankFullName:                 req.BankFullName,
                cardProduct:                  req.CardProduct,
                outputType:                   (StatementOutputType)req.OutputType,
                formatterKey:                 req.FormatterKey,
                processingModes:              (ProcessingMode)req.ProcessingModes,
                cardType:                     (CardType)req.CardType,
                statementType:                (StatementType)req.StatementType,
                statementTypeSuffix:          req.StatementTypeSuffix,
                rewardContractCondition:      req.RewardContractCondition,
                installmentContractCondition: req.InstallmentContractCondition,
                appendMode:                   req.AppendMode,
                useCorporateAccountNumber:    req.UseCorporateAccountNumber,
                reportTemplateName:           req.ReportTemplateName,
                pdfPasswordScheme:            NullIfEmpty(req.PdfPasswordScheme),
                emailFromAddress:             req.EmailFromAddress,
                bankWebsiteUrl:               req.BankWebsiteUrl,
                facebookUrl:                  req.FacebookUrl,
                linkedInUrl:                  req.LinkedInUrl,
                youTubeUrl:                   req.YouTubeUrl,
                maxTransactionsPerPage:       req.MaxTransactionsPerPage,
                maxTransactionsLastPage:      req.MaxTransactionsLastPage,
                emailWaitPeriodMs:            req.EmailWaitPeriodMs,
                fieldSeparator:               req.FieldSeparator,
                scheduledDay:                 req.ScheduledDay,
                displayName:                  req.DisplayName,
                sortOrder:                    req.SortOrder,
                isActive:                     req.IsActive);

            _root.BankProductCfgRepo.Update(entity);

            return NoContent();
        }

        // ── DELETE /api/bank-config/{id} ──────────────────────────────────────

        [HttpDelete("{id:int}")]
        public IActionResult Delete(int id)
        {
            if (_root.BankProductCfgRepo.GetById(id) == null)
                return NotFound();

            _root.BankProductCfgRepo.Delete(id);
            return NoContent();
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static string NullIfEmpty(string s)
            => string.IsNullOrEmpty(s) ? null : s;
    }
}

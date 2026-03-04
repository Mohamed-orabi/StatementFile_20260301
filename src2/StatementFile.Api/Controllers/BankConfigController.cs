using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.AspNetCore.Mvc;
using StatementFile.Application.DTOs;
using StatementFile.Domain.Entities;
using StatementFile.Domain.Enums;
using StatementFile.Infrastructure.Configuration;

namespace StatementFile.Api.Controllers
{
    /// <summary>
    /// CRUD operations for bank / product configuration records stored in
    /// the Oracle table STAT_BANK_PRODUCT_CONFIG.
    ///
    /// All mutations go through the domain entity's factory / update methods
    /// so business invariants (required fields, default values) are enforced
    /// regardless of which client is calling the API.
    /// </summary>
    [ApiController]
    [Route("api/bank-configs")]
    public sealed class BankConfigController : ControllerBase
    {
        private readonly CompositionRoot _root;

        public BankConfigController(CompositionRoot root) => _root = root;

        // ── Queries ───────────────────────────────────────────────────────────

        /// <summary>Returns all bank / product configurations (active + inactive).</summary>
        [HttpGet]
        public ActionResult<IReadOnlyList<BankConfigDto>> GetAll()
        {
            try
            {
                var list = _root.BankProductCfgRepo.GetAll();
                return Ok(list.Select(MapToDto).ToList());
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }

        /// <summary>Returns only active bank / product configurations.</summary>
        [HttpGet("active")]
        public ActionResult<IReadOnlyList<BankConfigDto>> GetActive()
        {
            try
            {
                var list = _root.BankProductCfgRepo.GetActive();
                return Ok(list.Select(MapToDto).ToList());
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }

        /// <summary>Returns a single configuration by primary key.</summary>
        [HttpGet("{id:int}")]
        public ActionResult<BankConfigDto> GetById(int id)
        {
            try
            {
                var entity = _root.BankProductCfgRepo.GetById(id);
                if (entity == null) return NotFound(new { error = $"Configuration #{id} not found." });
                return Ok(MapToDto(entity));
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }

        // ── Commands ──────────────────────────────────────────────────────────

        /// <summary>Creates a new bank / product configuration. Returns the new ID.</summary>
        [HttpPost]
        public ActionResult<int> Create([FromBody] BankConfigDto dto)
        {
            if (dto == null)
                return BadRequest(new { error = "Request body is required." });
            if (string.IsNullOrWhiteSpace(dto.BankName))
                return BadRequest(new { error = "BankName is required." });
            if (string.IsNullOrWhiteSpace(dto.FormatterKey))
                return BadRequest(new { error = "FormatterKey is required." });

            try
            {
                var entity = BankProductConfig.Create(
                    branchCode:                   dto.BranchCode,
                    bankName:                     dto.BankName,
                    bankFullName:                 dto.BankFullName,
                    cardProduct:                  dto.CardProduct,
                    outputType:                   (StatementOutputType)dto.OutputType,
                    formatterKey:                 dto.FormatterKey,
                    processingModes:              (ProcessingMode)dto.ProcessingModes,
                    cardType:                     (CardType)dto.CardType,
                    statementType:                (StatementType)dto.StatementType,
                    statementTypeSuffix:          dto.StatementTypeSuffix,
                    rewardContractCondition:      dto.RewardContractCondition,
                    installmentContractCondition: dto.InstallmentContractCondition,
                    appendMode:                   dto.AppendMode,
                    useCorporateAccountNumber:    dto.UseCorporateAccountNumber,
                    reportTemplateName:           dto.ReportTemplateName,
                    pdfPasswordScheme:            string.IsNullOrEmpty(dto.PdfPasswordScheme) ? null : dto.PdfPasswordScheme,
                    emailFromAddress:             dto.EmailFromAddress,
                    bankWebsiteUrl:               dto.BankWebsiteUrl,
                    facebookUrl:                  dto.FacebookUrl,
                    linkedInUrl:                  dto.LinkedInUrl,
                    youTubeUrl:                   dto.YouTubeUrl,
                    maxTransactionsPerPage:       dto.MaxTransactionsPerPage,
                    maxTransactionsLastPage:      dto.MaxTransactionsLastPage,
                    emailWaitPeriodMs:            dto.EmailWaitPeriodMs,
                    fieldSeparator:               dto.FieldSeparator,
                    scheduledDay:                 dto.ScheduledDay,
                    displayName:                  dto.DisplayName,
                    sortOrder:                    dto.SortOrder,
                    isActive:                     dto.IsActive);

                int newId = _root.BankProductCfgRepo.Insert(entity);
                return Ok(newId);
            }
            catch (Exception ex)
            {
                return BadRequest(new { error = ex.Message });
            }
        }

        /// <summary>Updates an existing bank / product configuration.</summary>
        [HttpPut("{id:int}")]
        public IActionResult Update(int id, [FromBody] BankConfigDto dto)
        {
            if (dto == null)
                return BadRequest(new { error = "Request body is required." });

            try
            {
                var entity = _root.BankProductCfgRepo.GetById(id);
                if (entity == null)
                    return NotFound(new { error = $"Configuration #{id} not found." });

                entity.Update(
                    bankFullName:                dto.BankFullName,
                    cardProduct:                 dto.CardProduct,
                    outputType:                  (StatementOutputType)dto.OutputType,
                    formatterKey:                dto.FormatterKey,
                    processingModes:             (ProcessingMode)dto.ProcessingModes,
                    cardType:                    (CardType)dto.CardType,
                    statementType:               (StatementType)dto.StatementType,
                    statementTypeSuffix:         dto.StatementTypeSuffix,
                    rewardContractCondition:     dto.RewardContractCondition,
                    installmentContractCondition:dto.InstallmentContractCondition,
                    appendMode:                  dto.AppendMode,
                    useCorporateAccountNumber:   dto.UseCorporateAccountNumber,
                    reportTemplateName:          dto.ReportTemplateName,
                    pdfPasswordScheme:           string.IsNullOrEmpty(dto.PdfPasswordScheme) ? null : dto.PdfPasswordScheme,
                    emailFromAddress:            dto.EmailFromAddress,
                    bankWebsiteUrl:              dto.BankWebsiteUrl,
                    facebookUrl:                 dto.FacebookUrl,
                    linkedInUrl:                 dto.LinkedInUrl,
                    youTubeUrl:                  dto.YouTubeUrl,
                    maxTransactionsPerPage:      dto.MaxTransactionsPerPage,
                    maxTransactionsLastPage:     dto.MaxTransactionsLastPage,
                    emailWaitPeriodMs:           dto.EmailWaitPeriodMs,
                    fieldSeparator:              dto.FieldSeparator,
                    scheduledDay:                dto.ScheduledDay,
                    displayName:                 dto.DisplayName,
                    sortOrder:                   dto.SortOrder,
                    isActive:                    dto.IsActive);

                _root.BankProductCfgRepo.Update(entity);
                return NoContent();
            }
            catch (Exception ex)
            {
                return BadRequest(new { error = ex.Message });
            }
        }

        /// <summary>Deletes a bank / product configuration permanently.</summary>
        [HttpDelete("{id:int}")]
        public IActionResult Delete(int id)
        {
            try
            {
                _root.BankProductCfgRepo.Delete(id);
                return NoContent();
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }

        // ── Mapping helper ────────────────────────────────────────────────────

        private static BankConfigDto MapToDto(BankProductConfig c) => new BankConfigDto
        {
            Id                           = c.Id,
            BranchCode                   = c.BranchCode,
            BankName                     = c.BankName,
            BankFullName                 = c.BankFullName,
            CardProduct                  = c.CardProduct,
            OutputType                   = (int)c.OutputType,
            FormatterKey                 = c.FormatterKey,
            ProcessingModes              = (long)c.ProcessingModes,
            CardType                     = (int)c.CardType,
            StatementType                = (int)c.StatementType,
            StatementTypeSuffix          = c.StatementTypeSuffix,
            RewardContractCondition      = c.RewardContractCondition,
            InstallmentContractCondition = c.InstallmentContractCondition,
            AppendMode                   = c.AppendMode,
            UseCorporateAccountNumber    = c.UseCorporateAccountNumber,
            ReportTemplateName           = c.ReportTemplateName,
            PdfPasswordScheme            = c.PdfPasswordScheme,
            EmailFromAddress             = c.EmailFromAddress,
            BankWebsiteUrl               = c.BankWebsiteUrl,
            FacebookUrl                  = c.FacebookUrl,
            LinkedInUrl                  = c.LinkedInUrl,
            YouTubeUrl                   = c.YouTubeUrl,
            MaxTransactionsPerPage       = c.MaxTransactionsPerPage,
            MaxTransactionsLastPage      = c.MaxTransactionsLastPage,
            EmailWaitPeriodMs            = c.EmailWaitPeriodMs,
            FieldSeparator               = c.FieldSeparator,
            ScheduledDay                 = c.ScheduledDay,
            DisplayName                  = c.DisplayName,
            SortOrder                    = c.SortOrder,
            IsActive                     = c.IsActive,
            CreatedAt                    = c.CreatedAt,
            ModifiedAt                   = c.ModifiedAt,
        };
    }
}

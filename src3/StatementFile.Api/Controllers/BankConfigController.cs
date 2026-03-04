using System;
using System.Linq;
using Microsoft.AspNetCore.Mvc;
using StatementFile.Application.DTOs;
using StatementFile.Domain.Entities;
using StatementFile.Domain.Enums;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Api.Controllers
{
    /// <summary>
    /// CRUD management for bank / product configurations.
    ///
    /// Each record replaces one "case" block in the legacy
    /// frmStatementFileExtn.runStatement() switch statement.
    ///
    /// Routes:
    ///   GET    /api/bank-configs              → all configs
    ///   GET    /api/bank-configs/active       → active only
    ///   GET    /api/bank-configs/{id}         → single config
    ///   POST   /api/bank-configs              → create
    ///   PUT    /api/bank-configs/{id}         → update
    ///   DELETE /api/bank-configs/{id}         → delete
    /// </summary>
    [ApiController]
    [Route("api/bank-configs")]
    public sealed class BankConfigController : ControllerBase
    {
        private readonly IBankProductConfigRepository _repo;
        private readonly IFormatterRegistry           _formatters;

        public BankConfigController(
            IBankProductConfigRepository repo,
            IFormatterRegistry           formatters)
        {
            _repo       = repo       ?? throw new ArgumentNullException(nameof(repo));
            _formatters = formatters ?? throw new ArgumentNullException(nameof(formatters));
        }

        // ── Queries ───────────────────────────────────────────────────────────────

        [HttpGet]
        public IActionResult GetAll()
        {
            var configs = _repo.GetAll().Select(MapToDto);
            return Ok(configs);
        }

        [HttpGet("active")]
        public IActionResult GetActive()
        {
            var configs = _repo.GetActive().Select(MapToDto);
            return Ok(configs);
        }

        [HttpGet("{id:int}")]
        public IActionResult GetById(int id)
        {
            var config = _repo.GetById(id);
            if (config == null) return NotFound();
            return Ok(MapToDto(config));
        }

        // ── Commands ──────────────────────────────────────────────────────────────

        [HttpPost]
        public IActionResult Create([FromBody] BankConfigDto dto)
        {
            if (dto == null)
                return BadRequest(new { error = "Request body is required." });

            if (!_formatters.IsRegistered(dto.FormatterKey))
                return BadRequest(new { error = $"Unknown formatter key '{dto.FormatterKey}'. Register the formatter in InfrastructureServiceExtensions first." });

            try
            {
                var config = BankProductConfig.Create(
                    bankName:                  dto.BankName,
                    bankFullName:              dto.BankFullName,
                    bankCode:                  dto.BankCode,
                    branchCode:                dto.BranchCode,
                    statementTypeSuffix:       dto.StatementTypeSuffix,
                    cardType:                  dto.CardType,
                    cardProduct:               dto.CardProduct,
                    outputType:                dto.OutputType,
                    formatterKey:              dto.FormatterKey,
                    bankWebLink:               dto.BankWebLink,
                    bankLogo:                  dto.BankLogo,
                    backgroundImage:           dto.BackgroundImage,
                    midBannerImage:            dto.MidBannerImage,
                    bottomBannerImage:         dto.BottomBannerImage,
                    emailFromAddress:          dto.EmailFromAddress,
                    emailFromName:             dto.EmailFromName,
                    whereCondition:            dto.WhereCondition,
                    vipCondition:              dto.VipCondition,
                    rewardCondition:           dto.RewardCondition,
                    rewardContractCondition:   dto.RewardContractCondition,
                    currencyFilter:            dto.CurrencyFilter,
                    installmentCondition:      dto.InstallmentCondition,
                    paymentSystem:             dto.PaymentSystem,
                    processingModes:           (ProcessingMode)dto.ProcessingModes,
                    isRewardRun:               dto.IsRewardRun,
                    isSplitOutput:             dto.IsSplitOutput,
                    hasAttachment:             dto.HasAttachment,
                    saveDataset:               dto.SaveDataset,
                    showMessageBox:            dto.ShowMessageBox,
                    runNullCardDelete:         dto.RunNullCardDelete,
                    runCardBranchMatch:        dto.RunCardBranchMatch,
                    excludeReward:             dto.ExcludeReward,
                    waitPeriodSeconds:         dto.WaitPeriodSeconds);

                int newId = _repo.Add(config);
                return Ok(newId);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(new { error = ex.Message });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }

        [HttpPut("{id:int}")]
        public IActionResult Update(int id, [FromBody] BankConfigDto dto)
        {
            if (dto == null)
                return BadRequest(new { error = "Request body is required." });

            var config = _repo.GetById(id);
            if (config == null) return NotFound();

            if (!_formatters.IsRegistered(dto.FormatterKey))
                return BadRequest(new { error = $"Unknown formatter key '{dto.FormatterKey}'." });

            try
            {
                config.Update(
                    bankName:                  dto.BankName,
                    bankFullName:              dto.BankFullName,
                    bankCode:                  dto.BankCode,
                    branchCode:                dto.BranchCode,
                    statementTypeSuffix:       dto.StatementTypeSuffix,
                    cardType:                  dto.CardType,
                    cardProduct:               dto.CardProduct,
                    outputType:                dto.OutputType,
                    formatterKey:              dto.FormatterKey,
                    bankWebLink:               dto.BankWebLink,
                    bankLogo:                  dto.BankLogo,
                    backgroundImage:           dto.BackgroundImage,
                    midBannerImage:            dto.MidBannerImage,
                    bottomBannerImage:         dto.BottomBannerImage,
                    emailFromAddress:          dto.EmailFromAddress,
                    emailFromName:             dto.EmailFromName,
                    whereCondition:            dto.WhereCondition,
                    vipCondition:              dto.VipCondition,
                    rewardCondition:           dto.RewardCondition,
                    rewardContractCondition:   dto.RewardContractCondition,
                    currencyFilter:            dto.CurrencyFilter,
                    installmentCondition:      dto.InstallmentCondition,
                    paymentSystem:             dto.PaymentSystem,
                    processingModes:           (ProcessingMode)dto.ProcessingModes,
                    isRewardRun:               dto.IsRewardRun,
                    isSplitOutput:             dto.IsSplitOutput,
                    hasAttachment:             dto.HasAttachment,
                    saveDataset:               dto.SaveDataset,
                    showMessageBox:            dto.ShowMessageBox,
                    runNullCardDelete:         dto.RunNullCardDelete,
                    runCardBranchMatch:        dto.RunCardBranchMatch,
                    excludeReward:             dto.ExcludeReward,
                    waitPeriodSeconds:         dto.WaitPeriodSeconds,
                    isActive:                  dto.IsActive);

                _repo.Update(config);
                return NoContent();
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }

        [HttpDelete("{id:int}")]
        public IActionResult Delete(int id)
        {
            var config = _repo.GetById(id);
            if (config == null) return NotFound();

            try
            {
                _repo.Delete(id);
                return NoContent();
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }

        // ── Mapper ────────────────────────────────────────────────────────────────

        private static BankConfigDto MapToDto(BankProductConfig c) => new BankConfigDto
        {
            Id                       = c.Id,
            IsActive                 = c.IsActive,
            BankName                 = c.BankName,
            BankFullName             = c.BankFullName,
            BankCode                 = c.BankCode,
            BranchCode               = c.BranchCode,
            StatementTypeSuffix      = c.StatementTypeSuffix,
            CardType                 = c.CardType,
            CardProduct              = c.CardProduct,
            OutputType               = c.OutputType,
            FormatterKey             = c.FormatterKey,
            BankWebLink              = c.BankWebLink,
            BankLogo                 = c.BankLogo,
            BackgroundImage          = c.BackgroundImage,
            MidBannerImage           = c.MidBannerImage,
            BottomBannerImage        = c.BottomBannerImage,
            EmailFromAddress         = c.EmailFromAddress,
            EmailFromName            = c.EmailFromName,
            WhereCondition           = c.WhereCondition,
            VipCondition             = c.VipCondition,
            RewardCondition          = c.RewardCondition,
            RewardContractCondition  = c.RewardContractCondition,
            CurrencyFilter           = c.CurrencyFilter,
            InstallmentCondition     = c.InstallmentCondition,
            PaymentSystem            = c.PaymentSystem,
            ProcessingModes          = (long)c.ProcessingModes,
            IsRewardRun              = c.IsRewardRun,
            IsSplitOutput            = c.IsSplitOutput,
            HasAttachment            = c.HasAttachment,
            SaveDataset              = c.SaveDataset,
            ShowMessageBox           = c.ShowMessageBox,
            RunNullCardDelete        = c.RunNullCardDelete,
            RunCardBranchMatch       = c.RunCardBranchMatch,
            ExcludeReward            = c.ExcludeReward,
            WaitPeriodSeconds        = c.WaitPeriodSeconds,
        };
    }
}

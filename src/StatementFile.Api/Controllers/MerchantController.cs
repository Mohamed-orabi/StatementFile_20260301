using System;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using StatementFile.Application.DTOs;
using StatementFile.Application.UseCases.MerchantOnboarding;
using StatementFile.Infrastructure.Configuration;

namespace StatementFile.Api.Controllers
{
    /// <summary>
    /// Processes a merchant XML statement file through the 7-step merchant onboarding flow:
    ///   1. Parse XML → DataSet
    ///   2. Load MDB template
    ///   3. Write DataSet → MDB
    ///   4. Export Crystal Report → PDF
    ///   5. Apply PDF password
    ///   6. Send email to bank
    ///   7. Clean up temp files
    ///
    /// The XML file is accepted as a Base-64-encoded string in the JSON body
    /// so clients don't need multipart / form-data upload infrastructure.
    /// </summary>
    [ApiController]
    [Route("api/merchants")]
    public sealed class MerchantController : ControllerBase
    {
        private readonly CompositionRoot _root;

        public MerchantController(CompositionRoot root) => _root = root;

        /// <summary>Process a merchant XML statement file.</summary>
        [HttpPost("process")]
        public ActionResult<GenerationResultDto> Process([FromBody] ProcessMerchantApiRequest req)
        {
            if (req == null)
                return BadRequest(new GenerationResultDto { Success = false, Error = "Request body is required." });
            if (string.IsNullOrWhiteSpace(req.BankName))
                return BadRequest(new GenerationResultDto { Success = false, Error = "BankName is required." });
            if (string.IsNullOrWhiteSpace(req.XmlContentBase64))
                return BadRequest(new GenerationResultDto { Success = false, Error = "XmlContentBase64 is required." });

            string tempXml = Path.GetTempFileName() + ".xml";
            try
            {
                // Decode Base-64 → XML bytes → temp file
                byte[] xmlBytes = Convert.FromBase64String(req.XmlContentBase64);
                System.IO.File.WriteAllBytes(tempXml, xmlBytes);

                var cmd = new ProcessMerchantStatementCommand(
                    xmlSourceFilePath: tempXml,
                    bankFullName:      req.BankFullName  ?? req.BankName,
                    bankName:          req.BankName,
                    bankCode:          req.BankCode      ?? req.BankName,
                    processingDate:    req.ProcessingDate);

                var result = _root.MerchantHandler.Handle(cmd);

                if (result.IsSuccess)
                    return Ok(new GenerationResultDto { Success = true });

                return Ok(new GenerationResultDto { Success = false, Error = result.Error });
            }
            catch (FormatException)
            {
                return BadRequest(new GenerationResultDto
                {
                    Success = false,
                    Error   = "XmlContentBase64 is not valid Base-64."
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new GenerationResultDto { Success = false, Error = ex.Message });
            }
            finally
            {
                if (System.IO.File.Exists(tempXml))
                    System.IO.File.Delete(tempXml);
            }
        }
    }
}

using System;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using StatementFile.Application.UseCases.MerchantOnboarding;
using StatementFile.Infrastructure.Configuration;

namespace StatementFile.Api.Controllers
{
    /// <summary>
    /// REST API for processing merchant XML statement files.
    ///
    /// Routes
    ///   POST /api/merchant-statement/process — process one merchant XML file
    ///
    /// The XML file content is sent as a base64-encoded string in the JSON body.
    /// The controller decodes it to a temporary file, runs the handler, then deletes the temp file.
    /// </summary>
    [ApiController]
    [Route("api/merchant-statement")]
    public sealed class MerchantStatementController : ControllerBase
    {
        private readonly CompositionRoot _root;

        public MerchantStatementController(CompositionRoot root)
        {
            _root = root ?? throw new ArgumentNullException(nameof(root));
        }

        // ── POST /api/merchant-statement/process ──────────────────────────────

        [HttpPost("process")]
        public IActionResult Process([FromBody] MerchantProcessRequest req)
        {
            if (string.IsNullOrWhiteSpace(req.BankName))
                return BadRequest("BankName is required.");
            if (string.IsNullOrWhiteSpace(req.XmlContentBase64))
                return BadRequest("XmlContentBase64 is required.");

            byte[] xmlBytes;
            try
            {
                xmlBytes = Convert.FromBase64String(req.XmlContentBase64);
            }
            catch (FormatException)
            {
                return BadRequest("XmlContentBase64 is not valid base64.");
            }

            string tempXml = Path.GetTempFileName() + ".xml";
            try
            {
                System.IO.File.WriteAllBytes(tempXml, xmlBytes);

                var cmd = new ProcessMerchantStatementCommand(
                    xmlSourceFilePath: tempXml,
                    bankFullName:      req.BankFullName,
                    bankName:          req.BankName,
                    bankCode:          req.BankCode,
                    processingDate:    req.ProcessingDate);

                var result = _root.MerchantHandler.Handle(cmd);

                return Ok(new MerchantProcessResult
                {
                    Success = result.IsSuccess,
                    Message = result.IsSuccess ? "Processed successfully." : result.Error,
                });
            }
            finally
            {
                if (System.IO.File.Exists(tempXml))
                    System.IO.File.Delete(tempXml);
            }
        }
    }
}

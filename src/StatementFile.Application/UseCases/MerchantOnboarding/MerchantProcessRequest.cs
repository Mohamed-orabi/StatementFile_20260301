using System;

namespace StatementFile.Application.UseCases.MerchantOnboarding
{
    /// <summary>
    /// Request body for POST /api/merchant-statement/process.
    /// The XML file content is base64-encoded so it can be sent as JSON.
    /// The API decodes it to a temp file and passes the path to the handler.
    /// </summary>
    public sealed class MerchantProcessRequest
    {
        public string   BankFullName     { get; set; }
        public string   BankName         { get; set; }
        public string   BankCode         { get; set; }
        public DateTime ProcessingDate   { get; set; }

        /// <summary>Base64-encoded content of the merchant XML file.</summary>
        public string   XmlContentBase64 { get; set; }
    }
}

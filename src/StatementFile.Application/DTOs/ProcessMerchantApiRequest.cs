using System;

namespace StatementFile.Application.DTOs
{
    /// <summary>
    /// API request for the merchant statement processing endpoint.
    ///
    /// The XML source file is transmitted as a Base-64 string so the
    /// endpoint can be called with a standard JSON body — no multipart
    /// upload infrastructure required.
    /// </summary>
    public sealed class ProcessMerchantApiRequest
    {
        /// <summary>Base-64-encoded content of the merchant XML file.</summary>
        public string XmlContentBase64 { get; set; }

        public string BankFullName     { get; set; } = string.Empty;
        public string BankName         { get; set; } = string.Empty;
        public string BankCode         { get; set; } = string.Empty;
        public DateTime ProcessingDate { get; set; } = DateTime.Today;
    }
}

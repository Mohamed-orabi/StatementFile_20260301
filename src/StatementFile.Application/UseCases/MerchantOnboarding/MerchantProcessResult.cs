namespace StatementFile.Application.UseCases.MerchantOnboarding
{
    /// <summary>
    /// Response body for POST /api/merchant-statement/process.
    /// </summary>
    public sealed class MerchantProcessResult
    {
        public bool   Success { get; set; }
        public string Message { get; set; }
    }
}

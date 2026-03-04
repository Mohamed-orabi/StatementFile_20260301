using System;
using System.Collections.Generic;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Infrastructure.Formatters
{
    /// <summary>
    /// Resolves a formatter by the string key stored in <c>BankProductConfig.FormatterKey</c>.
    ///
    /// Replaces the per-case object instantiation pattern in frmStatementFileExtn.runStatement():
    ///   case 5:  obj = new clsStatHtmlUBA();  ...  break;
    ///   case 12: obj = new clsStatement_ExportRpt(); ... break;
    ///
    /// Formatters are registered at startup via <see cref="Register"/> and then
    /// resolved at runtime via <see cref="Resolve"/>.  Each formatter key follows
    /// the convention "{OUTPUT_TYPE}_{BANK_MNEMONIC}", e.g.:
    ///   "HTML_UBA"       → HtmlStatementFormatter seeded for UBA
    ///   "PDF_BDCA"       → PdfStatementFormatter seeded for BDCA
    ///   "TXT_NSGB"       → TextStatementFormatter seeded for NSGB
    ///   "RAWDATA_EXPORT" → RawDataExportFormatter
    ///
    /// The set of registered formatters is open-closed: new banks are added by
    /// registering a new key + formatter instance at application startup in
    /// <see cref="StatementFile.Infrastructure.Configuration.InfrastructureServiceExtensions"/>,
    /// without modifying existing code.
    /// </summary>
    public sealed class FormatterRegistry : IFormatterRegistry
    {
        private readonly Dictionary<string, IStatementFormatter> _map
            = new Dictionary<string, IStatementFormatter>(StringComparer.OrdinalIgnoreCase);

        public void Register(string key, IStatementFormatter formatter)
        {
            if (string.IsNullOrWhiteSpace(key))
                throw new ArgumentException("Formatter key is required.", nameof(key));
            if (formatter == null)
                throw new ArgumentNullException(nameof(formatter));

            _map[key] = formatter;
        }

        public IStatementFormatter Resolve(string formatterKey)
        {
            if (string.IsNullOrWhiteSpace(formatterKey)) return null;
            return _map.TryGetValue(formatterKey, out var f) ? f : null;
        }

        public bool IsRegistered(string formatterKey) =>
            !string.IsNullOrWhiteSpace(formatterKey) && _map.ContainsKey(formatterKey);

        public IReadOnlyCollection<string> RegisteredKeys => _map.Keys;
    }
}

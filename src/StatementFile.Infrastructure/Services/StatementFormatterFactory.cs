using System;
using System.Collections.Generic;
using StatementFile.Application.Interfaces;

namespace StatementFile.Infrastructure.Services
{
    /// <summary>
    /// Registry-based factory that resolves <see cref="IStatementFormatterService"/>
    /// implementations by a string key (e.g. "HTML_BAI", "TXT_UBA").
    ///
    /// The key is stored alongside the product configuration (database-driven or
    /// appsettings), so new bank/format combinations can be registered without
    /// changing any if/else logic — satisfying the Open/Closed Principle.
    ///
    /// Registration happens at application startup via DI (see DependencyInjection.cs).
    /// </summary>
    public sealed class StatementFormatterFactory : IStatementFormatterFactory
    {
        private readonly Dictionary<string, IStatementFormatterService> _registry;
        private readonly IStatementFormatterService                      _fallback;

        public StatementFormatterFactory(
            IEnumerable<IStatementFormatterService> formatters,
            IStatementFormatterService              fallback)
        {
            if (formatters == null) throw new ArgumentNullException(nameof(formatters));
            _fallback = fallback ?? throw new ArgumentNullException(nameof(fallback));

            _registry = new Dictionary<string, IStatementFormatterService>(
                StringComparer.OrdinalIgnoreCase);

            foreach (var f in formatters)
                _registry[f.FormatterKey] = f;
        }

        public IStatementFormatterService GetFormatter(string formatterKey)
        {
            if (string.IsNullOrWhiteSpace(formatterKey))
                return _fallback;

            return _registry.TryGetValue(formatterKey, out var formatter)
                ? formatter
                : _fallback;
        }
    }
}

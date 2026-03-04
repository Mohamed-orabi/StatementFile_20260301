# StatementFile — Clean Architecture

## Solution Structure

```
src/
├── StatementFile.Domain/               ← Zero external dependencies
│   ├── Entities/
│   │   ├── Statement.cs                ← Aggregate root (60+ fields from TSTATEMENTMASTERTABLE)
│   │   ├── StatementTransaction.cs     ← Value object (23+ fields from TSTATEMENTDETAILTABLE)
│   │   ├── MerchantStatement.cs        ← Merchant aggregate root (XML → MDB flow)
│   │   ├── MerchantOperation.cs        ← Merchant transaction line
│   │   └── Bank.cs                     ← Branch reference data
│   ├── Enums/
│   │   ├── StatementType.cs            ← Normal | Corporate | Merchant
│   │   ├── CardType.cs                 ← Credit | Debit | Prepaid | Reward
│   │   └── ProcessingStatus.cs
│   ├── Interfaces/
│   │   ├── IUnitOfWork.cs              ← Oracle connection scope
│   │   ├── Repositories/
│   │   │   ├── IStatementRepository.cs
│   │   │   ├── IMerchantStatementRepository.cs
│   │   │   └── IBankRepository.cs
│   │   └── Services/
│   │       ├── IEmailService.cs
│   │       ├── IFtpService.cs
│   │       ├── IReportService.cs
│   │       ├── IStatementFormatterService.cs
│   │       └── IDataMaintenanceService.cs
│   └── Common/
│       └── Result.cs                   ← Discriminated-union (no exception-driven flow)
│
├── StatementFile.Application/          ← Orchestration; depends on Domain only
│   ├── DTOs/
│   │   ├── StatementDto.cs
│   │   ├── StatementTransactionDto.cs
│   │   ├── MerchantStatementDto.cs
│   │   ├── BankDto.cs
│   │   └── EmailRecipientDto.cs
│   ├── Interfaces/
│   │   ├── IConfigurationService.cs    ← Config abstraction (no System.Configuration here)
│   │   └── IStatementFormatterFactory.cs
│   ├── UseCases/
│   │   ├── MerchantOnboarding/
│   │   │   ├── ProcessMerchantStatementCommand.cs
│   │   │   ├── ProcessMerchantStatementHandler.cs  ← 7-step merchant flow
│   │   │   └── MerchantStatementXmlMapper.cs
│   │   ├── BulkProcessing/
│   │   │   ├── RunBulkMaintenanceCommand.cs
│   │   │   └── RunBulkMaintenanceHandler.cs        ← 3-step pre-processing
│   │   └── StatementGeneration/
│   │       ├── GenerateStatementCommand.cs
│   │       └── GenerateStatementHandler.cs         ← 6-step generation flow
│   └── Mapping/
│       └── StatementMappingProfile.cs              ← Entity ↔ DTO (no AutoMapper)
│
├── StatementFile.Infrastructure/       ← All I/O; depends on Domain + Application
│   ├── Data/
│   │   ├── OracleConnectionFactory.cs
│   │   ├── UnitOfWork.cs               ← Shares one OracleConnection per scope
│   │   └── Repositories/
│   │       ├── StatementRepository.cs  ← Oracle ODP.NET implementation
│   │       ├── BankRepository.cs
│   │       └── MerchantStatementRepository.cs  ← OleDb / MS-Access implementation
│   ├── Services/
│   │   ├── DataMaintenanceService.cs   ← Batch PL/SQL, card-branch match, Arabic fix
│   │   ├── EmailService.cs             ← System.Net.Mail / SMTP
│   │   ├── FtpService.cs               ← FtpWebRequest, Live-Banks/{bank}/To/
│   │   ├── ReportService.cs            ← Crystal Reports export
│   │   └── StatementFormatterFactory.cs ← Registry-based; add bank formatters here
│   ├── Formatters/
│   │   └── GenericHtmlStatementFormatter.cs    ← Fallback formatter
│   └── Configuration/
│       ├── AppConfigurationService.cs  ← Reads App.config + JSON files
│       └── DependencyInjection.cs      ← Static composition root
│
└── StatementFile.Presentation/         ← Windows Forms; depends on all layers
    ├── Program.cs                      ← STAThread entry; builds CompositionRoot
    ├── App.config                      ← All environment settings here
    └── Forms/
        ├── frmLogin.cs                 ← Oracle credential validation via DI
        └── frmGenerateStatement.cs     ← Product selection + BackgroundWorker
```

---

## Key Design Decisions

### 1. Domain Purity
`StatementFile.Domain` has **no NuGet references** — only BCL types.
Oracle, Crystal Reports, JSON, and Windows Forms are all external concerns that live
in Infrastructure or Presentation.

### 2. Database-Driven / Configurable Behaviour
| Legacy hardcoded | Clean Architecture equivalent |
|---|---|
| `if (bankName == "QNB_ALAHLI")` filter valid flag | `AppConfigurationService.ValidatedEmailBanks` HashSet (configurable) |
| `if (bankName == "SSB")` routing email | `ResolveBankSpecificTo()` — extend via config table |
| Table name string concatenation | `SessionContext.MainTable` / `DetailTable` (set per run from UI) |
| Crystal Reports template path hardcoded | `GetReportTemplate()` convention + App.config override |
| MDB template path hardcoded | `App.config["MerchantMdbTemplatePath"]` |

### 3. Open/Closed Formatter Registration
New bank/format combinations are added by:
1. Implementing `IStatementFormatterService` with a unique `FormatterKey`.
2. Adding the instance to the `formatters[]` array in `DependencyInjection.Compose()`.
No existing code needs changing.

### 4. Repository Pattern + Unit of Work
One `OracleConnection` is opened per statement run and shared across all repositories
in that scope via `IUnitOfWork`. The connection is disposed when the handler finishes.

### 5. Result<T> Pattern
Use-case handlers return `Result` / `Result<T>` instead of throwing exceptions for
expected business errors. The UI inspects `IsSuccess` / `Error` and shows appropriate
messages without try/catch at every call site.

### 6. Business Logic Preservation
The following flows are structurally identical to the legacy implementation:
- **Merchant Onboarding**: XML → DataSet → MDB → Crystal PDF → Email (7 steps)
- **Bulk Processing**: NULL-card delete → card-branch match → Arabic fix (3 steps)
- **Statement Generation**: load Oracle → format → export Crystal → FTP + Email (6 steps)
All field names, SQL hints (`/*+ parallel */`), filter conditions, and email routing
are preserved verbatim.

---

## Adding a New Bank-Specific Formatter

```csharp
// 1. Implement the interface
public sealed class BaiHtmlStatementFormatter : IStatementFormatterService
{
    public string FormatterKey => "HTML_BAI";

    public IEnumerable<string> Format(DataSet ds, string dir, int branch, string product)
    {
        // Bank-specific HTML generation logic (mirrors legacy clsStatHtmlBAI)
        ...
    }
}

// 2. Register in DependencyInjection.Compose()
var formatters = new IStatementFormatterService[]
{
    genericFormatter,
    new BaiHtmlStatementFormatter(),   // ← add here only
};
```

The product's `FormatterKey` field in the database / configuration routes each
product to its formatter at runtime — no if/else chain required.

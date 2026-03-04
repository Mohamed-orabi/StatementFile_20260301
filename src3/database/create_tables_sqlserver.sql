-- =============================================================================
-- StatementFile src3 – SQL Server DDL
-- =============================================================================
-- Run this script once against your SQL Server database, OR let EF Core
-- create the schema automatically by running:
--
--   cd src3/StatementFile.Api
--   dotnet ef migrations add InitialCreate --project ../StatementFile.Infrastructure
--   dotnet ef database update
-- =============================================================================

-- ---------------------------------------------------------------------------
-- STAT_BANK_PRODUCT_CONFIG
-- ---------------------------------------------------------------------------
CREATE TABLE [dbo].[STAT_BANK_PRODUCT_CONFIG]
(
    [ID]                         INT IDENTITY(1,1)   NOT NULL PRIMARY KEY,
    [IS_ACTIVE]                  BIT                 NOT NULL DEFAULT 1,

    -- Bank identification
    [BANK_NAME]                  NVARCHAR(50)        NOT NULL,
    [BANK_FULL_NAME]             NVARCHAR(200)           NULL,
    [BANK_CODE]                  NVARCHAR(20)            NULL,
    [BRANCH_CODE]                INT                 NOT NULL,

    -- Product / statement type
    [STATEMENT_TYPE_SUFFIX]      NVARCHAR(20)            NULL,
    [CARD_TYPE]                  TINYINT             NOT NULL,   -- 1=Credit 2=Debit 3=Corporate
    [CARD_PRODUCT]               NVARCHAR(100)           NULL,
    [OUTPUT_TYPE]                TINYINT             NOT NULL,   -- 1=Html 2=Pdf 3=Text 4=Raw 5=Xml 6=Email
    [FORMATTER_KEY]              NVARCHAR(100)       NOT NULL,

    -- Branding
    [BANK_WEB_LINK]              NVARCHAR(500)           NULL,
    [BANK_LOGO]                  NVARCHAR(500)           NULL,
    [BACKGROUND_IMAGE]           NVARCHAR(500)           NULL,
    [MID_BANNER_IMAGE]           NVARCHAR(500)           NULL,
    [BOTTOM_BANNER_IMAGE]        NVARCHAR(500)           NULL,

    -- Email
    [EMAIL_FROM_ADDRESS]         NVARCHAR(200)           NULL,
    [EMAIL_FROM_NAME]            NVARCHAR(200)           NULL,

    -- Oracle/SQL query filters
    [WHERE_CONDITION]            NVARCHAR(2000)          NULL,
    [VIP_CONDITION]              NVARCHAR(2000)          NULL,
    [REWARD_CONDITION]           NVARCHAR(2000)          NULL,
    [REWARD_CONTRACT_CONDITION]  NVARCHAR(2000)          NULL DEFAULT ('''New Reward Contract'''),
    [CURRENCY_FILTER]            NVARCHAR(2000)          NULL,
    [INSTALLMENT_CONDITION]      NVARCHAR(2000)          NULL,
    [PAYMENT_SYSTEM]             NVARCHAR(50)            NULL,

    -- Processing flags
    [PROCESSING_MODES]           BIGINT              NOT NULL DEFAULT 0,
    [IS_REWARD_RUN]              BIT                 NOT NULL DEFAULT 0,
    [IS_SPLIT_OUTPUT]            BIT                 NOT NULL DEFAULT 0,
    [HAS_ATTACHMENT]             BIT                 NOT NULL DEFAULT 0,
    [SAVE_DATASET]               BIT                 NOT NULL DEFAULT 0,
    [SHOW_MESSAGE_BOX]           BIT                 NOT NULL DEFAULT 0,
    [RUN_NULL_CARD_DELETE]       BIT                 NOT NULL DEFAULT 1,
    [RUN_CARD_BRANCH_MATCH]      BIT                 NOT NULL DEFAULT 1,
    [EXCLUDE_REWARD]             BIT                 NOT NULL DEFAULT 1,
    [WAIT_PERIOD_SECONDS]        INT                 NOT NULL DEFAULT 0,

    -- Metadata
    [CREATED_AT]                 DATETIME2           NOT NULL DEFAULT SYSUTCDATETIME(),
    [UPDATED_AT]                 DATETIME2           NOT NULL DEFAULT SYSUTCDATETIME()
);

-- ---------------------------------------------------------------------------
-- STAT_STATEMENT_RUN
-- ---------------------------------------------------------------------------
CREATE TABLE [dbo].[STAT_STATEMENT_RUN]
(
    [ID]               INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    [CONFIG_ID]        INT               NOT NULL REFERENCES [STAT_BANK_PRODUCT_CONFIG]([ID]),
    [STATEMENT_DATE]   DATE              NOT NULL,
    [STARTED_AT]       DATETIME2         NOT NULL DEFAULT SYSUTCDATETIME(),
    [FINISHED_AT]      DATETIME2             NULL,
    [IS_SUCCESS]       BIT               NOT NULL DEFAULT 0,
    [ERROR_MESSAGE]    NVARCHAR(4000)        NULL,
    [FILES_GENERATED]  INT               NOT NULL DEFAULT 0,
    [EMAILS_SENT]      INT               NOT NULL DEFAULT 0,
    [STATEMENTS_COUNT] INT               NOT NULL DEFAULT 0,
    [OUTPUT_DIRECTORY] NVARCHAR(1000)        NULL
);

-- ---------------------------------------------------------------------------
-- Stored procedures called by SqlBulkMaintenanceService
-- (Implement these bodies as needed for your environment)
-- ---------------------------------------------------------------------------
GO
CREATE OR ALTER PROCEDURE [dbo].[usp_DeleteNullCardRecords]
    @branchCode INT
AS
BEGIN
    SET NOCOUNT ON;
    -- TODO: implement based on legacy Oracle proc STMT.ZM_STMT_APP.DeleteNullCardRecords
    PRINT 'usp_DeleteNullCardRecords executed for branch ' + CAST(@branchCode AS NVARCHAR);
END
GO

CREATE OR ALTER PROCEDURE [dbo].[usp_MatchCardBranch]         @branchCode INT AS BEGIN SET NOCOUNT ON; END GO
CREATE OR ALTER PROCEDURE [dbo].[usp_FixArabicAddress]        @branchCode INT AS BEGIN SET NOCOUNT ON; END GO
CREATE OR ALTER PROCEDURE [dbo].[usp_FixArabicAddressLanguage]@branchCode INT AS BEGIN SET NOCOUNT ON; END GO
CREATE OR ALTER PROCEDURE [dbo].[usp_FixAddress]              @branchCode INT AS BEGIN SET NOCOUNT ON; END GO
CREATE OR ALTER PROCEDURE [dbo].[usp_DeleteOnHoldRecords]     @branchCode INT AS BEGIN SET NOCOUNT ON; END GO
CREATE OR ALTER PROCEDURE [dbo].[usp_MergeMarkUpFees]         @branchCode INT AS BEGIN SET NOCOUNT ON; END GO

CREATE OR ALTER PROCEDURE [dbo].[usp_ProcessRewardData]
    @branchCode INT, @rewardContractCondition NVARCHAR(2000)
AS BEGIN SET NOCOUNT ON; END
GO

CREATE OR ALTER PROCEDURE [dbo].[usp_ExcludeInstallmentData]
    @branchCode INT, @installmentCondition NVARCHAR(2000)
AS BEGIN SET NOCOUNT ON; END
GO

-- ---------------------------------------------------------------------------
-- Sample seed data
-- ---------------------------------------------------------------------------
INSERT INTO [dbo].[STAT_BANK_PRODUCT_CONFIG]
    (BANK_NAME, BANK_FULL_NAME, BANK_CODE, BRANCH_CODE, STATEMENT_TYPE_SUFFIX,
     CARD_TYPE, CARD_PRODUCT, OUTPUT_TYPE, FORMATTER_KEY,
     EMAIL_FROM_ADDRESS, EMAIL_FROM_NAME)
VALUES
    ('UBA', 'United Bank for Africa', 'UBA01', 1, 'CR',
     1, 'UBA Credit Card', 1, 'HTML_UBA',
     'cardservices@emp-group.com', 'UBA Card Services');
GO

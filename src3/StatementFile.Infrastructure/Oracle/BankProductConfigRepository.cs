using System;
using System.Collections.Generic;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using StatementFile.Domain.Entities;
using StatementFile.Domain.Enums;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Infrastructure.Oracle
{
    /// <summary>
    /// Persists <see cref="BankProductConfig"/> rows in the Oracle table
    /// STAT_BANK_PRODUCT_CONFIG — the database-driven replacement for the
    /// 500-case switch statement in frmStatementFileExtn.runStatement().
    /// </summary>
    public sealed class BankProductConfigRepository : IBankProductConfigRepository
    {
        private readonly IDbConnectionFactory _factory;

        public BankProductConfigRepository(IDbConnectionFactory factory) =>
            _factory = factory ?? throw new ArgumentNullException(nameof(factory));

        // ── Read ─────────────────────────────────────────────────────────────────

        public IReadOnlyList<BankProductConfig> GetAll()
        {
            const string sql = @"
                SELECT * FROM STAT_BANK_PRODUCT_CONFIG
                ORDER BY BANK_NAME, STATEMENT_TYPE_SUFFIX";

            return Query(sql);
        }

        public IReadOnlyList<BankProductConfig> GetActive()
        {
            const string sql = @"
                SELECT * FROM STAT_BANK_PRODUCT_CONFIG
                WHERE IS_ACTIVE = 1
                ORDER BY BANK_NAME, STATEMENT_TYPE_SUFFIX";

            return Query(sql);
        }

        public BankProductConfig GetById(int id)
        {
            const string sql = @"
                SELECT * FROM STAT_BANK_PRODUCT_CONFIG
                WHERE ID = :id";

            using var conn = _factory.CreateConnection();
            using var cmd  = conn.CreateCommand();
            cmd.CommandText = sql;
            AddParam(cmd, ":id", id);

            using var reader = cmd.ExecuteReader();
            return reader.Read() ? Map(reader) : null;
        }

        // ── Write ─────────────────────────────────────────────────────────────────

        public int Add(BankProductConfig c)
        {
            const string sql = @"
                INSERT INTO STAT_BANK_PRODUCT_CONFIG (
                    IS_ACTIVE, BANK_NAME, BANK_FULL_NAME, BANK_CODE, BRANCH_CODE,
                    STATEMENT_TYPE_SUFFIX, CARD_TYPE, CARD_PRODUCT,
                    OUTPUT_TYPE, FORMATTER_KEY,
                    BANK_WEB_LINK, BANK_LOGO, BACKGROUND_IMAGE, MID_BANNER_IMAGE, BOTTOM_BANNER_IMAGE,
                    EMAIL_FROM_ADDRESS, EMAIL_FROM_NAME,
                    WHERE_CONDITION, VIP_CONDITION, REWARD_CONDITION, REWARD_CONTRACT_CONDITION,
                    CURRENCY_FILTER, INSTALLMENT_CONDITION, PAYMENT_SYSTEM,
                    PROCESSING_MODES, IS_REWARD_RUN, IS_SPLIT_OUTPUT, HAS_ATTACHMENT,
                    SAVE_DATASET, SHOW_MESSAGE_BOX, RUN_NULL_CARD_DELETE,
                    RUN_CARD_BRANCH_MATCH, EXCLUDE_REWARD, WAIT_PERIOD_SECONDS,
                    CREATED_AT, UPDATED_AT
                ) VALUES (
                    :isActive, :bankName, :bankFullName, :bankCode, :branchCode,
                    :stmtTypeSuffix, :cardType, :cardProduct,
                    :outputType, :formatterKey,
                    :bankWebLink, :bankLogo, :bgImage, :midBanner, :bottomBanner,
                    :emailFrom, :emailFromName,
                    :whereCond, :vipCond, :rewardCond, :rewardContractCond,
                    :currFilter, :installCond, :paySystem,
                    :procModes, :isReward, :isSplit, :hasAttach,
                    :saveDs, :showMsg, :nullCardDel,
                    :cardBranchMatch, :exclReward, :waitPeriod,
                    :createdAt, :updatedAt
                )
                RETURNING ID INTO :newId";

            using var conn = _factory.CreateConnection();
            using var cmd  = conn.CreateCommand();
            cmd.CommandText = sql;

            BindConfigParams(cmd, c);

            var pNewId = ((OracleCommand)cmd).Parameters.Add(":newId", OracleDbType.Int32);
            pNewId.Direction = ParameterDirection.Output;

            cmd.ExecuteNonQuery();
            return Convert.ToInt32(pNewId.Value.ToString());
        }

        public void Update(BankProductConfig c)
        {
            const string sql = @"
                UPDATE STAT_BANK_PRODUCT_CONFIG SET
                    IS_ACTIVE                 = :isActive,
                    BANK_NAME                 = :bankName,
                    BANK_FULL_NAME            = :bankFullName,
                    BANK_CODE                 = :bankCode,
                    BRANCH_CODE               = :branchCode,
                    STATEMENT_TYPE_SUFFIX     = :stmtTypeSuffix,
                    CARD_TYPE                 = :cardType,
                    CARD_PRODUCT              = :cardProduct,
                    OUTPUT_TYPE               = :outputType,
                    FORMATTER_KEY             = :formatterKey,
                    BANK_WEB_LINK             = :bankWebLink,
                    BANK_LOGO                 = :bankLogo,
                    BACKGROUND_IMAGE          = :bgImage,
                    MID_BANNER_IMAGE          = :midBanner,
                    BOTTOM_BANNER_IMAGE       = :bottomBanner,
                    EMAIL_FROM_ADDRESS        = :emailFrom,
                    EMAIL_FROM_NAME           = :emailFromName,
                    WHERE_CONDITION           = :whereCond,
                    VIP_CONDITION             = :vipCond,
                    REWARD_CONDITION          = :rewardCond,
                    REWARD_CONTRACT_CONDITION = :rewardContractCond,
                    CURRENCY_FILTER           = :currFilter,
                    INSTALLMENT_CONDITION     = :installCond,
                    PAYMENT_SYSTEM            = :paySystem,
                    PROCESSING_MODES          = :procModes,
                    IS_REWARD_RUN             = :isReward,
                    IS_SPLIT_OUTPUT           = :isSplit,
                    HAS_ATTACHMENT            = :hasAttach,
                    SAVE_DATASET              = :saveDs,
                    SHOW_MESSAGE_BOX          = :showMsg,
                    RUN_NULL_CARD_DELETE      = :nullCardDel,
                    RUN_CARD_BRANCH_MATCH     = :cardBranchMatch,
                    EXCLUDE_REWARD            = :exclReward,
                    WAIT_PERIOD_SECONDS       = :waitPeriod,
                    UPDATED_AT                = :updatedAt
                WHERE ID = :id";

            using var conn = _factory.CreateConnection();
            using var cmd  = conn.CreateCommand();
            cmd.CommandText = sql;

            BindConfigParams(cmd, c);
            AddParam(cmd, ":id", c.Id);

            cmd.ExecuteNonQuery();
        }

        public void Delete(int id)
        {
            using var conn = _factory.CreateConnection();
            using var cmd  = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM STAT_BANK_PRODUCT_CONFIG WHERE ID = :id";
            AddParam(cmd, ":id", id);
            cmd.ExecuteNonQuery();
        }

        // ── Helpers ──────────────────────────────────────────────────────────────

        private IReadOnlyList<BankProductConfig> Query(string sql)
        {
            var list = new List<BankProductConfig>();
            using var conn   = _factory.CreateConnection();
            using var cmd    = conn.CreateCommand();
            cmd.CommandText  = sql;
            using var reader = cmd.ExecuteReader();
            while (reader.Read()) list.Add(Map(reader));
            return list;
        }

        private static BankProductConfig Map(IDataReader r)
        {
            var config = BankProductConfig.Create(
                bankName:                  r.GetString(r.GetOrdinal("BANK_NAME")),
                bankFullName:              SafeString(r, "BANK_FULL_NAME"),
                bankCode:                  SafeString(r, "BANK_CODE"),
                branchCode:                r.GetInt32(r.GetOrdinal("BRANCH_CODE")),
                statementTypeSuffix:       SafeString(r, "STATEMENT_TYPE_SUFFIX"),
                cardType:                  (CardType)r.GetInt32(r.GetOrdinal("CARD_TYPE")),
                cardProduct:               SafeString(r, "CARD_PRODUCT"),
                outputType:                (StatementOutputType)r.GetInt32(r.GetOrdinal("OUTPUT_TYPE")),
                formatterKey:              r.GetString(r.GetOrdinal("FORMATTER_KEY")),
                bankWebLink:               SafeString(r, "BANK_WEB_LINK"),
                bankLogo:                  SafeString(r, "BANK_LOGO"),
                backgroundImage:           SafeString(r, "BACKGROUND_IMAGE"),
                midBannerImage:            SafeString(r, "MID_BANNER_IMAGE"),
                bottomBannerImage:         SafeString(r, "BOTTOM_BANNER_IMAGE"),
                emailFromAddress:          SafeString(r, "EMAIL_FROM_ADDRESS"),
                emailFromName:             SafeString(r, "EMAIL_FROM_NAME"),
                whereCondition:            SafeString(r, "WHERE_CONDITION"),
                vipCondition:              SafeString(r, "VIP_CONDITION"),
                rewardCondition:           SafeString(r, "REWARD_CONDITION"),
                rewardContractCondition:   SafeString(r, "REWARD_CONTRACT_CONDITION"),
                currencyFilter:            SafeString(r, "CURRENCY_FILTER"),
                installmentCondition:      SafeString(r, "INSTALLMENT_CONDITION"),
                paymentSystem:             SafeString(r, "PAYMENT_SYSTEM"),
                processingModes:           (ProcessingMode)r.GetInt64(r.GetOrdinal("PROCESSING_MODES")),
                isRewardRun:               r.GetInt32(r.GetOrdinal("IS_REWARD_RUN")) == 1,
                isSplitOutput:             r.GetInt32(r.GetOrdinal("IS_SPLIT_OUTPUT")) == 1,
                hasAttachment:             r.GetInt32(r.GetOrdinal("HAS_ATTACHMENT")) == 1,
                saveDataset:               r.GetInt32(r.GetOrdinal("SAVE_DATASET")) == 1,
                showMessageBox:            r.GetInt32(r.GetOrdinal("SHOW_MESSAGE_BOX")) == 1,
                runNullCardDelete:         r.GetInt32(r.GetOrdinal("RUN_NULL_CARD_DELETE")) == 1,
                runCardBranchMatch:        r.GetInt32(r.GetOrdinal("RUN_CARD_BRANCH_MATCH")) == 1,
                excludeReward:             r.GetInt32(r.GetOrdinal("EXCLUDE_REWARD")) == 1,
                waitPeriodSeconds:         r.GetInt32(r.GetOrdinal("WAIT_PERIOD_SECONDS")));

            // Reflect the DB ID back onto the entity (private setter)
            typeof(BankProductConfig)
                .GetProperty(nameof(BankProductConfig.Id))!
                .SetValue(config, r.GetInt32(r.GetOrdinal("ID")));

            var isActive = r.GetInt32(r.GetOrdinal("IS_ACTIVE")) == 1;
            if (!isActive) config.Deactivate();

            return config;
        }

        private static void BindConfigParams(IDbCommand cmd, BankProductConfig c)
        {
            AddParam(cmd, ":isActive",            c.IsActive ? 1 : 0);
            AddParam(cmd, ":bankName",             c.BankName);
            AddParam(cmd, ":bankFullName",         c.BankFullName);
            AddParam(cmd, ":bankCode",             (object)c.BankCode ?? DBNull.Value);
            AddParam(cmd, ":branchCode",           c.BranchCode);
            AddParam(cmd, ":stmtTypeSuffix",       (object)c.StatementTypeSuffix ?? DBNull.Value);
            AddParam(cmd, ":cardType",             (int)c.CardType);
            AddParam(cmd, ":cardProduct",          (object)c.CardProduct ?? DBNull.Value);
            AddParam(cmd, ":outputType",           (int)c.OutputType);
            AddParam(cmd, ":formatterKey",         c.FormatterKey);
            AddParam(cmd, ":bankWebLink",          (object)c.BankWebLink ?? DBNull.Value);
            AddParam(cmd, ":bankLogo",             (object)c.BankLogo ?? DBNull.Value);
            AddParam(cmd, ":bgImage",              (object)c.BackgroundImage ?? DBNull.Value);
            AddParam(cmd, ":midBanner",            (object)c.MidBannerImage ?? DBNull.Value);
            AddParam(cmd, ":bottomBanner",         (object)c.BottomBannerImage ?? DBNull.Value);
            AddParam(cmd, ":emailFrom",            (object)c.EmailFromAddress ?? DBNull.Value);
            AddParam(cmd, ":emailFromName",        (object)c.EmailFromName ?? DBNull.Value);
            AddParam(cmd, ":whereCond",            (object)c.WhereCondition ?? DBNull.Value);
            AddParam(cmd, ":vipCond",              (object)c.VipCondition ?? DBNull.Value);
            AddParam(cmd, ":rewardCond",           (object)c.RewardCondition ?? DBNull.Value);
            AddParam(cmd, ":rewardContractCond",   (object)c.RewardContractCondition ?? DBNull.Value);
            AddParam(cmd, ":currFilter",           (object)c.CurrencyFilter ?? DBNull.Value);
            AddParam(cmd, ":installCond",          (object)c.InstallmentCondition ?? DBNull.Value);
            AddParam(cmd, ":paySystem",            (object)c.PaymentSystem ?? DBNull.Value);
            AddParam(cmd, ":procModes",            (long)c.ProcessingModes);
            AddParam(cmd, ":isReward",             c.IsRewardRun ? 1 : 0);
            AddParam(cmd, ":isSplit",              c.IsSplitOutput ? 1 : 0);
            AddParam(cmd, ":hasAttach",            c.HasAttachment ? 1 : 0);
            AddParam(cmd, ":saveDs",               c.SaveDataset ? 1 : 0);
            AddParam(cmd, ":showMsg",              c.ShowMessageBox ? 1 : 0);
            AddParam(cmd, ":nullCardDel",          c.RunNullCardDelete ? 1 : 0);
            AddParam(cmd, ":cardBranchMatch",      c.RunCardBranchMatch ? 1 : 0);
            AddParam(cmd, ":exclReward",           c.ExcludeReward ? 1 : 0);
            AddParam(cmd, ":waitPeriod",           c.WaitPeriodSeconds);
            AddParam(cmd, ":createdAt",            c.CreatedAt);
            AddParam(cmd, ":updatedAt",            c.UpdatedAt);
        }

        private static void AddParam(IDbCommand cmd, string name, object value)
        {
            var p = cmd.CreateParameter();
            p.ParameterName = name;
            p.Value         = value ?? DBNull.Value;
            cmd.Parameters.Add(p);
        }

        private static string SafeString(IDataReader r, string column)
        {
            int ord = r.GetOrdinal(column);
            return r.IsDBNull(ord) ? null : r.GetString(ord);
        }
    }
}

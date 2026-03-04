using System;
using System.Collections.Generic;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using StatementFile.Domain.Entities;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Infrastructure.Oracle
{
    public sealed class StatementRunRepository : IStatementRunRepository
    {
        private readonly IDbConnectionFactory _factory;

        public StatementRunRepository(IDbConnectionFactory factory) =>
            _factory = factory ?? throw new ArgumentNullException(nameof(factory));

        public int Add(StatementRun run)
        {
            const string sql = @"
                INSERT INTO STAT_STATEMENT_RUN
                    (CONFIG_ID, STATEMENT_DATE, STARTED_AT, IS_SUCCESS,
                     ERROR_MESSAGE, FILES_GENERATED, EMAILS_SENT,
                     STATEMENTS_COUNT, OUTPUT_DIRECTORY)
                VALUES
                    (:configId, :stmtDate, :startedAt, :isSuccess,
                     :errMsg, :filesGen, :emailsSent,
                     :stmtsCount, :outputDir)
                RETURNING ID INTO :newId";

            using var conn = _factory.CreateConnection();
            using var oCmd = new OracleCommand(sql, (OracleConnection)conn);

            oCmd.Parameters.Add(":configId",   run.ConfigId);
            oCmd.Parameters.Add(":stmtDate",   run.StatementDate);
            oCmd.Parameters.Add(":startedAt",  run.StartedAt);
            oCmd.Parameters.Add(":isSuccess",  run.IsSuccess ? 1 : 0);
            oCmd.Parameters.Add(":errMsg",     (object)run.ErrorMessage ?? DBNull.Value);
            oCmd.Parameters.Add(":filesGen",   run.FilesGenerated);
            oCmd.Parameters.Add(":emailsSent", run.EmailsSent);
            oCmd.Parameters.Add(":stmtsCount", run.StatementsCount);
            oCmd.Parameters.Add(":outputDir",  (object)run.OutputDirectory ?? DBNull.Value);

            var pId = oCmd.Parameters.Add(":newId", OracleDbType.Int32);
            pId.Direction = ParameterDirection.Output;

            oCmd.ExecuteNonQuery();
            return Convert.ToInt32(pId.Value.ToString());
        }

        public void Update(StatementRun run)
        {
            const string sql = @"
                UPDATE STAT_STATEMENT_RUN SET
                    FINISHED_AT      = :finishedAt,
                    IS_SUCCESS       = :isSuccess,
                    ERROR_MESSAGE    = :errMsg,
                    FILES_GENERATED  = :filesGen,
                    EMAILS_SENT      = :emailsSent,
                    STATEMENTS_COUNT = :stmtsCount,
                    OUTPUT_DIRECTORY = :outputDir
                WHERE ID = :id";

            using var conn = _factory.CreateConnection();
            using var cmd  = conn.CreateCommand();
            cmd.CommandText = sql;

            AddParam(cmd, ":finishedAt", (object)run.FinishedAt ?? DBNull.Value);
            AddParam(cmd, ":isSuccess",  run.IsSuccess ? 1 : 0);
            AddParam(cmd, ":errMsg",     (object)run.ErrorMessage ?? DBNull.Value);
            AddParam(cmd, ":filesGen",   run.FilesGenerated);
            AddParam(cmd, ":emailsSent", run.EmailsSent);
            AddParam(cmd, ":stmtsCount", run.StatementsCount);
            AddParam(cmd, ":outputDir",  (object)run.OutputDirectory ?? DBNull.Value);
            AddParam(cmd, ":id",         run.Id);

            cmd.ExecuteNonQuery();
        }

        public IReadOnlyList<StatementRun> GetByConfigId(int configId, int maxRows = 50)
        {
            const string sql = @"
                SELECT * FROM (
                    SELECT * FROM STAT_STATEMENT_RUN
                    WHERE CONFIG_ID = :configId
                    ORDER BY STARTED_AT DESC
                ) WHERE ROWNUM <= :maxRows";

            var list = new List<StatementRun>();
            using var conn   = _factory.CreateConnection();
            using var cmd    = conn.CreateCommand();
            cmd.CommandText  = sql;
            AddParam(cmd, ":configId", configId);
            AddParam(cmd, ":maxRows",  maxRows);

            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var run = StatementRun.Start(
                    configId:        reader.GetInt32(reader.GetOrdinal("CONFIG_ID")),
                    statementDate:   reader.GetDateTime(reader.GetOrdinal("STATEMENT_DATE")),
                    outputDirectory: SafeString(reader, "OUTPUT_DIRECTORY"));

                list.Add(run);
            }
            return list;
        }

        private static void AddParam(IDbCommand cmd, string name, object value)
        {
            var p = cmd.CreateParameter();
            p.ParameterName = name;
            p.Value         = value ?? DBNull.Value;
            cmd.Parameters.Add(p);
        }

        private static string SafeString(IDataReader r, string col)
        {
            int ord = r.GetOrdinal(col);
            return r.IsDBNull(ord) ? null : r.GetString(ord);
        }
    }
}

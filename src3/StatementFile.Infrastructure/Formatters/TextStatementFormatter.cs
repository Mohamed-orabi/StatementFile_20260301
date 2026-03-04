using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using Microsoft.Data.SqlClient;
using StatementFile.Domain.Entities;
using StatementFile.Domain.Interfaces;

namespace StatementFile.Infrastructure.Formatters
{
    /// <summary>
    /// Generates fixed-width plain-text statement files.
    /// Equivalent to the legacy <c>clsStatement_Export</c> class used in multiple
    /// switch-case branches in frmStatementFileExtn.runStatement().
    /// </summary>
    public sealed class TextStatementFormatter : IStatementFormatter
    {
        private readonly IEmailService _email;

        public TextStatementFormatter(IEmailService email) =>
            _email = email ?? throw new ArgumentNullException(nameof(email));

        public FormatterResult Format(BankProductConfig config, FormatterContext ctx)
        {
            string suffix      = ResolveTableSuffix(config);
            string masterTable = $"TSTATEMENTMASTER{suffix}";
            string detailTable = $"TSTATEMENTDETAIL{suffix}";

            int filesGenerated  = 0;
            int emailsSent      = 0;
            int statementsCount = 0;
            var errors          = new List<string>();

            try
            {
                using var conn = new SqlConnection(ctx.ConnectionString);
                conn.Open();

                var accounts = FetchAccounts(conn, masterTable, config, ctx.StatementDate);
                statementsCount = accounts.Rows.Count;

                foreach (DataRow account in accounts.Rows)
                {
                    try
                    {
                        var acctNo   = account["ACCOUNT_NUMBER"]?.ToString();
                        var txns     = FetchTransactions(conn, detailTable, acctNo, ctx.StatementDate);
                        var content  = RenderText(config, ctx, account, txns);
                        var fileName = $"{config.BankName}_{acctNo}_{ctx.StatementDate:yyyyMM}.txt";
                        var filePath = Path.Combine(ctx.OutputDirectory, fileName);

                        var fileMode = ctx.AppendData ? FileMode.Append : FileMode.Create;
                        using var fs = new FileStream(filePath, fileMode, FileAccess.Write);
                        using var sw = new StreamWriter(fs, Encoding.UTF8);
                        sw.Write(content);

                        filesGenerated++;

                        var email = ctx.EmailOverride ?? account["EMAIL"]?.ToString();
                        if (!string.IsNullOrWhiteSpace(email) && config.HasAttachment)
                        {
                            _email.Send(new EmailMessage
                            {
                                FromAddress = config.EmailFromAddress,
                                FromName    = config.EmailFromName,
                                ToAddresses = new[] { email },
                                Subject     = $"{config.BankFullName} Statement - {ctx.StatementDate:MMMM yyyy}",
                                Body        = "Please find your statement attached.",
                                IsBodyHtml  = false,
                                Attachments = new[] { filePath },
                            });
                            emailsSent++;
                        }
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Account {account["ACCOUNT_NUMBER"]}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                return new FormatterResult { Success = false, Error = ex.Message };
            }

            return new FormatterResult
            {
                Success         = errors.Count == 0,
                Error           = errors.Count > 0 ? string.Join("; ", errors) : null,
                FilesGenerated  = filesGenerated,
                EmailsSent      = emailsSent,
                StatementsCount = statementsCount,
                OutputDirectory = ctx.OutputDirectory,
            };
        }

        private static string ResolveTableSuffix(BankProductConfig c)
        {
            string s = c.StatementTypeSuffix?.ToUpperInvariant() ?? "CR";
            if (s.StartsWith("DB")) return "DB";
            if (s.StartsWith("CP")) return "CP";
            return "CR";
        }

        private static DataTable FetchAccounts(SqlConnection conn, string table,
            BankProductConfig cfg, DateTime stmtDate)
        {
            var sql = $@"SELECT M.* FROM {table} M
                         WHERE M.BRANCH_CODE = :bc
                           AND TRUNC(M.STATEMENT_DATE) = TRUNC(:sd)";

            if (!string.IsNullOrWhiteSpace(cfg.WhereCondition))
                sql += $" AND ({cfg.WhereCondition})";

            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add(":bc", cfg.BranchCode);
            cmd.Parameters.Add(":sd", stmtDate);
            using var a = new SqlDataAdapter(cmd);
            var dt = new DataTable();
            a.Fill(dt);
            return dt;
        }

        private static DataTable FetchTransactions(SqlConnection conn, string table,
            string acctNo, DateTime stmtDate)
        {
            var sql = $@"SELECT D.* FROM {table} D
                         WHERE D.ACCOUNT_NUMBER = :an
                           AND TRUNC(D.STATEMENT_DATE) = TRUNC(:sd)
                         ORDER BY D.TRANSACTION_DATE";

            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add(":an", acctNo);
            cmd.Parameters.Add(":sd", stmtDate);
            using var a = new SqlDataAdapter(cmd);
            var dt = new DataTable();
            a.Fill(dt);
            return dt;
        }

        private static string RenderText(BankProductConfig config, FormatterContext ctx,
            DataRow account, DataTable transactions)
        {
            var sb = new StringBuilder();
            sb.AppendLine(new string('=', 80));
            sb.AppendLine($"  {config.BankFullName.ToUpperInvariant()} - ACCOUNT STATEMENT");
            sb.AppendLine(new string('=', 80));
            sb.AppendLine($"  Account No : {account["ACCOUNT_NUMBER"]}");
            sb.AppendLine($"  Customer   : {account["CUSTOMER_NAME"]}");
            sb.AppendLine($"  Stmt Date  : {ctx.StatementDate:dd-MMM-yyyy}");
            sb.AppendLine(new string('-', 80));
            sb.AppendLine($"  {"Date",-12} {"Description",-35} {"Debit",10} {"Credit",10} {"Balance",10}");
            sb.AppendLine(new string('-', 80));

            foreach (DataRow txn in transactions.Rows)
            {
                sb.AppendLine(string.Format("  {0,-12} {1,-35} {2,10} {3,10} {4,10}",
                    txn["TRANSACTION_DATE"],
                    txn["DESCRIPTION"],
                    txn["DEBIT_AMOUNT"],
                    txn["CREDIT_AMOUNT"],
                    txn["BALANCE"]));
            }

            sb.AppendLine(new string('=', 80));
            return sb.ToString();
        }
    }
}

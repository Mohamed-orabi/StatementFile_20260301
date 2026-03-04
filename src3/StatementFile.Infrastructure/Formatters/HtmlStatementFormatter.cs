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
    /// Generates HTML statement files, one file per account.
    ///
    /// Equivalent to the legacy <c>clsBasStatementHtml</c> and
    /// <c>clsStatHtml*</c> family of classes used across many switch-case
    /// branches in frmStatementFileExtn.runStatement().
    ///
    /// Behaviour:
    ///   1. Queries TSTATEMENTMASTER{suffix} for accounts matching the config filters.
    ///   2. For each account, queries TSTATEMENTDETAIL{suffix} for transactions.
    ///   3. Renders an HTML template with bank branding, account summary, and transaction list.
    ///   4. Writes the file to OutputDirectory with the naming convention
    ///      {BankName}_{AccountNumber}_{StatementDate:yyyyMM}.html.
    ///   5. Optionally emails the statement to the account holder.
    /// </summary>
    public sealed class HtmlStatementFormatter : IStatementFormatter
    {
        private readonly IEmailService _email;

        public HtmlStatementFormatter(IEmailService email) =>
            _email = email ?? throw new ArgumentNullException(nameof(email));

        public FormatterResult Format(BankProductConfig config, FormatterContext ctx)
        {
            string suffix = ResolveTableSuffix(config);
            string masterTable = $"TSTATEMENTMASTER{suffix}";
            string detailTable = $"TSTATEMENTDETAIL{suffix}";

            int filesGenerated   = 0;
            int emailsSent       = 0;
            int statementsCount  = 0;
            var errors           = new List<string>();

            try
            {
                using var conn = new SqlConnection(ctx.ConnectionString);
                conn.Open();

                var accounts = FetchAccounts(conn, masterTable, config, ctx.StatementDate);
                statementsCount = accounts.Count;

                foreach (DataRow account in accounts.Rows)
                {
                    try
                    {
                        var acctNo   = account["ACCOUNT_NUMBER"]?.ToString();
                        var txns     = FetchTransactions(conn, detailTable, acctNo, ctx.StatementDate, config);
                        var html     = RenderHtml(config, ctx, account, txns);
                        var fileName = BuildFileName(config, acctNo, ctx.StatementDate);
                        var filePath = Path.Combine(ctx.OutputDirectory, fileName);

                        File.WriteAllText(filePath, html, Encoding.UTF8);
                        filesGenerated++;

                        var email = ctx.EmailOverride ?? account["EMAIL"]?.ToString();
                        if (!string.IsNullOrWhiteSpace(email) && ShouldEmail(config))
                        {
                            SendStatement(config, email, filePath, ctx.StatementDate);
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

            string errorSummary = errors.Count > 0 ? string.Join("; ", errors) : null;
            return new FormatterResult
            {
                Success          = errors.Count == 0,
                Error            = errorSummary,
                FilesGenerated   = filesGenerated,
                EmailsSent       = emailsSent,
                StatementsCount  = statementsCount,
                OutputDirectory  = ctx.OutputDirectory,
            };
        }

        // ── Helpers ──────────────────────────────────────────────────────────────

        private static string ResolveTableSuffix(BankProductConfig config)
        {
            // The StatementTypeSuffix may be "CR", "DB", "CP", "CRCORP", etc.
            // Strip any extra suffix beyond the card-type prefix.
            string s = config.StatementTypeSuffix?.ToUpperInvariant() ?? "CR";
            if (s.StartsWith("DB"))  return "DB";
            if (s.StartsWith("CP"))  return "CP";
            return "CR";
        }

        private static DataTable FetchAccounts(SqlConnection conn, string table,
            BankProductConfig cfg, DateTime stmtDate)
        {
            var sb = new StringBuilder($@"
                SELECT M.*
                FROM {table} M
                WHERE M.BRANCH_CODE = :branchCode
                  AND TRUNC(M.STATEMENT_DATE) = TRUNC(:stmtDate)");

            if (!string.IsNullOrWhiteSpace(cfg.WhereCondition))
                sb.Append($" AND ({cfg.WhereCondition})");
            if (!string.IsNullOrWhiteSpace(cfg.CurrencyFilter))
                sb.Append($" AND ({cfg.CurrencyFilter})");
            if (!string.IsNullOrWhiteSpace(cfg.VipCondition))
                sb.Append($" AND ({cfg.VipCondition})");
            if (cfg.ExcludeReward && !string.IsNullOrWhiteSpace(cfg.RewardCondition))
                sb.Append($" AND NOT ({cfg.RewardCondition})");

            using var cmd = new SqlCommand(sb.ToString(), conn);
            cmd.Parameters.Add(":branchCode", cfg.BranchCode);
            cmd.Parameters.Add(":stmtDate",   stmtDate);

            using var adapter = new SqlDataAdapter(cmd);
            var dt = new DataTable();
            adapter.Fill(dt);
            return dt;
        }

        private static DataTable FetchTransactions(SqlConnection conn, string table,
            string accountNumber, DateTime stmtDate, BankProductConfig cfg)
        {
            var sb = new StringBuilder($@"
                SELECT D.*
                FROM {table} D
                WHERE D.ACCOUNT_NUMBER = :acctNo
                  AND TRUNC(D.STATEMENT_DATE) = TRUNC(:stmtDate)");

            if (!string.IsNullOrWhiteSpace(cfg.InstallmentCondition))
                sb.Append($" AND NOT ({cfg.InstallmentCondition})");

            using var cmd = new SqlCommand(sb.ToString(), conn);
            cmd.Parameters.Add(":acctNo",   accountNumber);
            cmd.Parameters.Add(":stmtDate", stmtDate);

            using var adapter = new SqlDataAdapter(cmd);
            var dt = new DataTable();
            adapter.Fill(dt);
            return dt;
        }

        private static string RenderHtml(BankProductConfig config, FormatterContext ctx,
            DataRow account, DataTable transactions)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html><html><head>");
            sb.AppendLine($"<meta charset=\"UTF-8\"><title>{config.BankFullName} Statement</title>");
            sb.AppendLine("<style>body{font-family:Arial,sans-serif;font-size:12px;} table{border-collapse:collapse;width:100%;} td,th{border:1px solid #ccc;padding:4px;}</style>");
            sb.AppendLine("</head><body>");

            // Header banner
            if (!string.IsNullOrWhiteSpace(config.BankLogo))
                sb.AppendLine($"<img src=\"{config.BankLogo}\" alt=\"{config.BankFullName}\" /><br/>");

            sb.AppendLine($"<h2>{config.BankFullName}</h2>");
            sb.AppendLine($"<p>Statement Date: {ctx.StatementDate:dd-MMM-yyyy}</p>");
            sb.AppendLine($"<p>Account: {account["ACCOUNT_NUMBER"]}</p>");
            sb.AppendLine($"<p>Customer: {account["CUSTOMER_NAME"]}</p>");

            // Transactions table
            sb.AppendLine("<h3>Transactions</h3>");
            sb.AppendLine("<table><tr><th>Date</th><th>Description</th><th>Debit</th><th>Credit</th><th>Balance</th></tr>");
            foreach (DataRow txn in transactions.Rows)
            {
                sb.AppendLine("<tr>");
                sb.AppendLine($"<td>{txn["TRANSACTION_DATE"]}</td>");
                sb.AppendLine($"<td>{txn["DESCRIPTION"]}</td>");
                sb.AppendLine($"<td>{txn["DEBIT_AMOUNT"]}</td>");
                sb.AppendLine($"<td>{txn["CREDIT_AMOUNT"]}</td>");
                sb.AppendLine($"<td>{txn["BALANCE"]}</td>");
                sb.AppendLine("</tr>");
            }
            sb.AppendLine("</table>");

            if (!string.IsNullOrWhiteSpace(config.BottomBannerImage))
                sb.AppendLine($"<img src=\"{config.BottomBannerImage}\" />");

            if (!string.IsNullOrWhiteSpace(config.BankWebLink))
                sb.AppendLine($"<p><a href=\"{config.BankWebLink}\">{config.BankWebLink}</a></p>");

            sb.AppendLine("</body></html>");
            return sb.ToString();
        }

        private static string BuildFileName(BankProductConfig config, string accountNumber, DateTime stmtDate) =>
            $"{config.BankName}_{accountNumber}_{stmtDate:yyyyMM}.html";

        private static bool ShouldEmail(BankProductConfig config) =>
            config.OutputType == StatementFile.Domain.Enums.StatementOutputType.Email
            || config.HasAttachment;

        private void SendStatement(BankProductConfig config, string toEmail, string filePath, DateTime stmtDate)
        {
            _email.Send(new EmailMessage
            {
                FromAddress  = config.EmailFromAddress,
                FromName     = config.EmailFromName,
                ToAddresses  = new[] { toEmail },
                Subject      = $"{config.BankFullName} Statement - {stmtDate:MMMM yyyy}",
                Body         = $"<p>Dear Customer,</p><p>Please find your {config.BankFullName} statement attached.</p>",
                IsBodyHtml   = true,
                Attachments  = new[] { filePath },
            });
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using StatementFile.Domain.Entities;

namespace StatementFile.Application.UseCases.MerchantOnboarding
{
    /// <summary>
    /// Converts the raw XML DataSet (parsed from the merchant XML file) into
    /// a collection of <see cref="MerchantStatement"/> aggregate roots.
    /// Preserves all field-mapping logic from the legacy clsStatementMrchXml.
    /// </summary>
    internal static class MerchantStatementXmlMapper
    {
        private const string ColEndDate       = "EndDate";
        private const string ColStartDate     = "StartDate";
        private const string ColStatementNo   = "StatementNo";
        private const string ColAccount       = "Account";
        private const string ColExternalAcct  = "ExternalAccount";
        private const string ColOD            = "OD";
        private const string ColTD            = "TD";
        private const string ColO             = "O";
        private const string ColA             = "A";
        private const string ColOA            = "OA";
        private const string ColCF            = "CF";
        private const string ColS             = "S";
        private const string ColD             = "D";

        public static IEnumerable<MerchantStatement> MapFromDataSet(
            DataSet ds,
            string  bankCode,
            string  bankName,
            string  bankFullName,
            DateTime statDate)
        {
            if (!ds.Tables.Contains("Statement"))
                yield break;

            int statMasterCode = 0;
            int detailCode     = 0;
            int branchInt      = int.TryParse(bankCode, out int b) ? b : 0;

            foreach (DataRow mRow in ds.Tables["Statement"].Rows)
            {
                statMasterCode++;

                DateTime start = ParseDate(mRow, ColStartDate);
                DateTime end   = ParseDate(mRow, ColEndDate);

                string statNo   = mRow.Table.Columns.Contains(ColStatementNo)
                                  ? mRow[ColStatementNo].ToString()
                                  : statMasterCode.ToString();
                string account  = mRow.Table.Columns.Contains(ColAccount)
                                  ? mRow[ColAccount].ToString()
                                  : string.Empty;
                string extAcct  = mRow.Table.Columns.Contains(ColExternalAcct)
                                  ? mRow[ColExternalAcct].ToString()
                                  : string.Empty;

                var statement = MerchantStatement.Create(
                    statMasterCode, branchInt,
                    bankName, bankFullName,
                    statDate, statNo,
                    account, extAcct,
                    start, end);

                // Map child operation rows
                DataRow[] childRows = ds.Relations.Contains("StaementNoDR")
                    ? mRow.GetChildRows("StaementNoDR")
                    : new DataRow[0];

                foreach (DataRow dRow in childRows)
                {
                    // Skip reimbursement-payment rows (legacy filter)
                    if (dRow.Table.Columns.Contains(ColD) &&
                        dRow[ColD].ToString() == "Reimbursemnt - Payment")
                        continue;

                    detailCode++;

                    DateTime opDate   = ParseDate(dRow, ColOD);
                    DateTime transDate = dRow.Table.Columns.Contains(ColTD) &&
                                        !string.IsNullOrWhiteSpace(dRow[ColTD].ToString())
                                        ? ParseDate(dRow, ColTD)
                                        : opDate; // TD defaults to OD when null

                    var op = MerchantOperation.Create(
                        detailCode, statMasterCode, branchInt,
                        description:     GetString(dRow, ColD),
                        originalAmount:  GetDecimal(dRow, ColO),
                        amount:          GetDecimal(dRow, ColA),
                        otherAmount:     GetDecimal(dRow, ColOA),
                        commissionFee:   GetDecimal(dRow, ColCF),
                        settlement:      GetDecimal(dRow, ColS),
                        operationDate:   opDate,
                        transactionDate: transDate);

                    statement.AddOperation(op);
                }

                yield return statement;
            }
        }

        // ── Helpers ──────────────────────────────────────────────────────────────

        private static DateTime ParseDate(DataRow row, string column)
        {
            if (!row.Table.Columns.Contains(column)) return DateTime.MinValue;
            string raw = row[column].ToString().Trim();
            if (string.IsNullOrEmpty(raw)) return DateTime.MinValue;
            return DateTime.TryParse(raw, out DateTime d) ? d : DateTime.MinValue;
        }

        private static decimal GetDecimal(DataRow row, string column)
        {
            if (!row.Table.Columns.Contains(column)) return 0m;
            string raw = row[column].ToString().Trim();
            if (string.IsNullOrEmpty(raw)) return 0m;
            return decimal.TryParse(raw, out decimal d) ? d : 0m;
        }

        private static string GetString(DataRow row, string column)
        {
            if (!row.Table.Columns.Contains(column)) return string.Empty;
            return row[column].ToString();
        }
    }
}

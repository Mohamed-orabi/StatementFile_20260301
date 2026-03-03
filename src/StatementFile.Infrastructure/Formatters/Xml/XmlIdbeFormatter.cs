using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using StatementFile.Application.UseCases.StatementGeneration;
using StatementFile.Domain.Interfaces.Services;

namespace StatementFile.Infrastructure.Formatters.Xml
{
    /// <summary>
    /// IDBE VIP XML statement formatter.
    /// Writes the master DataSet as XML with schema
    /// (DataSet.WriteXml WriteSchema mode, matching the legacy clsStatXML_IDBE).
    /// VIP filtering is already applied upstream by LoadVipOnly() before this
    /// formatter is called.
    /// Produces a .xml file + .MD5 checksum file.
    /// </summary>
    public sealed class XmlIdbeFormatter : IStatementFormatterService
    {
        public string FormatterKey => "XML_IDBE";

        public IEnumerable<string> Format(
            StatementDataContext     ctx,
            string                   outputDirectory,
            GenerateStatementCommand cmd)
        {
            Directory.CreateDirectory(outputDirectory);

            string xmlFile = Path.Combine(outputDirectory,
                $"{cmd.BankName}_{cmd.StatementDate:yyyyMM}.xml");
            string md5File = Path.Combine(outputDirectory,
                $"{cmd.BankName}_{cmd.StatementDate:yyyyMM}.MD5");

            // Write DataSet XML with embedded schema (matching DataSet.WriteXml WriteSchema)
            ctx.MasterDataSet?.WriteXml(xmlFile, System.Data.XmlWriteMode.WriteSchema);

            // Write MD5 checksum
            if (File.Exists(xmlFile))
            {
                using var md5 = MD5.Create();
                using var fs  = File.OpenRead(xmlFile);
                byte[] hash   = md5.ComputeHash(fs);
                string hex    = BitConverter.ToString(hash)
                                .Replace("-", string.Empty)
                                .ToUpperInvariant();
                File.WriteAllText(md5File,
                    $"{Path.GetFileName(xmlFile)}  >>  {hex}");
            }

            return new[] { xmlFile, md5File };
        }
    }
}

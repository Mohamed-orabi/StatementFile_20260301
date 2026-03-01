using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Collections.Specialized;

namespace StatementFile.StatementFile.mailConfiguration
{
    internal class ConfigurationReader
    {
        private static readonly List<EmailConfiguration> _EmailConfig;
        private static readonly PathConfiguration _PathConfig;

        private static readonly string filepath = $@"{System.IO.Directory.GetCurrentDirectory()}{ConfigurationManager.AppSettings["mailConfigFilePath"]}";
        
        private static readonly string pathConfigurationFilepath = $@"{System.IO.Directory.GetCurrentDirectory()}{ConfigurationManager.AppSettings["pathConfiguration"]}";

        static ConfigurationReader()
        {
            _EmailConfig = LoadConfigurations();
            _PathConfig = LoadPathConfiguration();
        }
        /// <summary>
        /// load the JSON mails file 
        /// get the mail configuration in the _EMailConfig
        /// </summary>
        public static List<EmailConfiguration> LoadConfigurations()
        {
            //clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt","pathFileName"+ ConfigurationManager.AppSettings["mailConfigFilePath"].ToString());

            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine +"pathDirectory: " + filepath + Environment.NewLine, System.IO.FileMode.Append);


            try
            {
                if (System.IO.File.Exists(filepath))
                {

                    string text = System.IO.File.ReadAllText(filepath);
                    if (!string.IsNullOrWhiteSpace(text))
                        return JsonConvert.DeserializeObject<List<EmailConfiguration>>(text);

                }
                else
                {
                    clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", $@"{Environment.NewLine}expected filepath: exe\statement\mailConfiguration.json But GetCurrentDirectory:{filepath}{Environment.NewLine}", System.IO.FileMode.Append);
                    throw new Exception( );
                }

            }
            catch (Exception ex)
            {
                clsBasErrors.catchError(ex);
                System.Environment.Exit(0);
            }

            return default;
        }

        /// <summary>
        /// if Bank has special validate( QNB_ALAHLI or ICBG or HBLN ) add the special validation to them
        /// if it _ValidateEmail == true return emails with tag null and true
        /// if it _ValidateEmail == false return emails with tag null and false
        /// if _EmailConfig return null and ClsError will raise 
        /// </summary>
        /// <param name="bankName"></param>
        /// <param name="bankType"></param>
        /// <returns></returns>
        public static EmailConfiguration LoadBankMailConfiguration(string bankName, bool bankType)
        {
            if (_EmailConfig != null)
            {
                EmailConfiguration emailConfig = _EmailConfig.FirstOrDefault(e => e.name == bankName);
                if (bankName.Equals("QNB_ALAHLI") || bankName.Equals("ICBG") || bankName.Equals("HBLN"))
                {
                    emailConfig.to = emailConfig.to.Where(v => bankType ? v.valid != false : v.valid != true).ToList();
                }
            
                return emailConfig;
            }
            return null;
        }

        
        
        public static PathConfiguration LoadPathConfiguration() {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + "pathDirectory: " + pathConfigurationFilepath + Environment.NewLine, System.IO.FileMode.Append);
            try
            {
                if (System.IO.File.Exists(pathConfigurationFilepath))
                {

                    string text = System.IO.File.ReadAllText(pathConfigurationFilepath);
                    if (!string.IsNullOrWhiteSpace(text))
                        return JsonConvert.DeserializeObject<PathConfiguration>(text);

                }
                else
                {
                    clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", $@"{Environment.NewLine}expected filepath: exe\statement\pathConfiguration.json But GetCurrentDirectory:{pathConfigurationFilepath}{Environment.NewLine}", System.IO.FileMode.Append);
                    throw new Exception();
                }

            }
            catch (Exception ex)
            {
                clsBasErrors.catchError(ex);
                System.Environment.Exit(0);
            }

            return default;
        }


    }
}

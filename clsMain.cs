using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace StatementFile
{
    /// <summary>
    /// Summary description for clsMain.
    /// </summary>
    public class clsMain
  {
    public clsMain()
    {

    }

    [STAThread]
    static public void Main()
    {
          clsBasUserData.applicationName = "StatementFile";
          clsBasUserData.databaseType = "Oracle";
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);

            #region Zip file
            //SharpZip zip = new SharpZip();


            //ArrayList aryLstFilesStmt = new ArrayList();

            //string nam1 = @"D:\TEMP\P20Files\Statement\_WaitReply\202204BDCA_Credit_MCStandard_ClientsEmails"
            //            + "\\" + "BDCA" + "_" + "Credit_MCStandard_ClientsEmails" + "_STMT" + "202204" + ".zip";

            //string pdfPath = @"D:\Temp\P20Files\Statement\_WaitReply\202204BDCA_Credit_MCStandard_ClientsEmails\STMT";

            //string fname = "D:\\TEMP\\P20Files\\Statement\\_WaitReply\\202204BDCA_Credit_MCStandard_ClientsEmails\\STMT\\";

            //List<string> Pdf = new DirectoryInfo(pdfPath).GetFiles("*.pdf", SearchOption.TopDirectoryOnly).Select(filename => fname + filename).ToList();

            //foreach (var item in Pdf)
            //{
            //    aryLstFilesStmt.Add(item);
            //}

            //zip.createZip(aryLstFilesStmt, nam1, "");


            #endregion

            #region test mailConfiguration JSON
            //try
            //{
            //    EmailConfiguration mailConfig4 = ConfigurationReader.LoadBankMailConfiguration("QNB_ALAHLI", false);
            //    string msg = "", to = "", cc = "", path = "";
            //    if (mailConfig4 != null)
            //    {
            //        msg = mailConfig4.message;
            //        path = mailConfig4.path;

            //        mailConfig4.to.ForEach(item =>
            //        {
            //            to += (item.email + "\n");
            //        });
            //        mailConfig4.cc.ForEach(item =>
            //        {
            //            cc += (item.email + item.name + "\n");
            //        });
            //    }
            //}
            //catch (Exception ex)
            //{
            //    clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);
            //    clsBasErrors.catchError(ex);
            //    System.Environment.Exit(0);
            //}
            #endregion

            
            try
            {
                EventLog.WriteEntry("STMT", "Applican start", EventLogEntryType.Error);
                Application.Run(new frmBasLogin());
            }
            catch (Exception ex)
            {
                clsBasErrors.catchError(ex);
                System.Environment.Exit(0);
            }
            //clsDbOracleLayer.clearAllPools();
            clsBasUserData.b4CloseProgram();
    }
  }
}
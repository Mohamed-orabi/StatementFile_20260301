using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Data;
using Oracle.DataAccess.Client;
using System.Collections;
//using System.Data.OleDb;
//
public class clsBasUpdateStat
{


  public void UpdateStat(string pStrFileNameIn)
  {
    FileStream fileStrmIn = null;
    StreamReader strmRead = null;
    string strReadLine;
    string[] intPos;
    string strSqlActn = string.Empty;
    int SqlCnt;

    Cursor.Current = Cursors.WaitCursor;
    try
    {
      OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
      conn.Open();

      fileStrmIn = new FileStream(pStrFileNameIn, FileMode.Open);
      strmRead = new StreamReader(fileStrmIn, Encoding.ASCII);

      strSqlActn = "begin "; SqlCnt = 0;
      while ((strReadLine = strmRead.ReadLine()) != null)
      {
        SqlCnt++;
        intPos = strReadLine.Split('|');
        strSqlActn = strSqlActn + "update " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t set t.customerzipcode = '" + intPos[1] + "', t.barcode = '" + intPos[2] + "' where t.branch = 4 and t.clientid = " + intPos[0] + ";";
        if (SqlCnt > 1000)
        {
          strSqlActn = strSqlActn + " end;";
          (new OracleCommand(strSqlActn, conn)).ExecuteNonQuery();
          strSqlActn = "begin "; SqlCnt = 0;
        }
      }

      if (SqlCnt > 0)
      {
        strSqlActn = strSqlActn + " end;";
        (new OracleCommand(strSqlActn, conn)).ExecuteNonQuery();
      }
    }
    catch (NotSupportedException ex)  //(Exception ex)  //
    {
      clsBasErrors.catchError(ex);
    }
    finally
    {
      fileStrmIn.Close();
      strmRead.Close();
      Cursor.Current = Cursors.Default;
    }
  }


  public clsBasUpdateStat()
  {
  }

  ~clsBasUpdateStat()
  {
  }
}

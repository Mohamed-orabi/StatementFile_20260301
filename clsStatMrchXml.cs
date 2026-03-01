using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections.Generic;

class clsStatMrchXml
{
  private FileStream fileStrm, fileSummary;
  private StreamWriter streamWrit, streamSummary;
  private DataSet DSstatement;


  public void readXmlFile(string pStrFileName)
  {
    DSstatement = new DataSet();
    DSstatement.ReadXml(pStrFileName, XmlReadMode.ReadSchema);//XmlReadMode.InferSchema
    frmShowFileGrid showFile = new frmShowFileGrid();
    showFile.showData(DSstatement);
    showFile.Show();
  }


  public void printCardNum(DataSet pDsStatement, string pStrFileName)
  {
    try
    {
      DSstatement = pDsStatement;

      // open output file
      fileStrm = new FileStream(pStrFileName, FileMode.Create); //Create
      streamWrit = new StreamWriter(fileStrm, Encoding.Default);

      foreach (DataRow mRow in DSstatement.Tables["tStatementMasterTable"].Rows)
      {
        streamWrit.WriteLine(mRow["cardno"]);
      } //end of Master foreach
    }
    catch (OracleException ex)
    {
      clsDbOracleLayer.catchError(ex);
    }
    catch (NotSupportedException ex)  //(Exception ex)  //
    {
      clsBasErrors.catchError(ex);
    }
    finally
    {
      // Close output File
      streamWrit.Flush();
      streamWrit.Close();
      fileStrm.Close();

    }

  }


  public clsStatMrchXml()
  {
  }

  ~clsStatMrchXml()
  {
  }
}

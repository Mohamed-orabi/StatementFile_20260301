using System;
using System.IO;
using System.Text;
using System.Windows.Forms;


//
public class clsValidatePageSize
{
  private FileStream fileStrmIn, fileStrmOut;
  private StreamReader strmRead;
  private StreamWriter strmWrite;
  private string strEndOfPage = "\u000C"; //+ basText.replicat(" ",85) + "F 2"  ;  //+ "\u000D"+ "M" ^L + ^M
  protected string strEndOfAccount = "----------------- END OF STATEMENT -----------------";  //  + basText.replicat(" ",85) + "F 2"  + "\u000D"+ "M" ^L + ^M
  private string strMessage = string.Empty;

  public clsValidatePageSize()
  {
  }

  public int ValidatePageSize(string pStrFileNameIn, int pPageSize, string pCharacter)
  {
    string strReadLine, strWriteLine = string.Empty, strFileName = string.Empty;
    int numOfLine = 0, numOfErr = 0, curPageRow = 0;
    strEndOfPage = pCharacter;

    try
    {
        // open output file
        fileStrmIn = new FileStream(pStrFileNameIn, FileMode.Open);
        strmRead = new StreamReader(fileStrmIn, Encoding.ASCII);

        strFileName = clsBasFile.getPathWithoutExtn(pStrFileNameIn) + "_Err2." + clsBasFile.getFileExtn(pStrFileNameIn);
        fileStrmOut = new FileStream(strFileName, FileMode.Create);
        strmWrite = new StreamWriter(fileStrmOut, Encoding.Default);

        strReadLine = strmRead.ReadLine();
        while ((strReadLine = strmRead.ReadLine()) != null)
        {
            numOfLine++;
            curPageRow++;
            if (strReadLine == strEndOfPage || (strReadLine == strEndOfPage + strEndOfAccount))
            {
                if (curPageRow != pPageSize)
                {
                    strmWrite.WriteLine(numOfLine.ToString());
                    numOfErr++;
                }
                curPageRow = 0;
            }
            /*if(curPageRow == pPageSize)
            {
              curPageRow = 1;
            }*/
        }
    }
    catch (Exception x)
    {
        clsBasErrors.catchError(x);
    }

    //catch (NotSupportedException ex)  //(Exception ex)  //
    //{
    //    clsBasErrors.catchError(ex);
    //}
    finally
    {
        // Close Input & output File
        fileStrmIn.Close();
        strmRead.Close();

        strmWrite.Flush();
        strmWrite.Close();
        fileStrmOut.Close();

        strMessage = "File Contain Errors loged to file" + strFileName;
        if (numOfErr == 0)
        {
            strMessage = "File Contain No Errors";
            clsBasFile.deleteFile(strFileName);
        }
    }
    return numOfErr;
  }


  protected void printHeader()
  {
    strmWrite.WriteLine(String.Empty);
  }

  public string outMessage
  {
    get { return strMessage; }
    //    set {	strMessage = value;}
  }//outMessage

  ~clsValidatePageSize()
  {
  }

}

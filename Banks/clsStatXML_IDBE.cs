using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;

// Branch 16
public class clsStatXML_IDBE : clsBasStatement
    {
    protected string strBankName;

    protected string strOutputPath;
    protected DateTime vCurDate;
    protected frmStatementFile frmMain;
    protected string strFileNam, stmntType;
    protected ArrayList aryLstFiles = new ArrayList();
    protected string XmlFileName = string.Empty;
    protected int pBranch = 0;

    public clsStatXML_IDBE()
        {

        }

    public virtual string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
        {
        string rtrnStr = "Successfully Generate " + pBankName, ppStrFileName = pStrFileName;
        int curMonth = pCurDate.Month;
        strFileNam = pStrFileName;
        stmntType = pStmntType;
        pBranch = pBankCode;

        try
            {
            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
            pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".xml";
            XmlFileName = pStrFileName;
            curBranchVal = pBankCode; // 6; //6  = real   1 = test

            FillStatementDataSet(pBankCode, "vip"); //DSstatement =  //16); //3
            // Write Data to XML File 

            DSstatement.WriteXml(pStrFileName, XmlWriteMode.WriteSchema);

            aryLstFiles.Add(pStrFileName);
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
            finalizeStat();
            }
        return rtrnStr;
        }

    public void finalizeStat()
        {
        clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(XmlFileName) + ".MD5");
        aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(XmlFileName) + ".MD5");
        SharpZip zip = new SharpZip();
        zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(XmlFileName) + ".zip", "");
        DSstatement.Dispose();

        }

    public string mainSortOrder
        {
        set { strOrder = value; }
        }// mainSortOrder

    public string bankName
        {
        get { return strBankName; }
        set { strBankName = value; }
        }// bankName

    public frmStatementFile setFrm
        {
        set { frmMain = value; }
        }// setFrm

    ~clsStatXML_IDBE()
        {
        DSstatement.Dispose();
        }
    }

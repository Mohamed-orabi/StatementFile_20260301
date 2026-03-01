using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;


public class clsStatHtmlMrch 
{
  protected string strBankName;
  protected FileStream fileSummary, fileEmails, fileNoEmails;
  protected StreamWriter streamSummary, streamEmails, streamNoEmails;
  protected FileStream fileStrmErr;
  protected StreamWriter strmWriteErr;
  protected DataRow masterRow;
  protected DataRow detailRow;
  protected string curAccountNo, prevAccountNo = String.Empty;//,PrevCardNo
  protected decimal totNetUsage = 0;
  protected DataRow[] cardsRows, accountRows, rewardRows;
  protected DataRow[] mainRows;

  protected string strOutputPath, strOutputFile, fileSummaryName;
  protected DateTime vCurDate;
  protected ArrayList aryLstFiles;

  protected StringBuilder emailStr = new StringBuilder("");
  protected string emailLabel, strBankLogo;
  protected string strEmailFrom = "statement@emp-group.com", strEmailFromTmp = string.Empty;  //"cardservices@zenithbank.com"
  protected string strbankWebLink = "www.emp-group.com";  //www.zenithbank.com
  protected string strbankWebLinkService = "www.emp-group.com";  //www.zenithbank.com
  protected string emailTo = string.Empty, curAccountNumber, curCardNumber, curClientID;//, curEmail
  protected string strEmailFileName, strNoEmailFileName;
  //protected frmStatementFile frmMain;
  protected int totRec = 1, numMail, numOfTry = 0;
  protected ArrayList pLstCC = new ArrayList(), pLstBCC = new ArrayList();
  protected string logoAlignmentStr = "center";
  protected string strbackGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
  protected string strVisaLogo = string.Empty;
  protected string statMessageFileVal = string.Empty;
  protected string statMessage = "&nbsp;";//Null
  protected string statMessageFileMonthlyVal = string.Empty;
  protected string statMessageMonthly = "&nbsp;";//Null
  protected string isSuplStr = string.Empty;//Null
  protected DataRow[] emailRows = null;
  protected DataRow masterRelatedRow;
  //protected clsEmail sndMail;
  protected string emailFromNameStr;
  protected string strMobileNum = string.Empty;
  protected int waitPeriodVal = 7000;
  protected int noOfEmails, noOfBadEmails, noOfWithoutEmails;
  protected string emailLabelTmp;  //"cardservices@zenithbank.com"
  protected DataRow emailRow;
  protected ArrayList pLstTo = new ArrayList(), pLstAttachedFile = new ArrayList();

  private string strStatementPath;
  protected int totNoOfTransactions ;
  protected string fontTypeSize = "<font size='2'>"; //"<font color='#808080' size='2'>";

  private string bankFullName, bankName;
  private decimal totTrans = 0, totComm = 0, totAmount = 0;

  public clsStatHtmlMrch()
  {
  }

  public string Statement(string pStrFileName, DataSet DSstatement, string pBankName, string pBankFullName, DateTime pDate)
  {
    string rtrnStr = "Successfully Generate " + pBankName;
    int curMonth = pDate.Month;
    vCurDate = pDate;

    bool preExit = true;
    bankName = pBankName;
    bankFullName = pBankFullName;
    curMonth = pDate.Month;
    aryLstFiles = new ArrayList();
    try
    {
      strStatementPath = clsBasFile.getPathWithoutFile(pStrFileName) + pBankName + "_MerchantStatement_" + pDate.ToString("yyyyMMdd_HHmmss");
      //clsBasFile.createDirectory(strStatementPath);
      //strDestDataFile = strStatementPath + "\\" + pBankName + "_MerchantStatement_" + pDate.ToString("yyyyMMdd_HHmmss") + ".mdb";
      strBankName = pBankName;
      emailLabel = pBankName + " Statement for " + vCurDate.ToString("dd/MM/yyyy"); //"BAI statement for 02/2008"
      strOutputFile = pStrFileName;
      strEmailFileName = clsBasFile.getPathWithoutExtn( pStrFileName) + "_MerchantEmails.txt" ;//+ ".txt"
      strNoEmailFileName = clsBasFile.getPathWithoutExtn(pStrFileName) + "_MerchantNoEmails.txt";//+ ".txt"

      // open emails file
      fileEmails = new FileStream(strEmailFileName, FileMode.Create); //+ ".txt"Create
      streamEmails = new StreamWriter(fileEmails, Encoding.Default);
      //streamEmails.WriteLine("AccountNumber" + "|" + "CardNumber" + "|" + mClientid + "|" + "Email");
      streamEmails.WriteLine("AccountNumber" + "|" + "ClientID" + "|" + "Email" + "|"+ "MobilePhone" + "|"+"Date Time");
      streamEmails.AutoFlush = true;

      // open No emails file
      fileNoEmails = new FileStream(strNoEmailFileName, FileMode.Create); //Create
      streamNoEmails = new StreamWriter(fileNoEmails, Encoding.Default);
      streamNoEmails.WriteLine("AccountNumber" + "|" + "ClientID" + "|" + "Email" + "|" + "Mobile Phone");//"AccountNumber" + "|" + "ClientID"
      streamNoEmails.AutoFlush = true;
      
      if (!string.IsNullOrEmpty(statMessageFileVal))
      {
        FileStream filRead = null;
        StreamReader filStream = null;
        filRead = new FileStream(statMessageFileVal, FileMode.Open);
        filStream = new StreamReader(filRead, Encoding.Default);
        statMessage = filStream.ReadToEnd();
        filStream.Close();
        filRead.Close();
      }

      if (!string.IsNullOrEmpty(statMessageFileMonthlyVal))
      {
        FileStream filReadM = null;
        StreamReader filStreamM = null;
        filReadM = new FileStream(statMessageFileMonthlyVal, FileMode.Open);
        filStreamM = new StreamReader(filReadM, Encoding.Default);
        statMessageMonthly = filStreamM.ReadToEnd();
        filStreamM.Close();
        filReadM.Close();
      }

      // open Summary file
      fileSummaryName = pStrFileName;
      fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
        "_SummaryEmail.txt";//clsBasFile.getFileExtn(fileSummaryName)
      fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
      streamSummary = new StreamWriter(fileSummary, Encoding.Default);

      reSetupTheValues();
      //FillStatementDataSet(pBankCode); 
      curAccountNo = String.Empty;
      //frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
        //frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
      foreach (DataRow mRow in DSstatement.Tables["Statement"].Rows)
      {
        if (mRow["Email"].ToString().Trim() == string.Empty)
          continue;
        masterRow = mRow;
        //mRow[myColumn] + "#,";
        printHeader(mRow);
        //cardsRows = mRow.GetChildRows("StaementNoDR"); //StatementNoDRel, DataRowVersion.Proposed
        cardsRows = DSstatement.Tables["Operation"].Select("StatementNo = '" + mRow["StatementNo"].ToString().Trim() + "'"); 
        if(cardsRows.Length < 0)
          continue;

        foreach (DataRow dRow in cardsRows) // mRow.GetChildRows(StatementNoDRel)
        {
          printDetail(dRow);
            //dRow[myColumn].ToString().Trim();
        }//dRow
        writeFooter(mRow);
        SendEmail(emailStr.ToString(), "", mRow["Email"].ToString());
      }//mRow    
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
      printStatementSummary();

      // Close Summary File
      streamSummary.Flush();
      streamSummary.Close();
      fileSummary.Close();

      streamEmails.Flush();
      streamEmails.Close();
      fileEmails.Close();

      streamNoEmails.Flush();
      streamNoEmails.Close();
      fileNoEmails.Close();
      aryLstFiles.Add(fileSummaryName);
      aryLstFiles.Add(strEmailFileName + ".txt");
      clsBasFile.generateFileMD5(aryLstFiles, strEmailFileName + ".MD5");
      aryLstFiles.Add(strEmailFileName + ".MD5");
      SharpZip zip = new SharpZip();
      zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".zip", "");

      DSstatement.Dispose();
    }
    return rtrnStr;
  }


  protected virtual void printHeader(DataRow pMainRow)
  {
    totTrans = 0; totComm = 0; totAmount = 0;
    emailStr = new StringBuilder("");
    emailStr.Append(@"<html><head><meta http-equiv='Content-Type' content='text/html; charset=windows-1252'><meta http-equiv='Content-Language' content='en-us'>");
    emailStr.Append(@"<title>" + validateEmpty(pMainRow["Client"].ToString()) + " - Statement" + @"</title></head>");
    emailStr.Append(@"<body background='cid:" + clsBasFile.getFileFromPath(strbackGround) + @"' leftmargin='0' rightmargin='0' bgcolor='#CACACA'>");
    emailStr.Append(@"<table id='table15' width='100%' border='0' cellpadding='0' height='100%'><tr><td colspan='2' height='82'><table id='table25' width='100%' border='0'><tr>");
    emailStr.Append(@"<td height='82' width='26%'><a href='http://" + strbankWebLink + @"'>");
    emailStr.Append(@"<img border='0' src='cid:" + clsBasFile.getFileFromPath(strBankLogo) + @"' width='103' height='104'></a></td>");
    emailStr.Append(@"<td height='82' align='center'><b><font size='4'>" + bankFullName + "</font></b></td>");
    emailStr.Append(@"<td height='82' width='141'>");
    emailStr.Append(@"<p align='right'><a href='http://" + strbankWebLink + @"'>");
    emailStr.Append(@"<img border='0' src='cid:" + clsBasFile.getFileFromPath(strVisaLogo) + @"' width='105' height='34'></a></p></td></tr></table></td></tr>");
    emailStr.Append(@"<tr><td width='50%' height='100%'><table id='table16' height='100%' width='100%' border='1'><tr><td><table border='0' width='100%' cellspacing='0' cellpadding='0' id='table37' height='100%'><tr><td width='100%' height='22'>");
    emailStr.Append(@"<p align='center'>" + fontTypeSize + validateEmpty(pMainRow["Client"].ToString()) + "</font></p>");
    emailStr.Append(@"</td></tr><tr><td height='23'>");
    emailStr.Append(@"<p align='center'>" + fontTypeSize + validateEmpty(pMainRow["Address"].ToString()) + "</font></p>");
    emailStr.Append(@"</td></tr><tr>");
    emailStr.Append(@"<td align='center' height='23'>" + fontTypeSize + validateEmpty(pMainRow["StreetAddress"].ToString()) + "</font></td>");
    emailStr.Append(@"</tr><tr>");
    emailStr.Append(@"<td align='center' height='20'>" + fontTypeSize + validateEmpty(pMainRow["City"].ToString() + pMainRow["Country"].ToString()) + "</font></td></tr></table></td></tr></table></td>");
    emailStr.Append(@"<td width='50%'><table id='table18' height='100%' width='100%' border='1'><tr><td height='105'><table id='table38' height='100' width='100%' border='0' cellspacing='0' cellpadding='0'><tr>");
    emailStr.Append(@"<td width='129' height='4'><b>" + fontTypeSize + "Account Number:</font></b></td>");
    emailStr.Append(@"<td height='19'>" + fontTypeSize + validateEmpty(pMainRow["ExternalAccount"].ToString()) + "</font></td>");
    emailStr.Append(@"</tr><tr><td width='129' height='19'><b>" + fontTypeSize + "Statement Date:</font></b></td>");
    emailStr.Append(@"<td height='19'>" + fontTypeSize + validateEmpty(Convert.ToDateTime(pMainRow["StatDate"]).ToString("dd/MM/yyyy")) + "</font></td>");
    emailStr.Append(@"</tr><tr><td width='129' height='20'><b>" + fontTypeSize + "Opening Balance:</font></b></td>");
    emailStr.Append(@"<td height='20'>" + fontTypeSize + validateEmpty(pMainRow["StartBalance"].ToString()) + "</font></td>");
    emailStr.Append(@"</tr><tr><td width='129' height='19'><b>" + fontTypeSize + "Closing Balance:</font></b></td>");
    emailStr.Append(@"<td height='19'>" + fontTypeSize + validateEmpty(pMainRow["EndBalance"].ToString()) + "</font></td>");
    emailStr.Append(@"</tr><tr><td width='129' height='19'><b>" + fontTypeSize + "Currency:</font></b></td>");
    emailStr.Append(@"<td height='19'>" + fontTypeSize + validateEmpty(pMainRow["CurrencyName"].ToString()) + "</font></td></tr></table></td></tr></table></td></tr>");
    emailStr.Append(@"<tr><td width='98%' colspan='2' height='19'>&nbsp;</td></tr>");
    emailStr.Append(@"<tr><td width='98%' colspan='2' height='19'>&nbsp;</td></tr>");
    emailStr.Append(@"<tr><td width='98%' colspan='2' height='88'><table id='table21' height='101%' width='100%' border='1' style='padding: 0'><tr>");
    emailStr.Append(@"<td align='middle' colspan='2' height='19' nowrap style='padding: 0'>" + fontTypeSize + "Date of</font></td>");
    emailStr.Append(@"<td align='middle' width='59%' height='40' rowspan='2' style='padding: 0'><b>" + fontTypeSize + "Description</font></b></td>");
    emailStr.Append(@"<td align='middle' width='11%' height='40' rowspan='2' style='padding: 0'>" + fontTypeSize + "Trans. Amount</font></td>");
    emailStr.Append(@"<td align='middle' width='6%' height='40' rowspan='2' style='padding: 0'>" + fontTypeSize + "Comm. Amount</font></td>");
    emailStr.Append(@"<td align='middle' width='10%' height='40' rowspan='2' style='padding: 0'>" + fontTypeSize + "Amount to Pay</font></td></tr>");
    emailStr.Append(@"<tr><td align='middle' width='6%' height='19' style='padding: 0'>" + fontTypeSize + "Trans.</font></td>");
    emailStr.Append(@"<td align='middle' width='6%' height='19' style='padding: 0'>" + fontTypeSize + "Posting</font></td></tr>");
    //emailStr.Append(@"<html><head><meta http-equiv='Content-Type' content='text/html; charset=windows-1252'><meta http-equiv='Content-Language' content='pt'>");
    //emailStr.Append(@"<a href='http://" + strbankWebLink + @"'>");
    //emailStr.Append(@"<img border='0' src='cid:" + clsBasFile.getFileFromPath(strBankLogo) + @"' width='228' height='95' align='right'></td>");
    //emailStr.Append(@"</tr></table></td></tr>");
    //emailStr.Append(@"<td height='25'>" + fontTypeSize + validateEmpty(masterRow[mCardbranchpartname].ToString()) + "</font></td></tr>");
  }

  protected virtual void printDetail(DataRow pDetailRow)
  {
    string strDesc=string.Empty;
    if (pDetailRow["S"].ToString() != "0")
      strDesc = " (" + pDetailRow["S"].ToString() + ") ";
    if  (pDetailRow["P"].ToString().Trim() != "")
      strDesc += " [" + pDetailRow["P"].ToString().Substring(0, 6) + "XXXXXX" + pDetailRow["P"].ToString().Substring(11, 4) + "]"; //ToText({Operation.P})
    strDesc = pDetailRow["DE"].ToString() + " ( " + pDetailRow["D"].ToString() + " )" + strDesc;

    //emailStr.Append(@"<tr><td align='middle' width='6%' height='23'>" + fontTypeSize + validateEmpty(trnsDate.ToString("dd/MM")) + "</font></td>");
    totNoOfTransactions++;
    emailStr.Append(@"<tr>");
    emailStr.Append(@"<td align='middle' width='6%' height='23' style='padding: 0'>" + fontTypeSize + validateEmpty(Convert.ToDateTime(pDetailRow["TD"].ToString()).ToString("dd/MM")) + "</font></td>");
    emailStr.Append(@"<td align='middle' width='6%' height='23' style='padding: 0'>" + fontTypeSize + validateEmpty(Convert.ToDateTime(pDetailRow["OD"].ToString()).ToString("dd/MM")) + "</font></td>");
    emailStr.Append(@"<td width='59%' height='23' style='padding: 0'><p align='left'>" + fontTypeSize + validateEmpty(strDesc) + "</font></p></td>");
    emailStr.Append(@"<td width='11%' height='23' style='padding: 0' align='right'>" + fontTypeSize + validateEmpty(pDetailRow["OC"].ToString()) + " " + basText.formatNum(pDetailRow["OA"], "#,###,##0.00") + "</font></td>"); //pDetailRow["OA"].ToString()
    emailStr.Append(@"<td width='6%' height='23' style='padding: 0' align='right'>" + fontTypeSize + validateEmpty(basText.formatNum(pDetailRow["Commission"], "#,###,##0.00")) + "</font></td>");//pDetailRow["Commission"].ToString()
    emailStr.Append(@"<td width='10%' height='23' style='padding: 0' align='right'>" + fontTypeSize + validateEmpty(basText.formatNum(pDetailRow["Amount2Pay"], "#,###,##0.00")) + "</font></td></tr>");//pDetailRow["Amount2Pay"].ToString() 
    totTrans += Convert.ToDecimal(pDetailRow["OA"].ToString());
    totComm += Convert.ToDecimal(pDetailRow["Commission"].ToString());
    totAmount += Convert.ToDecimal(pDetailRow["Amount2Pay"].ToString());
  }

  protected virtual void writeFooter(DataRow pMainRow)
  {
    emailStr.Append(@"<tr><td align='middle' width='6%' height='1' style='padding: 0'></td><td align='middle' width='6%' height='1' style='padding: 0'></td><td width='59%' height='1' style='padding: 0'></td><td width='11%' height='1' style='padding: 0'></td><td width='6%' height='1' style='padding: 0'></td><td width='10%' height='1' style='padding: 0'></td></tr>");
    emailStr.Append(@"<tr><td align='middle' width='71%' height='22' colspan='3' style='padding: 0'><p align='right'><b>" + fontTypeSize + "Totals</font></b></p></td>");
    emailStr.Append(@"<td width='11%' height='22' style='padding: 0'><p align='right'>" + validateEmpty(basText.formatNum(totTrans, "#,###,##0.00").Trim()) + "</td>");
    emailStr.Append(@"<td width='6%' height='22' style='padding: 0'><p align='right'>" + validateEmpty(basText.formatNum(totComm, "#,###,##0.00").Trim()) + "</td>");
    emailStr.Append(@"<td width='10%' height='22' style='padding: 0'><p align='right'>" + validateEmpty(basText.formatNum(totAmount, "#,###,##0.00").Trim()) + "</td></tr></table></td></tr>");
    emailStr.Append(@"<tr><td width='98%' colspan='2'>&nbsp;</td></tr>");
    emailStr.Append(@"<tr><td width='98%' colspan='2'>&nbsp;</td></tr>");
    emailStr.Append(@"<tr><td width='100%' colspan='2'><p align='center'><font color='#808080'>");
    emailStr.Append(@"<a title='http://" + strbankWebLink + @"' href='http://" + strbankWebLink + @"'>http://" + strbankWebLink + @"</a></font></p></td></tr>");
    emailStr.Append(@"<tr><td width='98%' colspan='2'>&nbsp;</td></tr></table></body></html>");
  }

  protected void calcAccountRows()
  {
  }


  protected string MakeEmailStr(string pStr)
  {
    pStr = pStr.Replace("\"","\"\"");
    return pStr;
  }

  protected string MakeHeaderStr(string pStr, bool isBold, bool isHeader)
  {
    string color = string.Empty;
    pStr = pStr.Replace("&" ,"&amp;").Trim();
    if (isHeader)
      color = "#000080";
    else
      color = "#800000";

    if (pStr.Length < 1)
      pStr = "&nbsp;";
    else
    {
      //pStr = @"<font size=""2"" color=""" + color + @""">" + pStr + @"</font>";
      pStr = @"<font size=""2"">" + pStr + @"</font>";
      if (isBold)
        pStr = @"<b>" + pStr + "</b>";
    }

    return pStr;
  }

  protected string ValidateStr(string pStr)
  {
    if (pStr.Trim().Length < 1)
      pStr = "&nbsp;";
    return pStr;
  }

  protected void SendEmail(string pBody, string pSubject, string pTo)
  {
      clsEmail sndMail = new clsEmail();
      try
      {

          if (!basText.isValideEmail(pTo))
          {
              streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|"+clsCnfg.readSetting("strValidEmail")?? "Bad Email");
              noOfBadEmails++;
              return;
          }


          //streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + strMobileNum.Trim() + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
          pLstTo.Clear();
          pSubject = emailLabel;
          if (pTo.EndsWith("."))
              pTo = pTo + "com";
          pLstTo.Add(pTo);//"mmohammed@emp-group.com"
          //return;0
          pLstBCC.Add("statement@emp-group.com");
          //pLstBCC.Add("nazab@emp-group.com");
          //clsEmail sndMail = new clsEmail();
          sndMail.emailFromName = emailFromNameStr;
          //if (pTo.ToLower().IndexOf("socgen.com") > 0)
          //  strEmailFromTmp = "nsgb@emp-group.com";
          //else
          strEmailFromTmp = strEmailFrom;

          numOfTry = 0;
          noOfEmails++;
          while (!sndMail.sendEmailAttachPic(strEmailFromTmp, pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
          {
              //MessageBox.Show("Failure to send Email.", "Send Email Error",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
              System.Threading.Thread.Sleep(waitPeriodVal);//2000
              numOfTry++;
              if (numOfTry > 100)
              {
                  streamEmails.Write("\t\t Error while Send Email");
                  streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Exceed number of trials");
                  noOfBadEmails++; noOfEmails--;
                  break;
              }
          }
          streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + strMobileNum.Trim() + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

          numMail++;
          if (numMail % 400 == 0)
          {
              System.Threading.Thread.Sleep(waitPeriodVal);//2000
              GC.Collect();
              GC.WaitForPendingFinalizers();
          }
          //sndMail = null;
          //emailStr = new StringBuilder(""); emailTo = strMobileNum = string.Empty; curAccountNumber = string.Empty; curCardNumber = string.Empty; curClientID = string.Empty;
          //System.Threading.Thread.Sleep(waitPeriodVal);//100
      }
      catch (Exception ex) //(NotSupportedException ex) (Exception ex)  //
      {
          //clsBasErrors.catchError(ex);
          streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "| Email " + pTo + " , Err Message >> " + ex.Message + ", Err Desc >> ");
          noOfBadEmails++;
      }
      finally
      {
          sndMail = null;
          emailStr = new StringBuilder(""); emailTo = strMobileNum = string.Empty; curAccountNumber = string.Empty; curCardNumber = string.Empty; curClientID = string.Empty;
          System.Threading.Thread.Sleep(waitPeriodVal);//100

      }
  }


  protected void printStatementSummary()
  {
    streamSummary.WriteLine(strBankName + " E-Statement Summary");
    streamSummary.WriteLine("__________________________");
    streamSummary.WriteLine("");
    streamSummary.WriteLine("Statements Sent by Email   " + noOfEmails.ToString());
    streamSummary.WriteLine("Statements have bad Email  " + noOfBadEmails.ToString());
    streamSummary.WriteLine("Statements without Email  " + noOfWithoutEmails.ToString());
//    streamSummary.WriteLine("No of Transactions " + totNoOfTransactions.ToString());
  }

  protected virtual void reSetupTheValues()
  {
  }

  // Default 7000
  public int waitPeriod
  {
    get { return waitPeriodVal; }
    set { waitPeriodVal = value; }
  }//waitPeriod

  public string statMessageFile
  {
    get { return statMessageFileVal; }
    set { statMessageFileVal = value; }
  }//statMessageFile

  public string statMessageFileMonthly
  {
    get { return statMessageFileMonthlyVal; }
    set { statMessageFileMonthlyVal = value; }
  }//statMessageFileMonthly

  public string visaLogo
  {
    get { return strVisaLogo; }
    set { 
      strVisaLogo = value;
      pLstAttachedFile.Add(strVisaLogo);//@"D:\pC#\exe\FilesForPrograms\frmBackground.jpg"
  }
  }//visaLogo

  public string backGround
  {
    get { return strbackGround; }
    set 
    { 
      strbackGround = value;
      pLstAttachedFile.Add(strbackGround);//@"D:\pC#\exe\FilesForPrograms\frmBackground.jpg"
    }
  }// backGround

  public string bankNameVal
  {
    get { return strBankName; }
    set { strBankName = value; }
  }// bankNameVal

  public string emailFrom
  {
    get { return strEmailFrom; }
    set { strEmailFrom = value; }
  }// emailFrom

  public string bankWebLinkService
  {
    get { return strbankWebLinkService; }
    set { strbankWebLinkService = value; }
  }// bankWebLinkService

  public string bankWebLink
  {
    get { return strbankWebLink; }
    set { strbankWebLink = value; }
  }// bankWebLink

  public string bankLogo
  {
    get { return strBankLogo; }
    set 
    { 
      strBankLogo = value;
      pLstAttachedFile.Add(strBankLogo);//@"D:\Web\Email\BAI\Logo.gif"
    }
  }// bankLogo

  //public frmStatementFile setFrm
  //{
  //  set { frmMain = value; }
  //}// setFrm

  protected string validateEmpty(string pStr)
  {
    if (pStr.Length < 1)
      pStr = "&nbsp;";
    return pStr;
  }

  public string logoAlignment
  {
    set { logoAlignmentStr = value; }
  }// logoAlignment

  public string emailFromName
  {
    get { return emailFromNameStr; }
    set { emailFromNameStr = value; }
  }// emailFromName

  ~clsStatHtmlMrch()
  {
  }
}

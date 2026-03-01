using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;

//
public class clsMaintainData : clsBasStatement //clsBasStatementFunc
{
    protected string RewardCondStr = "'Reward Program (Airmile)'";
    protected bool notRwardVal = true;
    protected string installmentCond = string.Empty;
    protected bool isInstallmentVal = false;
    private const int MaxNoTrans = 500;//49
    protected bool withoutMsgBoxVal = false;

    public clsMaintainData()
    {
    }



    public int matchCardBranch4Account(int pBankCode)
    {
        string curClientId = "";
        int rtrnVal = 0;
        int clientId = 0, curBranchCode = 0, branchCode = 0, curBankCode = 0;
        string curBranchName = string.Empty, upateSql = string.Empty;
        bool needChange = false;
        long SqlCnt = 0;
        string strSqlActn = string.Empty;

        DataSet ds = null;
        string masterQuery;
        OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
        OracleDataAdapter myDataAdapter;
        conn.Open();
        if (notRwardVal)
        {
            //upateSql = "delete /*+ parallel (" + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + ",4) */ from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.cardno is null and branch =" + pBankCode + " and branch != 25";
            upateSql = "delete /*+ parallel (" + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + ",4) */ from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.cardno is null and branch =" + pBankCode + " and contracttype != " + RewardCondStr;
            rtrnVal = clsDbOracleLayer.doAction(upateSql);
        }

        if (isInstallmentVal)
        {
            upateSql = "delete /*+ parallel (" + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + ",4) */ from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.cardno is null and branch =" + pBankCode + " and contracttype not in " + installmentCond;
            rtrnVal = clsDbOracleLayer.doAction(upateSql);
        }


        masterQuery = "select /*+ index (" + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " iBranchTstatementmastertable) */ STATEMENTNO, branch, clientid, cardcreationdate, cardbranchpart, cardbranchpartname from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " where branch = " + pBankCode + " order by branch, clientid, cardcreationdate";



        myDataAdapter = new OracleDataAdapter(masterQuery, conn);
        ds = new DataSet("MasterDetailDS");
        myDataAdapter.Fill(ds, "MasterDetailDS");

        strSqlActn = "begin ";
        foreach (DataRow supRow in ds.Tables["MasterDetailDS"].Rows) //DataRow    ds.Tables["tStatementMasterTable"].Rows    mRow.GetChildRows(StatementNoDRel)
        {
            if (curClientId == supRow[mClientid].ToString())
            {
                if (supRow[mCardbranchpart].ToString() != "" && curBranchCode != Convert.ToInt32(supRow[mCardbranchpart].ToString()))
                {
                    branchCode = Convert.ToInt32(supRow[mCardbranchpart].ToString());
                    curBranchName = supRow[mCardbranchpartname].ToString();
                    curBankCode = Convert.ToInt32(supRow[mBranch].ToString());
                    clientId = Convert.ToInt32(supRow[mClientid].ToString());
                    needChange = true;
                }
            }
            else
            {
                if (needChange)
                {
                    upateSql = "update " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " set " +
                      " cardbranchpart = '" + basText.validateWriteField(branchCode.ToString()) + "', " +
                      " cardbranchpartname = '" + basText.validateWriteField(curBranchName.ToString()) + "' " +
                      " where branch = " + curBankCode +
                      " and clientid = " + clientId;
                    //clsDbOracleLayer.doAction(@upateSql);
                    strSqlActn += upateSql + ";";
                    SqlCnt++;
                    if (SqlCnt > MaxNoTrans)
                    {
                        strSqlActn = strSqlActn + " end;";
                        (new OracleCommand(strSqlActn, conn)).ExecuteNonQuery();
                        strSqlActn = "begin "; SqlCnt = 0;
                    }

                    rtrnVal++;
                }
                needChange = false;
                //curBranchCode = 0;
            }
            curClientId = supRow[mClientid].ToString();
            if (supRow[mCardbranchpart].ToString() != "")
                curBranchCode = Convert.ToInt32(supRow[mCardbranchpart].ToString());
        } // end of supRow
        if (needChange)
        {
            upateSql = "update " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " set " +
              " cardbranchpart = '" + basText.validateWriteField(branchCode.ToString()) + "', " +
              " cardbranchpartname = '" + basText.validateWriteField(curBranchName.ToString()) + "' " +
              " where branch = " + curBankCode +
              " and clientid = " + clientId;
            //clsDbOracleLayer.doAction(@upateSql);
            strSqlActn += upateSql + ";";
            rtrnVal++;
        }
        if (SqlCnt > 0)
        {
            strSqlActn = strSqlActn + " end;";
            (new OracleCommand(strSqlActn, conn)).ExecuteNonQuery();
        }

        ds.Dispose();
        conn.Close();
        GC.Collect();
        GC.WaitForPendingFinalizers();

        return rtrnVal;
    }

    public int fixArbicAddress(int pBankCode)
    {
        int rtrnVal = 0;
        string curBranchName = string.Empty, upateSql = string.Empty;
        DataSet ds = null;
        string masterQuery;
        long SqlCnt = 0;
        string strSqlActn = string.Empty;

        OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
        OracleDataAdapter myDataAdapter;
        conn.Open();

        //masterQuery = "select branch, STATEMENTNO, customeraddress1,customeraddress2,customeraddress3,accountaddress1,accountaddress2,accountaddress3,cardaddress1,cardaddress2,cardaddress3 from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " where branch = " + pBankCode; // where branch = 6 and STATEMENTNO = 342 where branch = 10
        //masterQuery = "select branch, statementnumber, customeraddress1,customeraddress2,customeraddress3,accountaddress1,accountaddress2,accountaddress3,cardaddress1,cardaddress2,cardaddress3 from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " where branch = " + pBankCode + " order by statementnumber"; // where branch = 6 and STATEMENTNO = 342 where branch = 10
        masterQuery = "select branch, statementnumber, customeraddress1,customeraddress2,customeraddress3,accountaddress1,accountaddress2,accountaddress3,cardaddress1,cardaddress2,cardaddress3 from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " where branch = " + pBankCode + "  and STATEMENTNUMBER IN (17655,17658) order by statementnumber";
        myDataAdapter = new OracleDataAdapter(masterQuery, conn);
        ds = new DataSet("MasterDetailDS");
        myDataAdapter.Fill(ds, "MasterDetailDS");

        strSqlActn = "begin ";
        foreach (DataRow supRow in ds.Tables["MasterDetailDS"].Rows) //DataRow    ds.Tables["tStatementMasterTable"].Rows    mRow.GetChildRows(StatementNoDRel)
        {
            if (supRow[mCustomeraddress1].ToString().Trim().Length > 3 && supRow[mCustomeraddress1].ToString().Trim().Substring(0, 3) == "???"
              || supRow[mCustomeraddress2].ToString().Trim().Length > 3 && supRow[mCustomeraddress2].ToString().Trim().Substring(0, 3) == "???"
              || supRow[mCustomeraddress3].ToString().Trim().Length > 3 && supRow[mCustomeraddress3].ToString().Trim().Substring(0, 3) == "???")
            {
                upateSql = "update " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " set " +
                  " customeraddress1 = '" + basText.validateWriteField(ValidateArbic(supRow[mCustomeraddress1].ToString().Trim())) + "', " +
                  " customeraddress2 = '" + basText.validateWriteField(ValidateArbic(supRow[mCustomeraddress2].ToString().Trim())) + "', " +
                  " customeraddress3 = '" + basText.validateWriteField(ValidateArbic(supRow[mCustomeraddress3].ToString().Trim())) + "', " +
                  " accountaddress1 = '" + basText.validateWriteField(ValidateArbic(supRow["accountaddress1"].ToString().Trim())) + "', " +
                  " accountaddress2 = '" + basText.validateWriteField(ValidateArbic(supRow["accountaddress2"].ToString().Trim())) + "', " +
                  " accountaddress3 = '" + basText.validateWriteField(ValidateArbic(supRow["accountaddress3"].ToString().Trim())) + "', " +
                  " cardaddress1 = '" + basText.validateWriteField(ValidateArbic(supRow["cardaddress1"].ToString().Trim())) + "', " +
                  " cardaddress2 = '" + basText.validateWriteField(ValidateArbic(supRow["cardaddress2"].ToString().Trim())) + "', " +
                  " cardaddress3 = '" + basText.validateWriteField(ValidateArbic(supRow["cardaddress3"].ToString().Trim())) + "' " +
                  " where branch = " + supRow[mBranch] +
                  //" and STATEMENTNO = " + supRow["STATEMENTNO"];
                  " and statementnumber = " + supRow["statementnumber"];
                //clsDbOracleLayer.doAction(@upateSql);
                strSqlActn += upateSql + ";";
                SqlCnt++;
                if (SqlCnt > MaxNoTrans)
                {
                    strSqlActn = strSqlActn + " end;";
                    (new OracleCommand(strSqlActn, conn)).ExecuteNonQuery();
                    strSqlActn = "begin "; SqlCnt = 0;
                }

            }
        } // end of supRow
        if (SqlCnt > 0)
        {
            strSqlActn = strSqlActn + " end;";
            (new OracleCommand(strSqlActn, conn)).ExecuteNonQuery();
        }
        ds.Dispose();
        conn.Close();
        return rtrnVal;
    }

    public int fixAddress(int pBankCode)
    {
        int rtrnVal = 0;
        string upateSql = string.Empty;
        DataSet ds = null;
        string masterQuery;
        string originalString;
        string newaddr1, newaddr2;
        OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
        OracleDataAdapter myDataAdapter;
        conn.Open();

        //masterQuery = "select branch, statementno,statementnumber,customeraddress1,customeraddress2 from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " where branch = " + pBankCode + " and length(customeraddress1) > 50 order by statementno";
        //masterQuery = "select branch, statementno,statementnumber,customeraddress1,customeraddress2 from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " where branch = " + pBankCode + " and length(customeraddress1) > 50 and customeraddress2 is null order by statementno";
        masterQuery = "select distinct  branch ,statementnumber,customeraddress1 from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " where branch = " + pBankCode + " and length(customeraddress1) > 50 and customeraddress2 is null order by statementnumber";
        myDataAdapter = new OracleDataAdapter(masterQuery, conn);
        try
        {
            ds = new DataSet("MasterDetailDS");
            myDataAdapter.Fill(ds, "MasterDetailDS");
            foreach (DataRow supRow in ds.Tables["MasterDetailDS"].Rows) //DataRow    ds.Tables["tStatementMasterTable"].Rows    mRow.GetChildRows(StatementNoDRel)
            {
                originalString = ValidateArbic(supRow[mCustomeraddress1].ToString());
                newaddr1 = newaddr2 = string.Empty;
                string[] splitArray = originalString.Split(' ');
                for (int i = 0; i < splitArray.Length; i++)
                {
                    newaddr1 += splitArray[i] + " ";
                    if (newaddr1.Length > 50)
                    {
                        newaddr1 = newaddr1.Remove(newaddr1.Length - splitArray[i].Length - 1, splitArray[i].Length);
                        break;
                    }
                }
                newaddr2 = originalString.Substring(newaddr1.Length - 1);
                newaddr1 = newaddr1.Replace("'", "''");
                newaddr2 = newaddr2.Replace("'", "''");
                upateSql = "update " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " set " +
                  " customeraddress1 = '" + newaddr1 + "', " +
                  " customeraddress2 = '" + newaddr2 + "', " +
                  " accountaddress1 = '" + newaddr1 + "', " +
                  " accountaddress2 = '" + newaddr2 + "', " +
                  " cardaddress1 = '" + newaddr1 + "', " +
                  " cardaddress2 = '" + newaddr2 + "'" +
                  " where branch = " + supRow[mBranch] +
                  //" and STATEMENTNO = " + supRow[mStatementno] +
                  " and statementnumber = " + supRow[mStatementnumber] +
                  "  and customeraddress2 is null" +
                  "  and accountaddress2 is null" +
                  "  and cardaddress2 is null";
                ;
                (new OracleCommand(upateSql, conn)).ExecuteNonQuery();
            }
            ds.Dispose();
            conn.Close();
        }
        catch (Exception ex)
        {
            clsBasErrors.catchError(ex);
        }
        return rtrnVal;
    }

    public static void fixAddress(int pBankCode, string originalString, out string newaddr1, out string newaddr2)
    {
        newaddr1 = newaddr2 = string.Empty;
        try
        {
            //originalString = ValidateArbic(supRow[mCustomeraddress1].ToString());
            string[] splitArray = originalString.Split(' ');
            for (int i = 0; i < splitArray.Length; i++)
            {
                newaddr1 += splitArray[i] + " ";
                if (newaddr1.Length > 50)
                {
                    newaddr1 = newaddr1.Remove(newaddr1.Length - splitArray[i].Length - 1, splitArray[i].Length);
                    break;
                }
            }
            newaddr2 = originalString.Substring(newaddr1.Length - 1);
        }
        catch (Exception ex)
        {
            clsBasErrors.catchError(ex);
        }
    }

    public int fixArbicAddressLang(int pBankCode)
    {
        int rtrnVal = 0;
        string curBranchName = string.Empty, upateSql = string.Empty;
        DataSet ds = null;
        string masterQuery;
        long SqlCnt = 0;
        string strSqlActn = string.Empty;

        OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
        OracleDataAdapter myDataAdapter;
        conn.Open();

        //masterQuery = "select branch, STATEMENTNO, customeraddress1,customeraddress2,customeraddress3,accountaddress1,accountaddress2,accountaddress3,cardaddress1,cardaddress2,cardaddress3 from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " where branch = " + pBankCode; // where branch = 6 and STATEMENTNO = 342 where branch = 10
        masterQuery = "select branch, statementnumber, customeraddress1,customeraddress2,customeraddress3,accountaddress1,accountaddress2,accountaddress3,cardaddress1,cardaddress2,cardaddress3 from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " where branch = " + pBankCode + " order by statementnumber"; // where branch = 6 and STATEMENTNO = 342 where branch = 10
        myDataAdapter = new OracleDataAdapter(masterQuery, conn);
        ds = new DataSet("MasterDetailDS");
        myDataAdapter.Fill(ds, "MasterDetailDS");

        strSqlActn = "begin ";
        foreach (DataRow supRow in ds.Tables["MasterDetailDS"].Rows) //DataRow    ds.Tables["tStatementMasterTable"].Rows    mRow.GetChildRows(StatementNoDRel)
        {
            if (supRow[mCustomeraddress1].ToString().Trim().Length > 3 && supRow[mCustomeraddress1].ToString().Trim().Substring(0, 3) == "ÑÞã")
            {
                upateSql = "update " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " set " +
                  " customeraddress1 = '" + basText.validateWriteField(ValidateArbic(supRow[mCustomeraddress1].ToString().Trim())) + "', " +
                  " customeraddress2 = '" + basText.validateWriteField(ValidateArbic(supRow[mCustomeraddress2].ToString().Trim())) + "', " +
                  " customeraddress3 = '" + basText.validateWriteField(ValidateArbic(supRow[mCustomeraddress3].ToString().Trim())) + "', " +
                  " accountaddress1 = '" + basText.validateWriteField(ValidateArbic(supRow["accountaddress1"].ToString().Trim())) + "', " +
                  " accountaddress2 = '" + basText.validateWriteField(ValidateArbic(supRow["accountaddress2"].ToString().Trim())) + "', " +
                  " accountaddress3 = '" + basText.validateWriteField(ValidateArbic(supRow["accountaddress3"].ToString().Trim())) + "', " +
                  " cardaddress1 = '" + basText.validateWriteField(ValidateArbic(supRow["cardaddress1"].ToString().Trim())) + "', " +
                  " cardaddress2 = '" + basText.validateWriteField(ValidateArbic(supRow["cardaddress2"].ToString().Trim())) + "', " +
                  " cardaddress3 = '" + basText.validateWriteField(ValidateArbic(supRow["cardaddress3"].ToString().Trim())) + "' " +
                  " where branch = " + supRow[mBranch] +
                  //" and STATEMENTNO = " + supRow["STATEMENTNO"];
                  " and statementnumber = " + supRow["statementnumber"];
                //clsDbOracleLayer.doAction(@upateSql);
                strSqlActn += upateSql + ";";
                SqlCnt++;
                if (SqlCnt > MaxNoTrans)
                {
                    strSqlActn = strSqlActn + " end;";
                    (new OracleCommand(strSqlActn, conn)).ExecuteNonQuery();
                    strSqlActn = "begin "; SqlCnt = 0;
                }
            }
            if (ValidateIsArbic(supRow[mCustomeraddress1].ToString().Trim()) == true)
            {
                upateSql = "update " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t set t.companycode = 1 where branch = " + supRow[mBranch] +
                //" and STATEMENTNO = " + supRow["STATEMENTNO"];
                " and statementnumber = " + supRow["statementnumber"];
            }
            else
            {
                upateSql = "update " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t set t.companycode = 0 where branch = " + supRow[mBranch] +
                //" and STATEMENTNO = " + supRow["STATEMENTNO"];
                " and statementnumber = " + supRow["statementnumber"];
            }
            //clsDbOracleLayer.doAction(@upateSql);
            strSqlActn += upateSql + ";";
            SqlCnt++;
            if (SqlCnt > MaxNoTrans)
            {
                strSqlActn = strSqlActn + " end;";
                (new OracleCommand(strSqlActn, conn)).ExecuteNonQuery();
                strSqlActn = "begin "; SqlCnt = 0;
            }
        } // end of supRow
        if (SqlCnt > 0)
        {
            strSqlActn = strSqlActn + " end;";
            (new OracleCommand(strSqlActn, conn)).ExecuteNonQuery();
        }

        ds.Dispose();
        conn.Close();
        return rtrnVal;
    }




    public void deleteOnHoldTrans(int pBranch, bool isReward)
    {
        Cursor.Current = Cursors.WaitCursor;
        string upateSql = string.Empty, suplCond = string.Empty;
        if (isReward)
        {
            suplCond = " and d.trandescription != 'Calculated Points'";
        }
        try
        {
            upateSql = "delete FROM " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " d where d.branch = " + pBranch + " and POSTINGDATE is null and DOCNO is null " + suplCond;
            clsDbOracleLayer.doAction(upateSql);
            return;

        }
        catch
        {
            //clsDbOracleLayer.catchError(ex);
        }
        finally
        {
            Cursor.Current = Cursors.Default;
        }
    }




    public void mergeMarkUpFees(int pBranch)
    {
        string masterQuery;
        string upateSql = string.Empty;
        string docNo = string.Empty;
        string mainRowId = string.Empty, supRowId = string.Empty;
        decimal totTrans = 0;
        long SqlCnt = 0;
        string strSqlActn = string.Empty;

        Cursor.Current = Cursors.WaitCursor;
        masterQuery = @"select t.rowid,t.* from " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " t where t.branch = " + pBranch + " and t.refereneno != ' ' and t.docno in(SELECT x.docno FROM " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " x where x.branch = " + pBranch + " and x.trandescription like '%Mark-Up Fee%' GROUP BY x.docno) and t.trandescription not like '%International%' order by t.docno, t.merchant desc";
        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter myDataAdapter = new OracleDataAdapter(masterQuery, conn);
            conn.Open();
            DataSet ds = new DataSet("mergeMarkUpFees");
            myDataAdapter.Fill(ds, "mergeMarkUpFees");

            strSqlActn = "begin ";
            foreach (DataRow supRow in ds.Tables["mergeMarkUpFees"].Rows) //DataRow    ds.Tables["tStatementMasterTable"].Rows    mRow.GetChildRows(StatementNoDRel)
            {
                if (docNo == supRow[dDocno].ToString())
                {
                    supRowId = supRow["rowid"].ToString();
                    totTrans = totTrans + Convert.ToDecimal(supRow[dBilltranamount].ToString());
                    upateSql = "delete from " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " where rowid = '" + supRowId + "'";
                    strSqlActn += upateSql + ";";
                    SqlCnt++;
                }
                else if (docNo != supRow[dDocno].ToString())
                {
                    if (docNo != "")
                    {
                        upateSql = "update " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " set billtranamount = " + totTrans + " where rowid = '" + mainRowId + "'";
                        strSqlActn += upateSql + ";";
                        SqlCnt++;
                    }
                    totTrans = 0; supRowId = "";
                    mainRowId = supRow["rowid"].ToString();
                    totTrans = Convert.ToDecimal(supRow[dBilltranamount].ToString());
                }
                docNo = supRow[dDocno].ToString();
                if (SqlCnt > MaxNoTrans)
                {
                    strSqlActn = strSqlActn + " end;";
                    (new OracleCommand(strSqlActn, conn)).ExecuteNonQuery();
                    strSqlActn = "begin "; SqlCnt = 0;
                }
            } // end of supRow
            if (supRowId != "")
            {
                upateSql = "update " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " set billtranamount = " + totTrans + " where rowid = '" + mainRowId + "'";
                strSqlActn += upateSql + ";";
                SqlCnt++;
            }
            if (SqlCnt > 0)
            {
                strSqlActn = strSqlActn + " end;";
                (new OracleCommand(strSqlActn, conn)).ExecuteNonQuery();
            }

            conn.Close();
        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
        finally
        {
            Cursor.Current = Cursors.Default;
        }
    }


    public int makeBranchAsMainCard(int pBankCode)
    {
        int rtrnVal = 0;
        string curBranchName = string.Empty, upateSql = string.Empty;
        string curMainCard, CurCardNo;
        DataSet ds = null;
        DataRow[] mainRows;
        string masterQuery, curBranchPart, branchPart, branchPartName;
        bool needChange;
        long SqlCnt = 0;
        string strSqlActn = string.Empty;

        Cursor.Current = Cursors.WaitCursor;
        OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
        OracleDataAdapter myDataAdapter;
        try
        {
            conn.Open();

            //masterQuery = "select branch, STATEMENTNO, accountno, cardno, cardprimary, cardstate, cardbranchpart, cardbranchpartname from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " where branch = " + pBankCode; // where branch = 6 and STATEMENTNO = 342 where branch = 10
            masterQuery = "select branch, STATEMENTNO, accountno, cardno, cardprimary, cardstate, cardbranchpart, cardbranchpartname from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " where branch = " + pBankCode + " and CARDNO in ('4644800374601505','4644800200079231')";
            myDataAdapter = new OracleDataAdapter(masterQuery, conn);
            ds = new DataSet("MasterDetailDS");
            myDataAdapter.Fill(ds, "MasterDetailDS");
            conn.Close();
            strSqlActn = "begin ";
            foreach (DataRow supRow in ds.Tables["MasterDetailDS"].Rows) //DataRow    ds.Tables["tStatementMasterTable"].Rows    mRow.GetChildRows(StatementNoDRel)
            {
                mainRows = ds.Tables["MasterDetailDS"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(supRow[mAccountno]) + "'");
                curMainCard = CurCardNo = curBranchPart = branchPart = branchPartName = "";
                needChange = false;
                foreach (DataRow mainRow in mainRows) //mRow.GetChildRows(StatementNoDRel)
                {
                    CurCardNo = mainRow[mCardno].ToString();
                    if (branchPart == "")
                    {
                        branchPart = mainRow[mCardbranchpart].ToString();
                        branchPartName = mainRow[mCardbranchpartname].ToString();
                    }
                    if (mainRow[mCardprimary].ToString() == "Y")
                        if (curMainCard == "")
                        {
                            curMainCard = CurCardNo; //mainRow[mCardno].ToString();
                            branchPart = mainRow[mCardbranchpart].ToString();
                            branchPartName = mainRow[mCardbranchpartname].ToString();
                        }
                    if (mainRow[mCardprimary].ToString() == "Y" && isValidateCard(mainRow[mCardstate].ToString()))
                    {
                        curMainCard = CurCardNo; //mainRow[mCardno].ToString();
                        branchPart = mainRow[mCardbranchpart].ToString();
                        branchPartName = mainRow[mCardbranchpartname].ToString();
                    }
                    if (curBranchPart != "" && curBranchPart != mainRow[mCardbranchpart].ToString())
                        needChange = true;
                    curBranchPart = mainRow[mCardbranchpart].ToString();
                }

                if (curMainCard == "")
                    curMainCard = CurCardNo;

                if (needChange)
                {
                    upateSql = "update " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " set " +
                        " cardbranchpart = '" + basText.ReplaceSpecialChr(branchPart) + "' " +
                        ", cardbranchpartname = '" + basText.ReplaceSpecialChr(branchPartName) + "' " +
                        " where branch = " + supRow[mBranch] +
                        " and ACCOUNTNO = '" + clsBasValid.validateStr(supRow[mAccountno]) + "'";
                    //clsDbOracleLayer.doAction(@upateSql);
                    strSqlActn += upateSql + ";";
                    SqlCnt++;
                    if (SqlCnt > MaxNoTrans)
                    {
                        strSqlActn = strSqlActn + " end;";
                        conn.Open();
                        (new OracleCommand(strSqlActn, conn)).ExecuteNonQuery();
                        conn.Close();
                        strSqlActn = "begin "; SqlCnt = 0;
                    }

                }
            } // end of supRow
            if (SqlCnt > 0)
            {
                strSqlActn = strSqlActn + " end;";
                conn.Open();
                (new OracleCommand(strSqlActn, conn)).ExecuteNonQuery();
                conn.Close();
            }

        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
        //    catch (NotSupportedException ex)  //(Exception ex)  //
        //    {
        //      //clsBasErrors.catchError(ex);
        //    }
        finally
        {
            ds.Dispose();
            conn.Close();
            Cursor.Current = Cursors.Default;
        }
        return rtrnVal;
    }



    protected string ValidateArbic(string pStrArbic)
    {
        string rtnStr = string.Empty, tmpStr = string.Empty;
        if (pStrArbic == null || pStrArbic.Trim() == "")
            rtnStr = "";
        else
        {
            rtnStr = pStrArbic;
            if (pStrArbic.Length > 3 && pStrArbic.Substring(0, 3) == "ÑÞã")
                rtnStr = pStrArbic.Substring(3);
            if (basText.isContainArbic(rtnStr))
                rtnStr = pStrArbic;
        }
        rtnStr = rtnStr.Trim();
        return rtnStr;
    }

    protected bool ValidateIsArbic(string pStrArbic)
    {
        string rtnStr = string.Empty, tmpStr = string.Empty;
        bool rtrnBool = false;
        if (pStrArbic == null || pStrArbic.Trim() == "")
            rtnStr = "";
        else
        {
            rtnStr = pStrArbic;
            if (pStrArbic.Length > 3 && pStrArbic.Substring(0, 3) == "ÑÞã")
                rtnStr = pStrArbic.Substring(3);
            if (basText.isContainArbic(rtnStr))
                rtnStr = pStrArbic;
        }
        rtnStr = rtnStr.Trim();
        if (basText.isContainArbic(rtnStr))
            rtrnBool = true;
        return rtrnBool;
    }


    public bool notRward //static public 
    {
        get { return notRwardVal; }
        set { notRwardVal = value; }
    }// notRward

    public bool isInstallment //static public 
    {
        get { return isInstallmentVal; }
        set { isInstallmentVal = value; }
    }

    public string curRewardCond //static public 
    {
        get { return RewardCondStr; }
        set { RewardCondStr = value; }
    }// curRewardCond

    public string curInstallementCond //static public 
    {
        get { return installmentCond; }
        set { installmentCond = value; }
    }


    ~clsMaintainData()
    {
    }


}

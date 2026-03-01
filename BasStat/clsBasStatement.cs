using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
//using Oracle.DataAccess.Client;
using Oracle.DataAccess.Client;
using System.Collections;

// Base Data for Statement
public class clsBasStatement : clsBasStatementFunc
{
    //protected DataSet StatementDataset;

    // statement tables fields
    protected const string mBranch = "branch";
    protected const string mStatementno = "statementno";
    protected const string mStatementnumber = "statementnumber";
    protected const string mCardno = "cardno";
    protected const string mMBR = "mbr";
    protected const string mCardproduct = "cardproduct";
    protected const string mCardprimary = "cardprimary";
    protected const string mPrinarycardno = "prinarycardno";
    protected const string mAccountno = "accountno";
    protected const string mStatementdatefrom = "statementdatefrom";
    protected const string mStatementdateto = "statementdateto";
    protected const string mStetementduedate = "stetementduedate";
    protected const string mCardstate = "cardstate";
    protected const string mCardstatus = "cardstatus";
    protected const string mExternalno = "externalno";
    protected const string mCustomername = "customername";
    protected const string mCardbranchpart = "cardbranchpart";
    protected const string mCardbranchpartname = "cardbranchpartname";
    protected const string mCustomeraddress1 = "customeraddress1";
    protected const string mCustomeraddress2 = "customeraddress2";
    protected const string mCustomeraddress3 = "customeraddress3";
    protected const string mCardaddress1 = "CARDADDRESS1";
    protected const string mCardaddress2 = "CARDADDRESS2";
    protected const string mCardaddress3 = "CARDADDRESS3";
    protected const string mCustomerregion = "customerregion";
    protected const string mCustomercity = "customercity";
    protected const string mCustomerzipcode = "customerzipcode";
    protected const string mAccountlim = "accountlim";
    protected const string mAccountavailablelim = "accountavailablelim";
    protected const string mCardlimit = "cardlimit";
    protected const string mCardavailablelimit = "cardavailablelimit";
    protected const string mTotaloverdueamount = "totaloverdueamount";
    protected const string mTotalDueAmount = "totaldueamount";
    protected const string mTotaldebits = "totaldebits";
    protected const string mTotalcredits = "totalcredits";
    protected const string mTotalpayments = "totalpayments";
    protected const string mTotalpurchases = "totalpurchases";
    protected const string mTotalcashwithdrawal = "totalcashwithdrawal";
    protected const string mTotalcharges = "totalcharges";
    protected const string mTotalinterest = "totalinterest";
    protected const string mMindueamount = "mindueamount";
    protected const string mOpeningbalance = "openingbalance";
    protected const string mClosingbalance = "closingbalance";
    protected const string mAccountcurrency = "accountcurrency";
    protected const string mAccounttype = "accounttype";
    protected const string mStatementmessageline1 = "statementmessageline1";
    protected const string mStatementmessageline2 = "statementmessageline2";
    protected const string mStatementmessageline3 = "statementmessageline3";
    protected const string mContracttype = "contracttype";
    protected const string mClientid = "clientid";
    protected const string mAccountzipcode = "accountzipcode";
    protected const string mCardaddressbarcode = "cardaddressbarcode";
    protected const string mCardpaymentmethod = "cardpaymentmethod";
    protected const string mCardexpirydate = "cardexpirydate";
    protected const string mCustomercountry = "customercountry";
    protected const string mCardaccountno = "cardaccountno";
    protected const string mContractno = "contractno";
    protected const string mAccountstatus = "accountstatus";
    protected const string mCardclientname = "cardclientname";
    protected const string mContractlimit = "contractlimit";
    protected const string mCustomertitle = "customertitle";
    protected const string mDept = "dept";
    protected const string mCardtitle = "cardtitle";
    protected const string mContactpersonename = "contactpersonename";
    protected const string mCardvip = "cardvip";
    protected const string mCarddafamount = "carddafamount";
    protected const string mDAFPercentage = "dafpercentage";
    protected const string mCardClientId = "cardclientid";
    protected const string mHOLSTMT = "holstmt";
    protected const string mBarcode = "barcode";
    protected const string mUserActField1 = "useractfield1";
    protected const string mEarnedBonus = "earnedbonus";
    protected const string mRedeemedBonus = "redeemedbonus";
    protected const string mExpiredBonus = "expiredbonus";
    protected const string mBonusAdjustment = "bonusadjustment";
    protected const string mExpiredBonusNextMonth = "expiredbonusnextmonth";
    protected const string mExpiredBonusDate = "expiredbonusdate";
    protected const string mIntSpentLimit = "intspentlimit";
    protected const string mMinPayPercentage = "minpaypercentage";
    protected const string mMonthsCount = "monthscount";
    protected const string mPackageName = "packagename";
    protected const string mInstallmentUsedLimit = "installmentusedlimit";
    protected const string mOverduedays = "OVERDUEDAYS";
    protected const string mcontractstate = "CONTRACTSTATE";
    protected const string mCreditcontracts = "CREDITCONTRACTS";
    protected const string dStatementno = "statementno";
    protected const string dStatementnumber = "statementnumber";
    protected const string dAccountno = "accountno";
    protected const string dAccountcurrency = "accountcurrency";
    protected const string dCardno = "cardno";
    protected const string dTransdate = "transdate";
    protected const string dPostingdate = "postingdate";
    protected const string dTrandescription = "trandescription";
    protected const string dMerchant = "merchant";
    protected const string dOrigtranamount = "origtranamount";
    protected const string dOrigtrancurrency = "origtrancurrency";
    protected const string dBilltranamount = "billtranamount";
    protected const string dBilltranamountsign = "billtranamountsign";
    protected const string dDocno = "docno";
    protected const string dRefereneno = "refereneno";
    protected const string dInstallmentData = "installmentdata";
    protected const string dOrigDocNo = "origdocno";
    protected const string dInstallmentMerchant = "installmentmerchant";
    protected const string dInstallmentMerchantLocation = "installmentmerchantlocation";
    protected const string dInstallmentStartDate = "installmentstartdate";
    protected const string dInstallmentLoan = "installmentloan";
    protected const string dInstallmentCycleNo = "installmentcycleno";
    protected const string dInstallmentCycles = "installmentcycles";
    protected const string dInstallmentRegRePayment = "installmentregrepayment";
    protected const string dInstallmentEndDate = "installmentenddate";
    protected const string dInstallmentOutBalance = "installmentoutbalance";

    protected const string dEntryNo = "entryno";
    protected const string dPackageName = "packagename";
    protected const string dApprovalCode = "approvalcode";

    protected const string cPan = "pan";
    protected const string cMBR = "mbr";
    protected const string cBranch = "branch";
    protected const string cClientid = "idclient";

    protected const string pCode = "code";
    protected const string pName = "name";

    protected const string bBranch = "branch";
    protected const string bBranchPart = "branchpart";
    protected const string bIdent = "ident";

    protected DataRelation StatementNoDRel;
    protected DataRelation ClientIDRel;
    protected int curBranch = 0;//static
    protected string curProduct = string.Empty;
    protected bool isCorporate = false, isCredite = true;
    protected string strOrder = " m.cardproduct,m.CARDBRANCHPART,m.accountno,m.cardprimary desc,m.cardno ", strOrderBy = string.Empty;
    protected string curCurrency = string.Empty;
    protected string strMainTableCond = string.Empty, strSubTableCond = string.Empty;
    protected string strSupHaving = string.Empty;
    protected DataSet DSstatement = null;//;custDS
    protected DataSet DSemails = null;
    protected DataSet DSBirthYandID = null;
    protected DataSet DSOemails = null;
    protected DataSet DSPassportNos = null;
    protected DataSet DStClientBank = null;
    protected DataSet SummaryCor = null;
    protected DataSet DStRefernceCity = null;
    protected DataSet DSreward = null;
    protected DataSet DSInstallment = null;
    protected DataSet DSclientContactData = null;
    protected DataSet DSaccountCurrencies = null;
    protected DataSet DSstatementForReports = null;
    protected DataSet DScontractCurr = null;
    protected DataSet DSProducts = null;
    protected DataSet DSstatementHist = null;
    protected DataSet DSCompany = null;
    protected DataSet DSBranches = null;
    protected string RewardCondStr = "'Reward Program (Airmile)'";
    protected string InstallmentCond = "'BDC Installment Card Program'";
    protected bool isFullFields = true;
    protected DataRowCollection dtlRowColl; //= new DataRowCollection()
    protected static DateTime currStatementdatefrom = DateTime.Now;
    protected static DateTime currStatementdateto = DateTime.Now;
    protected static string curBrnch = string.Empty;
    protected OracleConnection conn;
    protected bool isGetTotalVal = false;
    protected string masterQuery = string.Empty, detailQuery = string.Empty, cardsQuery = string.Empty;
    protected string MainSchema = clsCnfg.readSetting("MainSchema");

    protected void FillStatementDataSet(int pBrach)//DataSet public static 
    {
        //vmBranch = 233;
        curBrnch = pBrach.ToString();
        masterQuery = getMasterQuery;
        detailQuery = getDetailQuery;

        getBasicData();
        detailQuery += " and d.postingdate between to_date('" + currStatementdatefrom.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') and to_date('" + currStatementdateto.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') ";
        if (curProduct.Length > 0)
        {
            masterQuery += " and " + curProduct;
            //masterQuery += " and cardproduct ='" + curProduct + "' ";
        }

        if (curCurrency.Length > 0)
        {
            masterQuery += " and m.cardaccountno like '%" + curCurrency + "' ";
            detailQuery += " and d.accountno like '%" + curCurrency + "' ";
        }

        if (strMainTableCond.Length > 0)
        {
            masterQuery += " and " + strMainTableCond;
        }

        if (strSubTableCond.Length > 0)
        {
            detailQuery += " and " + strSubTableCond;
        }


        //strOrder = " m.accountno,m.cardprimary desc,m.cardno ";//,m.cardstatus
        if (isCorporate)
        {
            masterQuery += " and packagename = 'STTOTABLECORP'";
            detailQuery += " and packagename = 'STTOTABLECORP'";
            //strOrder = " m.accountcurrency,m.contractno,m.customerno,m.accountno,m.accountcurrency,m.cardaccountno,m.cardprimary,m.cardstatus,m.cardno ";
            strOrder = " m.contractno,m.accountno,m.cardaccountno,m.cardprimary,m.cardcreationdate ";
            //strOrder = " m.contractno,m.accountno,m.cardaccountno,m.cardprimary,m.cardstate desc,m.cardno ";
            masterQuery += " order by m.accountcurrency,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }
        else
        {
            masterQuery += " order by " + strOrder;
            //if (curBranch == 1)//
            //  masterQuery += " order by " + strOrder; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 3)//AWB || curBranch == 3
            //  masterQuery += " order by m.CARDBRANCHPART,m.accountno,m.cardprimary,m.cardstatus,m.cardno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 5)//AIB
            //  masterQuery += " order by m.CARDBRANCHPART,m.externalno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            ////masterQuery += " order by m.CARDBRANCHPART,m.cardno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 6)//AAIB
            //  masterQuery += " order by substr(m.cardbranchpartname,6)," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 10)//BNP
            //  masterQuery += " order by m.cardproduct,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            //else  //
            //  masterQuery += " order by " + strOrderBy + "m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }

        /*    if(curBranch == 5 || curBranch == 3)//AIB
              detailQuery += " order by accountno,cardno,postingdate";
            else
              detailQuery += " order by accountno,cardno,postingdate ";
        */
        detailQuery += strSupHaving;
        //detailQuery += " order by d.accountno,d.cardno,d.postingdate";
        //detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno";
        detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno,d.entryno";

        try
        {
            conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter MasterDA = new OracleDataAdapter(masterQuery, conn);
            OracleDataAdapter detailDA = new OracleDataAdapter(detailQuery, conn);

            conn.Open();

            //(new OracleCommand("alter session set optimizer_goal = RULE", conn)).ExecuteNonQuery();
            DSstatement = new DataSet("MasterDetailDS");
            MasterDA.Fill(DSstatement, "tStatementMasterTable");
            detailDA.Fill(DSstatement, "tStatementDetailTable");
            if (isCorporate)
            {
                //string contractQuery = "SELECT DISTINCT x.contractno FROM tstatementmastertable x where x.branch = " + curBranch + " order by x.accountcurrency";// + " and x.contractno = 'CORP_EGP_1069140001'"
                string contractQuery = "SELECT DISTINCT x.accountcurrency,x.contractno FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and packagename = 'STTOTABLECORP' order by x.accountcurrency";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter contractDA = new OracleDataAdapter(contractQuery, conn);
                contractDA.Fill(DSstatement, "contractTable");
                // create currency table
                string currencyQuery = "SELECT DISTINCT x.accountcurrency FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and packagename = 'STTOTABLECORP' and x.accountcurrency is not null";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter currencyDA = new OracleDataAdapter(currencyQuery, conn);
                currencyDA.Fill(DSstatement, "currencyTable");
            }
            if (isCredite)
            {
                //if (!isGetTotalVal)
                //{
                try
                {
                    StatementNoDRel = DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementno], DSstatement.Tables["tStatementDetailTable"].Columns[dStatementno]);
                }
                catch (Exception e)
                {
                    clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + e.Message, System.IO.FileMode.Append);

                    MessageBox.Show(e.Message);
                }
                //}
            }
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        //return DSstatement;
    }

    protected void FillStatementDataSetWithRemovingMarkupFee(int pBrach)//DataSet public static 
    {
        //vmBranch = 233;
        curBrnch = pBrach.ToString();
        masterQuery = getMasterQuery;
        detailQuery = getDetailQueryForMarkupFee;     //getDetailQuery;

        getBasicData();
        detailQuery += " and d.postingdate between to_date('" + currStatementdatefrom.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') and to_date('" + currStatementdateto.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') ";
        if (curProduct.Length > 0)
        {
            masterQuery += " and " + curProduct;
            //masterQuery += " and cardproduct ='" + curProduct + "' ";
        }

        if (curCurrency.Length > 0)
        {
            masterQuery += " and m.cardaccountno like '%" + curCurrency + "' ";
            detailQuery += " and d.accountno like '%" + curCurrency + "' ";
        }

        if (strMainTableCond.Length > 0)
        {
            masterQuery += " and " + strMainTableCond;
        }

        if (strSubTableCond.Length > 0)
        {
            detailQuery += " and " + strSubTableCond;
        }

        //strOrder = " m.accountno,m.cardprimary desc,m.cardno ";//,m.cardstatus
        if (isCorporate)
        {
            masterQuery += " and packagename = 'STTOTABLECORP'";
            detailQuery += " and packagename = 'STTOTABLECORP'";
            //strOrder = " m.accountcurrency,m.contractno,m.customerno,m.accountno,m.accountcurrency,m.cardaccountno,m.cardprimary,m.cardstatus,m.cardno ";
            strOrder = " m.contractno,m.accountno,m.cardaccountno,m.cardprimary,m.cardcreationdate ";
            //strOrder = " m.contractno,m.accountno,m.cardaccountno,m.cardprimary,m.cardstate desc,m.cardno ";
            masterQuery += " order by m.accountcurrency,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }
        else
        {
            masterQuery += " order by " + strOrder;
            //if (curBranch == 1)//
            //  masterQuery += " order by " + strOrder; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 3)//AWB || curBranch == 3
            //  masterQuery += " order by m.CARDBRANCHPART,m.accountno,m.cardprimary,m.cardstatus,m.cardno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 5)//AIB
            //  masterQuery += " order by m.CARDBRANCHPART,m.externalno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            ////masterQuery += " order by m.CARDBRANCHPART,m.cardno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 6)//AAIB
            //  masterQuery += " order by substr(m.cardbranchpartname,6)," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 10)//BNP
            //  masterQuery += " order by m.cardproduct,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            //else  //
            //  masterQuery += " order by " + strOrderBy + "m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }

        /*    if(curBranch == 5 || curBranch == 3)//AIB
              detailQuery += " order by accountno,cardno,postingdate";
            else
              detailQuery += " order by accountno,cardno,postingdate ";
        */
        detailQuery += strSupHaving;
        //detailQuery += " order by d.accountno,d.cardno,d.postingdate";
        //detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno";
        detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno,d.entryno";

        try
        {
            conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter MasterDA = new OracleDataAdapter(masterQuery, conn);
            OracleDataAdapter detailDA = new OracleDataAdapter(detailQuery, conn);

            conn.Open();

            //(new OracleCommand("alter session set optimizer_goal = RULE", conn)).ExecuteNonQuery();
            DSstatement = new DataSet("MasterDetailDS");
            MasterDA.Fill(DSstatement, "tStatementMasterTable");
            detailDA.Fill(DSstatement, "tStatementDetailTable");

            // Remove Mark-up fee 

            for (int i = DSstatement.Tables["tStatementDetailTable"].Rows.Count - 1; i >= 0; i--)
            {
                string strr_trx_desc = DSstatement.Tables["tStatementDetailTable"].Rows[i]["TRANDESCRIPTION"].ToString().Trim();
                if (!strr_trx_desc.Contains("MARK-UP"))
                    continue;
                else
                    DSstatement.Tables["tStatementDetailTable"].Rows[i].Delete();
            }
            DSstatement.Tables["tStatementDetailTable"].AcceptChanges();



            // Remove Mark-up fee



            if (isCorporate)
            {
                //string contractQuery = "SELECT DISTINCT x.contractno FROM tstatementmastertable x where x.branch = " + curBranch + " order by x.accountcurrency";// + " and x.contractno = 'CORP_EGP_1069140001'"
                string contractQuery = "SELECT DISTINCT x.accountcurrency,x.contractno FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and packagename = 'STTOTABLECORP' order by x.accountcurrency";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter contractDA = new OracleDataAdapter(contractQuery, conn);
                contractDA.Fill(DSstatement, "contractTable");
                // create currency table
                string currencyQuery = "SELECT DISTINCT x.accountcurrency FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and packagename = 'STTOTABLECORP' and x.accountcurrency is not null";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter currencyDA = new OracleDataAdapter(currencyQuery, conn);
                currencyDA.Fill(DSstatement, "currencyTable");
            }
            if (isCredite)
            {
                //if (!isGetTotalVal)
                //{
                try
                {
                    StatementNoDRel = DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementno], DSstatement.Tables["tStatementDetailTable"].Columns[dStatementno]);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
                //}
            }
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        //return DSstatement;
    }


    public DataSet getClientEmail(int pBrach)
    {
        string strQuery = "";

        //strQuery = "SELECT /*+ index (" + MainSchema + "TCLIENTPERSONE UK_CLIENTPERSONE) */t.idclient, t.email, t.mobilephone, t.phone , t.externalid, t.latfio,b.valuestr FROM " + MainSchema + "tClientPersone t JOIN " + MainSchema + "TCLIENT a ON (t.branch = a.branch AND t.idclient = a.idclient) LEFT OUTER JOIN " + MainSchema + "TOBJADDPROPDATA_CLIENT b ON (a.branch = b.branch AND a.OBJECTUID = b.OWNERID) WHERE t.branch =" + pBrach;
        strQuery = "SELECT t.idclient, t.email, t.mobilephone, t.phone , t.externalid, t.latfio FROM " + MainSchema + "tClientPersone t WHERE t.branch =" + pBrach;
        //EDT-465 => MSCC-6392
        if (pBrach == 21 || pBrach == 38)
            strQuery = "SELECT t.idclient, t.email, t.mobilephone, t.phone , t.externalid, t.latfio,b.email2 FROM " + MainSchema + "tClientPersone t JOIN " + MainSchema + "TCLIENT a ON (t.branch = a.branch AND t.idclient = a.idclient) LEFT OUTER JOIN " + MainSchema + "TUP$" + pBrach + "$CLIENT$ b ON (a.OBJECTUID = b.up$OWNERID) WHERE t.branch =" + pBrach;

        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter EmailsDA = new OracleDataAdapter(strQuery, conn);

            if ((pBrach == 56) || (pBrach == 74) || pBrach == 127)
            {
                conn.Open();
                DSemails = new DataSet("tClientPersone");
                EmailsDA.Fill(DSemails, "tClientPersone");
                conn.Close();
            }
            else
            {
                conn.Open();
                DSemails = new DataSet("Emails");
                EmailsDA.Fill(DSemails, "Emails");
                conn.Close();
            }
        }
        catch (OracleException ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return DSemails;
    }

    public DataSet getClientWithEmail(int pBrach)
    {
        string strQuery = "";

        strQuery = "SELECT t.idclient, t.email, t.mobilephone, t.phone , t.externalid, t.latfio FROM " + MainSchema + "tClientPersone t WHERE t.branch =" + pBrach + "and t.email is not null";
        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter OemailsDA = new OracleDataAdapter(strQuery, conn);
            if (pBrach == 127)
            {
                conn.Open();
                DSOemails = new DataSet("tClientPersone");
                OemailsDA.Fill(DSOemails, "tClientPersone");
                conn.Close();
            }
            else
            {
                conn.Open();
                DSOemails = new DataSet("Oemails");
                OemailsDA.Fill(DSOemails, "Oemails");
                conn.Close();
            }
        }
        catch (OracleException ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return DSOemails;
    }

    public DataSet getClientPassportNo(int pBrach)
    {
        string strQuery = "";

        //strQuery = "SELECT /*+ index (" + MainSchema + "TCLIENTPERSONE UK_CLIENTPERSONE) */t.idclient, t.email, t.mobilephone, t.phone , t.externalid, t.latfio,b.valuestr FROM " + MainSchema + "tClientPersone t JOIN " + MainSchema + "TCLIENT a ON (t.branch = a.branch AND t.idclient = a.idclient) LEFT OUTER JOIN " + MainSchema + "TOBJADDPROPDATA_CLIENT b ON (a.branch = b.branch AND a.OBJECTUID = b.OWNERID) WHERE t.branch =" + pBrach;
        strQuery = "SELECT t.idclient, t.NO FROM " + MainSchema + "tIdentity t WHERE t.branch =" + pBrach;

        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter PassportNODA = new OracleDataAdapter(strQuery, conn);

            conn.Open();
            DSPassportNos = new DataSet("PassportNos");
            PassportNODA.Fill(DSPassportNos, "PassportNos");
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return DSemails;
    }

    public DataSet getClientPasNoAndBirthYear(int pBrach)
    {
        string strQuery = "";

        //strQuery = "SELECT /*+ index (" + MainSchema + "TCLIENTPERSONE UK_CLIENTPERSONE) */t.idclient, t.email, t.mobilephone, t.phone , t.externalid, t.latfio,b.valuestr FROM " + MainSchema + "tClientPersone t JOIN " + MainSchema + "TCLIENT a ON (t.branch = a.branch AND t.idclient = a.idclient) LEFT OUTER JOIN " + MainSchema + "TOBJADDPROPDATA_CLIENT b ON (a.branch = b.branch AND a.OBJECTUID = b.OWNERID) WHERE t.branch =" + pBrach;
        //strQuery = "SELECT t.idclient, t.email, t.mobilephone, t.phone , t.externalid, t.latfio FROM " + MainSchema + "tClientPersone t WHERE t.branch =" + pBrach;
        strQuery = "SELECT  ti.idclient, replace(replace(ti.NO,'X',''),'x','') NO, decode(to_char(TC.BIRTHDAY,'YYYY'), NULL, '0000',to_char(TC.BIRTHDAY,'YYYY')) BIRTHyear FROM " + MainSchema + "tClientPersone tc, " + MainSchema + "tIdentity ti  WHERE tc.branch = ti.branch and tc.idclient = ti.idclient and length(replace(replace(ti.NO,'X',''),'x','')) > 4 and tc.branch = " + pBrach;

        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter BirthYandIDDA = new OracleDataAdapter(strQuery, conn);

            conn.Open();
            DSBirthYandID = new DataSet("ClientPasNoAndBirthYear");
            BirthYandIDDA.Fill(DSBirthYandID, "ClientPasNoAndBirthYear");
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return DSemails;
    }

    public int getTranCount(int pBrach, string cardno)
    {
        string strQuery = "";
        object counter = 0;
        //strQuery = "SELECT /*+ index (" + MainSchema + "TCLIENTPERSONE UK_CLIENTPERSONE) */t.idclient, t.email, t.mobilephone, t.phone , t.externalid, t.latfio,b.valuestr FROM " + MainSchema + "tClientPersone t JOIN " + MainSchema + "TCLIENT a ON (t.branch = a.branch AND t.idclient = a.idclient) LEFT OUTER JOIN " + MainSchema + "TOBJADDPROPDATA_CLIENT b ON (a.branch = b.branch AND a.OBJECTUID = b.OWNERID) WHERE t.branch =" + pBrach;
        strQuery = "SELECT count(cardno) from " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " d  WHERE d.branch =" + pBrach + " and d.cardno = '" + cardno + "' and docno is not null and postingdate is not null";

        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand countcmd = new OracleCommand(strQuery, conn);

            conn.Open();
            counter = countcmd.ExecuteScalar();
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return Convert.ToInt32(counter);
    }

    public decimal getTranTotal(int pBrach, string cardno)
    {
        string strQuery = "";
        object sum = 0;
        //strQuery = "SELECT /*+ index (" + MainSchema + "TCLIENTPERSONE UK_CLIENTPERSONE) */t.idclient, t.email, t.mobilephone, t.phone , t.externalid, t.latfio,b.valuestr FROM " + MainSchema + "tClientPersone t JOIN " + MainSchema + "TCLIENT a ON (t.branch = a.branch AND t.idclient = a.idclient) LEFT OUTER JOIN " + MainSchema + "TOBJADDPROPDATA_CLIENT b ON (a.branch = b.branch AND a.OBJECTUID = b.OWNERID) WHERE t.branch =" + pBrach;
        //strQuery = "SELECT sum(billtranamount) from " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " d  WHERE d.branch =" + pBrach + " and d.cardno = '" + cardno + "' and docno is not null and postingdate is not null";
        //strQuery = "SELECT cr.branch, db.db,cr.cr,db.db - cr.cr diff " +
        strQuery = "SELECT db.db - cr.cr diff " +
               "FROM (  SELECT SUM (db) db, branch " +
               "FROM(SELECT SUM (billtranamount) db,d.branch " +
               "FROM " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " d " +
               "WHERE     d.branch = " + pBrach + " " +
               "AND d.cardno = '" + cardno + "' " +
               "AND docno IS NOT NULL " +
               "AND postingdate IS NOT NULL " +
               "AND billtranamountsign = 'DB' " +
               "GROUP BY d.branch " +
               "UNION ALL " +
               "SELECT 0 db, d.branch " +
               "FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " d " +
               "WHERE     d.branch = " + pBrach + " AND ROWNUM = 1 " +
               "GROUP BY d.branch)" +
               "GROUP BY branch) db," +
               "(SELECT SUM (cr) cr, branch " +
               "FROM (SELECT SUM (billtranamount) cr, d.branch " +
               "FROM " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " d " +
               "WHERE     d.branch = " + pBrach + " " +
               "AND d.cardno = '" + cardno + "' " +
               "AND docno IS NOT NULL " +
               "AND postingdate IS NOT NULL " +
               "AND billtranamountsign = 'CR' " +
               "GROUP BY d.branch " +
               "UNION ALL " +
               "SELECT 0 cr, d.branch " +
               "FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " d " +
               "WHERE     d.branch = " + pBrach + " AND ROWNUM = 1 " +
               "GROUP BY d.branch)" +
               "GROUP BY branch) cr " +
               "WHERE cr.branch = db.branch";

        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand countcmd = new OracleCommand(strQuery, conn);

            conn.Open();
            sum = countcmd.ExecuteScalar();
            conn.Close();
            if (sum == DBNull.Value)
                sum = 0;
        }
        catch (OracleException ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return Convert.ToDecimal(sum);
    }

    public DataSet getCardProduct(int pBrach)
    {
        string strQuery;
        strQuery = "select cp.code, cp.name from " + MainSchema + "tReferenceCardProduct cp where cp.branch = " + pBrach + " order by 2";

        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter ProductsDA = new OracleDataAdapter(strQuery, conn);

            conn.Open();
            DSProducts = new DataSet("Products");
            ProductsDA.Fill(DSProducts, "Products");
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return DSProducts;
    }

    public DataSet getBranchPart(int pBrach)
    {
        string strQuery;
        strQuery = "select bp.branch, bp.branchpart ,bp.ident from " + MainSchema + "tBranchPart bp where bp.branch = " + pBrach;

        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter BranchesDA = new OracleDataAdapter(strQuery, conn);

            conn.Open();
            DSBranches = new DataSet("Brnaches");
            BranchesDA.Fill(DSBranches, "Brnaches");
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return DSBranches;
    }

    public DataSet getTClientBank(int pBrach)
    {
        string strQuery;
        //strQuery = "select /*+ index (" + clsSessionValues.mainDbSchema + "TCLIENTPERSONE UK_CLIENTPERSONE) */ t.idclient, t.email, t.mobilephone, t.phone from " + clsSessionValues.mainDbSchema + "tClientPersone t where t.branch = " + pBrach;
        strQuery = "select t.idclient, t.emaillegal, t.emailcont, t.phonelegal, t.phonecont from " + MainSchema + "tClientbank t where t.branch = " + pBrach;
        //strQuery = "select t.idclient, t.email from tClientPersone t where t.branch = " + pBrach;

        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter DAtClientBank = new OracleDataAdapter(strQuery, conn);

            conn.Open();
            DStClientBank = new DataSet("tClientBank");
            DAtClientBank.Fill(DStClientBank, "tClientBank");
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return DStClientBank;
    }


    public DataSet GetSummaryPerCorporate(int pBrach)
    {
        string strQuery;
        strQuery = $"select SUM(S.CLOSINGBALANCE) as CLOSINGBALANCE ,SUM(S.CARDLIMIT) as CARDLIMIT,SUM(S.TOTALOVERDUEAMOUNT)  as TOTALOVERDUEAMOUNT, s.clientid from {clsSessionValues.mainDbSchema + clsSessionValues.mainTable} s where branch = {pBrach} GROUP BY S.CLOSINGBALANCE,S.CARDLIMIT,S.TOTALOVERDUEAMOUNT, s.clientid";
        //strQuery = "select t.idclient, t.email from tClientPersone t where t.branch = " + pBrach;

        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter EmailsDA = new OracleDataAdapter(strQuery, conn);

            conn.Open();
            SummaryCor = new DataSet("SummaryCor");
            EmailsDA.Fill(SummaryCor, "SummaryCor");
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return SummaryCor;
    }

    public DataSet getClientEmailName(int pBrach)
    {
        string strQuery;
        strQuery = "select t.idclient, t.email, t.mobilephone, t.fio from " + MainSchema + "tClientPersone t where t.branch = " + pBrach;
        //strQuery = "select t.idclient, t.email from tClientPersone t where t.branch = " + pBrach;

        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter EmailsDA = new OracleDataAdapter(strQuery, conn);

            conn.Open();
            DSemails = new DataSet("Emails");
            EmailsDA.Fill(DSemails, "Emails");
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return DSemails;
    }


    //public void addClientPersone(int pBrach)
    //    {
    //    //string rewardSql = "select * from V_StatementMasterReward r where r.branch = " + pBrach;
    //    string clientSql = "select * from " + MainSchema + "tClientPersone c where c.branch = " + pBrach;
    //    OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
    //    OracleDataAdapter rewardDA = new OracleDataAdapter(clientSql, conn);
    //    conn.Open();
    //    rewardDA.Fill(DSstatement, "tClientPersone");
    //    }


    public DataSet getClientContactData(int pBrach)
    {
        string strQuery;
        strQuery = "select c.idclient,m.customername,m.customeraddress1, m.customeraddress2,m.customeraddress3, trim(m.customercity || ' ' || m.customerregion || ' ' || m.customercountry) as custArea, c.phone, c.mobilephone from " + MainSchema + "tClientPersone c, " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " m where m.clientid = c.idclient AND m.CONTRACTTYPE <> 'Reward Program (Airmile)' and  c.branch = " + pBrach;
        //strQuery = "select t.idclient, t.email from " + clsSessionValues.mainDbSchema + "tClientPersone t where t.branch = " + pBrach;

        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter ClientContactDataDA = new OracleDataAdapter(strQuery, conn);

            conn.Open();
            DSclientContactData = new DataSet("ClientContactData");
            ClientContactDataDA.Fill(DSclientContactData, "ClientContactData");
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return DSclientContactData;
    }

    public DataSet fillAccountCurrencies(int pBrach)
    {
        string strQuery;
        strQuery = "select distinct c." + mAccountcurrency + " from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " c where c.branch = " + pBrach;
        //strQuery = "select t.idclient, t.email from " + clsSessionValues.mainDbSchema + "tClientPersone t where t.branch = " + pBrach;

        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter AccountCurrenciesDA = new OracleDataAdapter(strQuery, conn);

            conn.Open();
            DSaccountCurrencies = new DataSet("AccountCurrencies");
            AccountCurrenciesDA.Fill(DSaccountCurrencies, "AccountCurrencies");
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return DSclientContactData;
    }

    public DataSet fillContractCurr(int pBrach)
    {
        string strQuery;
        strQuery = "select distinct c." + mAccountcurrency + " ,c." + mCardproduct + ",c." + mContractno + " from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " c where c.branch = " + pBrach + " and c." + mAccountcurrency + " is not null " + " order by c." + mAccountcurrency + " ,c." + mCardproduct + " ,c." + mContractno;//+ " ,c." + mAccountno + " ,c." + mCardno
        //strQuery = "select distinct c." + mAccountcurrency + " ,c." + mCardproduct + ",c." + mContractno + " from tstatementmastertable c where c.branch = " + pBrach + " and c." + mAccountcurrency + " is not null " + " and contractno = 'CORP_EGP_0001153001' " + " order by c." + mAccountcurrency + " ,c." + mCardproduct + " ,c." + mContractno;//+ " ,c." + mAccountno + " ,c." + mCardno
        //strQuery = "select distinct c." + mCardno + ",c." + mAccountno + ",c." + mContractno + ",c." + mAccountcurrency + " from tstatementmastertable c where c.branch = " + pBrach + " and c." + mAccountcurrency + " is not null " + " and contractno = 'CORP_EGP_0094417001' " + " order by c." + mAccountcurrency + " ,c." + mContractno + " ,c." + mAccountno + " ,c." + mCardno;
        //strQuery = "select distinct c." + mCardno + ",c." + mAccountno + ",c." + mContractno + ",c." + mAccountcurrency + " from tstatementmastertable c where c.branch = " + pBrach + " and c." + mAccountcurrency + " is not null " + " order by c." + mAccountcurrency + " ,c." + mContractno + " ,c." + mAccountno + " ,c." + mCardno;

        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter ContractCurrDA = new OracleDataAdapter(strQuery, conn);

            conn.Open();
            DScontractCurr = new DataSet("ContractCurr");
            ContractCurrDA.Fill(DScontractCurr, "ContractCurr");
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return DSclientContactData;
    }

    public DataSet getReward(int pBrach)
    {
        string strQuery;
        //strQuery = "select /*+ index (" + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + "  MASTER" + clsSessionValues.mainTable.Substring(21, 9) + "_idx_2) */ r.CLIENTID, r.OPENINGBALANCE, r.EARNEDBONUS, r.REDEEMEDBONUS , r.CLOSINGBALANCE, r.BONUSADJUSTMENT from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " r where r.contracttype in " + RewardCondStr + " and r.branch = " + pBrach; //  , r.bonusadjustment
        //if (pBrach == 36)
        strQuery = "select r.CLIENTID, r.OPENINGBALANCE, r.EARNEDBONUS, r.REDEEMEDBONUS , r.EXPIREDBONUS, r.CLOSINGBALANCE, r.BONUSADJUSTMENT , r.EXPIREDBONUSNEXTMONTH , r.EXPIREDBONUSDATE ,r.CREDITCONTRACTS from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " r where r.branch = " + pBrach + " and r.contracttype in " + RewardCondStr + ""; //  , r.bonusadjustment
        //else
        //    strQuery = "select /*+ index (" + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + "  MASTER" + clsSessionValues.mainTable.Substring(21, 9) + "_idx_2) */ r.CLIENTID, r.OPENINGBALANCE, r.EARNEDBONUS, r.REDEEMEDBONUS , r.CLOSINGBALANCE, r.BONUSADJUSTMENT from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " r where r.contracttype in " + RewardCondStr + " and r.branch = " + pBrach; //  , r.bonusadjustment
        //strQuery = "select t.idclient, t.email from tClientPersone t where t.branch = " + pBrach;

        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter rewardDA = new OracleDataAdapter(strQuery, conn);

            conn.Open();
            DSreward = new DataSet("Reward");
            rewardDA.Fill(DSreward, "Reward");
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return DSreward;
    }

    public DataSet getInstallment(int pBrach)
    {
        string strQuery;
        strQuery = "select i.CLIENTID, i.OPENINGBALANCE, i.CLOSINGBALANCE,i.CREDITCONTRACTS from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " i where i.contracttype in " + InstallmentCond + " and i.branch = " + pBrach; //  , r.bonusadjustment

        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter InstallmentDA = new OracleDataAdapter(strQuery, conn);

            conn.Open();
            DSInstallment = new DataSet("Installment");
            InstallmentDA.Fill(DSInstallment, "Installment");
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return DSInstallment;
    }

    public void getBasicData()
    {
        //string pStrQuery = "select * from tstatementmastertable m2 where m2.branch = " + pBrach + " and rownum = 1";
        string pStrQuery = "select m.Statementdatefrom ,m.Statementdateto from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " m where m.BRANCH = " + curBrnch + " and rownum = 1";
        OracleDataReader rtnDataReader = null;
        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand cmnd = new OracleCommand(pStrQuery, conn);
            conn.Open();
            rtnDataReader = cmnd.ExecuteReader(CommandBehavior.CloseConnection);
            if (rtnDataReader.HasRows)
            {
                rtnDataReader.Read();
                currStatementdatefrom = Convert.ToDateTime(rtnDataReader.GetValue(0));
                currStatementdateto = Convert.ToDateTime(rtnDataReader.GetValue(1));
                TimeSpan d = frmStatementFile.StDate - currStatementdateto;
                if (clsSessionValues.Prompt4Msg && (currStatementdateto > DateTime.Now || (DateTime.Now - currStatementdateto) > new TimeSpan(3, 0, 0, 0, 0)))
                {
                    System.Windows.Forms.MessageBox.Show("Error in date in bank " + curBrnch + "\n dateto: " + currStatementdateto.ToString() + "\n today: " + DateTime.Now.ToString(), "Text");
                }
            }
        }
        catch (OracleException ex)
        {
            clsBasErrors.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
    }//getDataReader

    public bool getBasicData(int bankCode, DateTime stdate, string bankName, string run)
    {

        curBrnch = Convert.ToString(bankCode);
        string DateTo = stdate.ToString("MM/dd/yyyy");
        //string pStrQuery = "select * from tstatementmastertable m2 where m2.branch = " + pBrach + " and rownum = 1";
        string pStrQuery = "select m.Statementdatefrom ,m.Statementdateto from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " m where m.BRANCH = " + curBrnch + " and rownum = 1 and Statementdateto=to_date('" + DateTo + "','mm/dd/yyyy')";
        OracleDataReader rtnDataReader = null;
        try
        {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand cmnd = new OracleCommand(pStrQuery, conn);
            conn.Open();
            rtnDataReader = cmnd.ExecuteReader(CommandBehavior.CloseConnection);
            if (rtnDataReader.HasRows)
            {
                rtnDataReader.Read();
                currStatementdatefrom = Convert.ToDateTime(rtnDataReader.GetValue(0));
                currStatementdateto = Convert.ToDateTime(rtnDataReader.GetValue(1));
                return true;
            }
            else
            {
                clsEmail sndMail = new clsEmail();
                sndMail.sendEmail(bankCode, stdate, bankName, currStatementdatefrom, currStatementdateto, run);
                return false;

            }

        }
        catch (OracleException ex)
        {
            clsBasErrors.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return true;
    }//getDataReader

    //protected void assignMasterFiledVal()
    //    {
    //    }

    //protected void assignDetailFieldVal()
    //    {
    //    }

    protected string curRewardCond //static public 
    {
        get { return RewardCondStr; }
        set { RewardCondStr = value; }
    }// curRewardCond

    protected string curInstallmentCond //static public 
    {
        get { return InstallmentCond; }
        set { InstallmentCond = value; }
    }// InstallmentCondStr

    protected int curBranchVal //static public 
    {
        get { return curBranch; }
        set { curBranch = value; }
    }// curBranchVal

    protected string mainTableCond
    {
        get { return curProduct; }
        set { curProduct = value; }
    }// mainTableCond

    protected string MainTableCond
    {
        get { return strMainTableCond; }
        set { strMainTableCond = value; }
    }// MainTableCond

    protected string supTableCond
    {
        get { return strSubTableCond; }
        set { strSubTableCond = value; }
    }// supTableCond

    protected string supHaving
    {
        get { return strSupHaving; }
        set { strSupHaving = value; }
    }// supHaving

    protected string curCurrencyVal
    {
        get { return curCurrency; }
        set { curCurrency = value; }
    }// curCurrencyVal


    protected bool isCorporateVal
    {
        get { return isCorporate; }
        set { isCorporate = value; }
    }// isCorporateVal

    protected bool isCreditVal
    {
        get { return isCredite; }
        set { isCredite = value; }
    }// isCreditVal

    protected string orderBy
    {
        get { return strOrderBy; }
        set { strOrderBy = value; }
    }// orderBy

    protected bool isGetTotal //check for if the ststement genrate total for Transactions
    {
        set { isGetTotalVal = value; }
    }

    protected string getMasterQuery
    {
        get
        {//MasterQuery 
            string mQry = string.Empty;
            //if (isFullFields)
            //if (curBranch != 24)
            //{
            mQry = "select m.branch,m.statementno,m.statementnumber,m.statementdatefrom,m.statementdateto,m.statementtype,m.statementsendtype,m.stetementduedate,m.statementmessageline1,m.statementmessageline2,m.statementmessageline3,m.contractno,m.contracttype,m.contractcreationdate,m.contractstatus,m.companycode,m.registrationnumber,m.contractlimit,m.customerno,m.customertitle,m.customername,m.customercreationdate,m.customertype,m.customerstatus,m.clientid,m.employer,m.dept,m.position,m.employeenumber,m.customercountry,m.customerregion,m.customercity,m.customerzipcode,m.customeraddress1,m.customeraddress2,m.customeraddress3,m.customeraddressparcode,m.accountno,m.externalno,m.accountcreationdate,m.accounttype,m.accountstatus,m.accountcurrency,m.accountcountry,m.accountregion,m.accountcity,m.accountzipcode,m.accountaddress1,m.accountaddress2,m.accountaddress3,m.accountaddressparcode,m.accountlim,m.accountavailablelim,m.accountcashlim,m.accountavailablecashlim,m.cardno,m.mbr,m.cardtype,m.cardcreationdate,m.cardexpirydate,m.cardlastmodificationdate,m.cardproduct,m.cardstate,m.cardstatus,m.cardstatusdate,m.cardcurrency,m.cardbirthdate,m.cardtitle,m.cardvip,m.cardprimary,m.prinarycardno,m.cardbranchpart,m.cardbranchpartname,m.cardaccountno,m.cardclientname,m.cardpaymentmethod,m.cardlimit,m.cardavailablelimit,m.cardcashlimit,m.cardavailablecashlimit,m.cardcountry,m.cardregion,m.cardcity,m.cardzipcode,m.cardaddress1,m.cardaddress2,m.cardaddress3,m.cardaddressbarcode,m.totaldueamount,m.mindueamount,m.openingbalance,m.closingbalance,m.totalpurchases,m.totalcashwithdrawal,m.totalinterest,m.totalpayments,m.totalcharges,m.totalcreditsandrefunds,m.totaldebits,m.totalcredits,m.totalsuspends,m.totaloverdueamount,m.totaloverlimitamount,m.carddafamount,m.dafpercentage"
            + ",m.holstmt,m.barcode,m.useractfield1,m.useractfield2,m.usercustfield1,m.usercustfield2,m.usercardfield1,m.usercardfield2,m.contactpersonename,m.cardclientid,m.totaldueamount,m.intspentlimit,m.minpaypercentage,m.monthscount,m.installmentusedlimit ,m.OVERDUEDAYS, m.CONTRACTSTATE";


            //}
            //    else
            //        {
            //        mQry = "select m.branch,m.statementno,m.statementnumber,m.statementdatefrom,m.statementdateto,m.statementtype,m.statementsendtype,m.stetementduedate,m.statementmessageline1,m.statementmessageline2,m.statementmessageline3,m.contractno,m.contracttype,m.contractstatus,m.companycode,m.registrationnumber,m.contractlimit,m.customerno,m.customertitle,m.customername,m.customertype,m.customerstatus,m.clientid,m.employer,m.dept,m.position,m.employeenumber,m.customercountry,m.customerregion,m.customercity,m.customerzipcode,m.customeraddress1,m.customeraddress2,m.customeraddress3,m.customeraddressparcode,m.accountno,m.externalno,m.accounttype,m.accountstatus,m.accountcurrency,m.accountcountry,m.accountregion,m.accountcity,m.accountzipcode,m.accountaddress1,m.accountaddress2,m.accountaddress3,m.accountaddressparcode,m.accountlim,m.accountavailablelim,m.accountcashlim,m.accountavailablecashlim,m.cardno,m.mbr,m.cardtype,m.cardproduct,m.cardstate,m.cardstatus,m.cardstatusdate,m.cardcurrency,m.cardbirthdate,m.cardtitle,m.cardvip,m.cardprimary,m.prinarycardno,m.cardbranchpart,m.cardbranchpartname,m.cardaccountno,m.cardclientname,m.cardpaymentmethod,m.cardlimit,m.cardavailablelimit,m.cardcashlimit,m.cardavailablecashlimit,m.cardcountry,m.cardregion,m.cardcity,m.cardzipcode,m.cardaddress1,m.cardaddress2,m.cardaddress3,m.cardaddressbarcode,m.totaldueamount,m.mindueamount,m.openingbalance,m.closingbalance,m.totalpurchases,m.totalcashwithdrawal,m.totalinterest,m.totalpayments,m.totalcharges,m.totalcreditsandrefunds,m.totaldebits,m.totalcredits,m.totalsuspends,m.totaloverdueamount,m.totaloverlimitamount,m.carddafamount,m.dafpercentage"
            //        + ",m.holstmt,m.barcode,m.useractfield1,m.useractfield2,m.usercustfield1,m.usercustfield2,m.usercardfield1,m.usercardfield2,m.contactpersonename,m.cardclientid,m.totaldueamount ";//
            //        }
            //else
            //    mQry = "select m." + mBranch + ",m." + mStatementno + ",m." + mCardno + ",m." + mMBR + ",m." + mCardproduct + ",m." + mCardprimary + ",m." + mPrinarycardno + ",m." + mAccountno + ",m." + mStatementdatefrom + ",m." + mStatementdateto + ",m." + mStetementduedate + ",m." + mCardstate + ",m." + mCardstatus + ",m." + mExternalno + ",m." + mCustomername + ",m." + mCardbranchpartname + ",m." + mCustomeraddress1 + ",m." + mCustomeraddress2 + ",m." + mCustomeraddress3 + ",m." + mCardaddress1 + ",m." + mCardaddress2 + ",m." + mCardaddress3 + ",m." + mCustomerregion + ",m." + mCustomercity + ",m." + mAccountlim + ",m." + mAccountavailablelim + ",m." + mCardlimit + ",m." + mTotaloverdueamount + ",m." + mTotaldebits + ",m." + mTotalcredits + ",m." + mTotalpayments + ",m." + mTotalpurchases + ",m." + mTotalcashwithdrawal + ",m." + mTotalcharges + ",m." + mTotalinterest + ",m." + mMindueamount + ",m." + mOpeningbalance + ",m." + mClosingbalance + ",m." + mAccountcurrency + ",m." + mAccounttype + ",m." + mCardbranchpart + ",m." + mStatementmessageline1 + ",m." + mStatementmessageline2 + ",m." + mStatementmessageline3 + ",m." + mContracttype + ",m." + mClientid + ",m." + mAccountzipcode + ",m." + mCardaddressbarcode + ",m." + mCardpaymentmethod + ",m." + mCardexpirydate + ",m." + mCustomercountry + ",m." + mCardavailablelimit + ",m." + mCardaccountno + ",m." + mContractno + ",m." + mAccountstatus + ",m." + mCardclientname + ",m." + mContractlimit + ",m." + mCustomertitle + ",m." + mCardtitle + ",m." + mDept + ",m." + mContactpersonename + ",m." + mCardvip + ",m." + mCarddafamount + ",m." + mCardClientId + ",m." + mHOLSTMT + ",m." + mBarcode + ",m." + mUserActField1 + ",m." + mDAFPercentage + ",m." + mTotalDueAmount + ",m." + mIntSpentLimit + "m." + mMinPayPercentage + "m." + mMonthsCount;

            //return mQry + " from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " m "
            //  + " where m.BRANCH = " + curBranch //,'011045350007301'

            if (frmStatementFile.Internal == true && !string.IsNullOrEmpty(frmStatementFile.internalAccNo))
            {
                if (clsBasUserData.sType == "CP")
                {
                    return mQry + " from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " m "
             + " where m.BRANCH = " + curBranch //,'011045350007301'
                                                //+ "and m.Statementdatefrom = to_date('" + currStatementdatefrom.ToString("dd/MM/yyyy") + "','dd/mm/yyyy')"
                                                //+ "and m.Statementdateto = to_date('" + currStatementdateto.ToString("dd/MM/yyyy") + "','dd/mm/yyyy')"
             + "and m.CARDACCOUNTNO ='" + frmStatementFile.internalAccNo + "'";
                }
                else
                {
                    return mQry + " from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " m "
             + " where m.BRANCH = " + curBranch //,'011045350007301'
                                                //+ "and m.Statementdatefrom = to_date('" + currStatementdatefrom.ToString("dd/MM/yyyy") + "','dd/mm/yyyy')"
                                                //+ "and m.Statementdateto = to_date('" + currStatementdateto.ToString("dd/MM/yyyy") + "','dd/mm/yyyy')"
             + "and m.accountno ='" + frmStatementFile.internalAccNo + "'";
                }

            }
            else
            {
                mQry += " from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " m "
                 + " where m.BRANCH = " + curBranch;
                // + " and ClientId like '6232287'";

                //if(curBranch == 29)
                //{
                //    mQry += "  ORDER BY " + "m.cardcreationdate DESC";
                //}

                return mQry;
            }
        }
    }// getMasterQuery

    protected string getMasterQueryWithOverDueDays
    {
        get
        {//MasterQuery 
            string mQry = string.Empty;
            //if (isFullFields)
            //if (curBranch != 24)
            //{
            mQry = "select m.branch,m.statementno,m.statementnumber,m.statementdatefrom,m.statementdateto,m.statementtype,m.statementsendtype,m.stetementduedate,m.statementmessageline1,m.statementmessageline2,m.statementmessageline3,m.contractno,m.contracttype,m.contractcreationdate,m.contractstatus,m.companycode,m.registrationnumber,m.contractlimit,m.customerno,m.customertitle,m.customername,m.customercreationdate,m.customertype,m.customerstatus,m.clientid,m.employer,m.dept,m.position,m.employeenumber,m.customercountry,m.customerregion,m.customercity,m.customerzipcode,m.customeraddress1,m.customeraddress2,m.customeraddress3,m.customeraddressparcode,m.accountno,m.externalno,m.accountcreationdate,m.accounttype,m.accountstatus,m.accountcurrency,m.accountcountry,m.accountregion,m.accountcity,m.accountzipcode,m.accountaddress1,m.accountaddress2,m.accountaddress3,m.accountaddressparcode,m.accountlim,m.accountavailablelim,m.accountcashlim,m.accountavailablecashlim,m.cardno,m.mbr,m.cardtype,m.cardcreationdate,m.cardexpirydate,m.cardlastmodificationdate,m.cardproduct,m.cardstate,m.cardstatus,m.cardstatusdate,m.cardcurrency,m.cardbirthdate,m.cardtitle,m.cardvip,m.cardprimary,m.prinarycardno,m.cardbranchpart,m.cardbranchpartname,m.cardaccountno,m.cardclientname,m.cardpaymentmethod,m.cardlimit,m.cardavailablelimit,m.cardcashlimit,m.cardavailablecashlimit,m.cardcountry,m.cardregion,m.cardcity,m.cardzipcode,m.cardaddress1,m.cardaddress2,m.cardaddress3,m.cardaddressbarcode,m.totaldueamount,m.mindueamount,m.openingbalance,m.closingbalance,m.totalpurchases,m.totalcashwithdrawal,m.totalinterest,m.totalpayments,m.totalcharges,m.totalcreditsandrefunds,m.totaldebits,m.totalcredits,m.totalsuspends,m.totaloverdueamount,m.totaloverlimitamount,m.carddafamount,m.dafpercentage"
            + ",m.holstmt,m.barcode,m.useractfield1,m.useractfield2,m.usercustfield1,m.usercustfield2,m.usercardfield1,m.usercardfield2,m.contactpersonename,m.cardclientid,m.totaldueamount,m.intspentlimit,m.minpaypercentage,m.monthscount,m.installmentusedlimit "
            + " , NVL((Select  eod.ODDAYS from A4M.ZM_EOD_CONT_ACCT eod where eod.BRANCH = m.BRANCH and eod.CONTRACTNO = m.CONTRACTNO  and eod.ACCOUNTNO = m.ACCOUNTNO and eod.OPDATE = m.statementdateto ), 0) as OverDueDays ";


            //}
            //    else
            //        {
            //        mQry = "select m.branch,m.statementno,m.statementnumber,m.statementdatefrom,m.statementdateto,m.statementtype,m.statementsendtype,m.stetementduedate,m.statementmessageline1,m.statementmessageline2,m.statementmessageline3,m.contractno,m.contracttype,m.contractstatus,m.companycode,m.registrationnumber,m.contractlimit,m.customerno,m.customertitle,m.customername,m.customertype,m.customerstatus,m.clientid,m.employer,m.dept,m.position,m.employeenumber,m.customercountry,m.customerregion,m.customercity,m.customerzipcode,m.customeraddress1,m.customeraddress2,m.customeraddress3,m.customeraddressparcode,m.accountno,m.externalno,m.accounttype,m.accountstatus,m.accountcurrency,m.accountcountry,m.accountregion,m.accountcity,m.accountzipcode,m.accountaddress1,m.accountaddress2,m.accountaddress3,m.accountaddressparcode,m.accountlim,m.accountavailablelim,m.accountcashlim,m.accountavailablecashlim,m.cardno,m.mbr,m.cardtype,m.cardproduct,m.cardstate,m.cardstatus,m.cardstatusdate,m.cardcurrency,m.cardbirthdate,m.cardtitle,m.cardvip,m.cardprimary,m.prinarycardno,m.cardbranchpart,m.cardbranchpartname,m.cardaccountno,m.cardclientname,m.cardpaymentmethod,m.cardlimit,m.cardavailablelimit,m.cardcashlimit,m.cardavailablecashlimit,m.cardcountry,m.cardregion,m.cardcity,m.cardzipcode,m.cardaddress1,m.cardaddress2,m.cardaddress3,m.cardaddressbarcode,m.totaldueamount,m.mindueamount,m.openingbalance,m.closingbalance,m.totalpurchases,m.totalcashwithdrawal,m.totalinterest,m.totalpayments,m.totalcharges,m.totalcreditsandrefunds,m.totaldebits,m.totalcredits,m.totalsuspends,m.totaloverdueamount,m.totaloverlimitamount,m.carddafamount,m.dafpercentage"
            //        + ",m.holstmt,m.barcode,m.useractfield1,m.useractfield2,m.usercustfield1,m.usercustfield2,m.usercardfield1,m.usercardfield2,m.contactpersonename,m.cardclientid,m.totaldueamount ";//
            //        }
            //else
            //    mQry = "select m." + mBranch + ",m." + mStatementno + ",m." + mCardno + ",m." + mMBR + ",m." + mCardproduct + ",m." + mCardprimary + ",m." + mPrinarycardno + ",m." + mAccountno + ",m." + mStatementdatefrom + ",m." + mStatementdateto + ",m." + mStetementduedate + ",m." + mCardstate + ",m." + mCardstatus + ",m." + mExternalno + ",m." + mCustomername + ",m." + mCardbranchpartname + ",m." + mCustomeraddress1 + ",m." + mCustomeraddress2 + ",m." + mCustomeraddress3 + ",m." + mCardaddress1 + ",m." + mCardaddress2 + ",m." + mCardaddress3 + ",m." + mCustomerregion + ",m." + mCustomercity + ",m." + mAccountlim + ",m." + mAccountavailablelim + ",m." + mCardlimit + ",m." + mTotaloverdueamount + ",m." + mTotaldebits + ",m." + mTotalcredits + ",m." + mTotalpayments + ",m." + mTotalpurchases + ",m." + mTotalcashwithdrawal + ",m." + mTotalcharges + ",m." + mTotalinterest + ",m." + mMindueamount + ",m." + mOpeningbalance + ",m." + mClosingbalance + ",m." + mAccountcurrency + ",m." + mAccounttype + ",m." + mCardbranchpart + ",m." + mStatementmessageline1 + ",m." + mStatementmessageline2 + ",m." + mStatementmessageline3 + ",m." + mContracttype + ",m." + mClientid + ",m." + mAccountzipcode + ",m." + mCardaddressbarcode + ",m." + mCardpaymentmethod + ",m." + mCardexpirydate + ",m." + mCustomercountry + ",m." + mCardavailablelimit + ",m." + mCardaccountno + ",m." + mContractno + ",m." + mAccountstatus + ",m." + mCardclientname + ",m." + mContractlimit + ",m." + mCustomertitle + ",m." + mCardtitle + ",m." + mDept + ",m." + mContactpersonename + ",m." + mCardvip + ",m." + mCarddafamount + ",m." + mCardClientId + ",m." + mHOLSTMT + ",m." + mBarcode + ",m." + mUserActField1 + ",m." + mDAFPercentage + ",m." + mTotalDueAmount + ",m." + mIntSpentLimit + "m." + mMinPayPercentage + "m." + mMonthsCount;

            //return mQry + " from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " m "
            //  + " where m.BRANCH = " + curBranch //,'011045350007301'

            if (frmStatementFile.Internal == true && !string.IsNullOrEmpty(frmStatementFile.internalAccNo))
            {
                if (clsBasUserData.sType == "CP")
                {
                    return mQry + " from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " m "
             + " where m.BRANCH = " + curBranch //,'011045350007301'
                                                //+ "and m.Statementdatefrom = to_date('" + currStatementdatefrom.ToString("dd/MM/yyyy") + "','dd/mm/yyyy')"
                                                //+ "and m.Statementdateto = to_date('" + currStatementdateto.ToString("dd/MM/yyyy") + "','dd/mm/yyyy')"
             + "and m.CARDACCOUNTNO ='" + frmStatementFile.internalAccNo + "'";
                }
                else
                {
                    return mQry + " from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " m "
             + " where m.BRANCH = " + curBranch //,'011045350007301'
                                                //+ "and m.Statementdatefrom = to_date('" + currStatementdatefrom.ToString("dd/MM/yyyy") + "','dd/mm/yyyy')"
                                                //+ "and m.Statementdateto = to_date('" + currStatementdateto.ToString("dd/MM/yyyy") + "','dd/mm/yyyy')"
             + "and m.accountno ='" + frmStatementFile.internalAccNo + "'";
                }

            }
            else
            {
                return mQry + " from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " m "
                  + " where m.BRANCH = " + curBranch //,'011045350007301'
                                                     //+ "and m.Statementdatefrom = to_date('" + currStatementdatefrom.ToString("dd/MM/yyyy") + "','dd/mm/yyyy')"
                                                     //+ "and m.Statementdateto = to_date('" + currStatementdateto.ToString("dd/MM/yyyy") + "','dd/mm/yyyy')"
                                                     //+ "and m.cardproduct = 'Visa Classic'"
                                                     //+ "and m.cardproduct = 'MasterCard Classic Credit'"
                                                     //+ "and m.cardproduct = 'MasterCard Gold Credit'"
                                                     //+ "and m.cardproduct = 'Visa Business - Individuals'"
                                                     //+ "and m.cardproduct = 'MasterCard Platinum Credit'"
                                                     //+ "and m.cardproduct in('MasterCard Classic Credit','MasterCard Gold Credit','MasterCard Platinum Credit')"
                                                     //+ "and m.CONTRACTTYPE IN ('MasterCard Platinum Credit')"
                                                     //+ "and m.CONTRACTNO NOT IN ('V-CR-GOLD_0000269','V-CR-GOLD_0000398','V-CR-GOLD_0000400','V-CR-GOLD_0000404','V-CR-GOLD_0000410')"
                                                     //+ "and m.CONTRACTNO IN ('MC_Inst_000000001126','52003100000000013040','INST_CARD_0000001123')"
                                                     //+ "and m.CONTRACTNO IN ('Gold_Stand_46575','Rew3_GLD_3822693')" //NSGB
                                                     //+ "and m.CONTRACTNO IN ('Plat_Staff_000000051','Rew3_PLT_394022')" //NSGB
                                                     //+ "and m.CONTRACTNO IN ('40284200000000075364')" 
                                                     //  +"and m.accountno IN ('CR_NGN_0000082176', 'Rew_68925' )"
                                                     //+ " and accountno = 'Cre_GHS_0000000056'"
                                                     //+ "and m.CONTRACTNO IN ('Plat_Stand_000002811','Rew3_PLT_392407')" //NSGB
                                                     //+ " and m.accountno in('CR_NGN_0000015960')" //DBN
                                                     //+ " and m.accountno in('Credit_AON_4024','Credit_AON_768')" //BAI
                                                     // + " and m.accountno in('Credit_586001','Credit_589822','Credit_551066')"
                                                     // + " and m.accountno in('Credit_586001','Credit_589822','Credit_551066','Credit_512357','Credit_589729')"
                                                     // + " and m.accountno in('Credit_AON_20')"
                                                     // + " and m.accountno in('Credit_AON_2676')"
                                                     // + " and m.accountno in('CR_NGN_0000011840')"
                                                     //+ " and m.accountno in('4042390001394')"
                                                     //+ " and m.cardclientname = 'KEN SPANN'"
                                                     //+ " and m.externalno = '0003159731'"
                                                     //+ " and m.accountno = 'Prepaid_0000000017'" 
                                                     //+ " and m.accountno = 'MC_CR_KES_0000000173'"
                                                     //+ " and m.accountno = 'MC_CR_KES_0000000158'"
                                                     //+ " and m.accountno = '00002185001'"
                                                     //+ " and m.accountno = 'Prepaid_0000175855'"                
                                                     //+ " and m.externalno IN ('0110081710120501x', '0121175710120601', '1703300','0110253610120501')"
                                                     //+ " and externalno = '0110253610120501'"
                                                     //+ " and cardaccountno = '99999-326-818-B'"
                                                     //+ "and statementno in (104818,76304)"
                                                     //+ " and m.accountno = 'MC_CR_EGP_0000012010'"
                                                     //+ " and cardbranchpart = 1"
                                                     //+ "and cardaccountno = '4174390000042439'"
                                                     //+ "and clientid = 7178504"
                ;
            }
        }
    }// getMasterQuery


    protected string getDetailQuery
    {
        get
        {//detailQuery 
            string dQry = string.Empty;
            //if (isGetTotalVal)
            //    {
            //    dQry = "select d." + dStatementno + ",d." + dCardno + ",d." + dAccountno + ",d." + dDocno + ",d." + dTransdate + ",d." + dPostingdate + ",d." + dRefereneno + ",d." + dAccountcurrency + ",d." + dOrigtranamount + ",d." + dOrigtrancurrency + ",d." + dBilltranamount + ",d." + dBilltranamountsign + ",d." + dTrandescription + ",d." + dMerchant + ",m." + mCardprimary;// +",d." + dInstallmentData + ",d." + dOrigDocNo + ",d. dEntryNo +";
            //    dQry += " from " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " d ," + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " m "
            //  + " where d." + dStatementno + " = m." + mStatementno + " and d.BRANCH = " + curBranch + " and m.BRANCH = " + curBranch //,'011045350007301'
            //        //+ "and statementno in (104818,76304)"
            //  ;
            //    }
            //else
            //    {
            //    if (isFullFields)
            if (curBranch == 69)
            {
                dQry = "select d.branch,d.statementno,d.statementnumber,d.contractno,d.accountno,d.accountcurrency,d.cardno,d.mbr,d.transdate,d.postingdate,d.trandescription,d.merchant,d.origtranamount,d.origtrancurrency,d.billtranamount,d.billtranamountsign,d.docno,d.refereneno,d.installmentdata,d.origdocno,d.installmentmerchant,d.installmentmerchantlocation,d.installmentstartdate,d.installmentloan,d.installmentcycleno,d.installmentcycles,d.installmentregrepayment,d.installmentenddate,d.installmentoutbalance,d.entryno,d.approvalcode";//,d.approvalcode";
            }
            else
            {
                dQry = "select d.branch,d.statementno,d.statementnumber,d.contractno,d.accountno,d.accountcurrency,d.cardno,d.mbr,d.transdate,d.postingdate,d.trandescription,d.merchant,d.origtranamount,d.origtrancurrency,d.billtranamount,d.billtranamountsign,d.docno,d.refereneno,d.installmentdata,d.origdocno,d.installmentmerchant,d.installmentmerchantlocation,d.installmentstartdate,d.installmentloan,d.installmentcycleno,d.installmentcycles,d.installmentregrepayment,d.installmentenddate,d.installmentoutbalance,d.entryno";//,d.approvalcode";
            }
            //else
            //    {
            //    dQry = "select d." + dStatementno + ",d." + dCardno + ",d." + dAccountno + ",d." + dDocno + ",d." + dTransdate + ",d." + dPostingdate + ",d." + dRefereneno + ",d." + dAccountcurrency + ",d." + dOrigtranamount + ",d." + dOrigtrancurrency + ",d." + dBilltranamount + ",d." + dBilltranamountsign + ",d." + dTrandescription + ",d." + dMerchant + ",d." + dInstallmentData + ",d." + dOrigDocNo + ",d." + dInstallmentMerchant + ",d." + dInstallmentMerchantLocation + ",d." + dEntryNo;
            //    }
            //  dQry += " from " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " d "
            //+ " where d.BRANCH = " + curBranch //,'011045350007301'

            if (frmStatementFile.Internal == true && !string.IsNullOrEmpty(frmStatementFile.internalAccNo))
            {
                dQry += " from " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " d "
          + " where d.BRANCH = " + curBranch + "and accountno ='" + frmStatementFile.internalAccNo + "'";
            }
            else
            {
                dQry += " from " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " d "
              + " where d.BRANCH = " + curBranch //,'011045350007301'
                                                 //;
                                                 //}
                                                 //dQry += ""
                                                 // + "and d.accountno IN ('CR_NGN_0000082176', 'Rew_68925' )"
                                                 //+ " and d.accountno in('CR_NGN_0000015960')" //DBN
                                                 //+ " and m.accountno in('Credit_AON_4024','Credit_AON_768')" //BAI
                                                 //+ " and d.accountno in('Credit_586001','Credit_589822','Credit_551066')"
                                                 //+ " and d.accountno in('Credit_586001','Credit_589822','Credit_551066','Credit_512357','Credit_589729')"
                                                 //+ " and d.accountno in('Credit_AON_20')"
                                                 //+ " and d.accountno in('Credit_AON_2676')"
                                                 //+ "and d.CONTRACTNO IN ('MC_Inst_000000001126','52003100000000013040','INST_CARD_0000001123')"
                                                 //+ "and d.CONTRACTNO IN ('Gold_Stand_46575','Rew3_GLD_3822693')"
                                                 //+ "and d.CONTRACTNO IN ('Plat_Staff_000000051','Rew3_PLT_394022')"
                                                 //+ "and d.CONTRACTNO IN ('4042390000023')"
                                                 //+ "and d.CONTRACTNO IN ('40284200000000075364')"
                                                 //+ " and accountno = 'Cre_GHS_0000000056'"
                                                 //+ "and d.CONTRACTNO IN ('Plat_Stand_000002811','Rew3_PLT_392407')"
                                                 //+ " and d.accountno in('CR_NGN_0000005982')"
                                                 //+ " and d.accountno = '99999-326-818-B'"
                                                 //+ " and d.accountno = 'MC_CR_EGP_0000012010'"
                                                 //+ " and d.accountno = 'Prep_EUR_00008697'"
                                                 //+ " and d.accountno = 'MC_CR_KES_0000000173'"
                                                 //+ " and d.accountno = 'Prepaid_0000000017'"
                                                 //+ " and d.accountno = 'Cre_GHS_0000000056'"
                                                 //+ " and d.accountno = 'Prepaid_0000175855'"
                                                 //+ "and accountno = 'Credit_00031387'"
                                                 //+ "and accountno = '4174390000042439'"
                                                 //+ " and statementno in (104818,76304)"
                                                 //+ " and d.accountno in('4042390001394')"
                                                 //+ " and statementno in (select statementno from stmt.tstatementmastercr_16040101 where branch = 23 and cardbranchpart = 1)"
                                                 //+ " and statementno in (select statementno from stmt.tstatementmastercr_16040101 where branch = 23 and externalno IN ('0110081710120501x', '0121175710120601', '1703300','0110253610120501'))"
          ;
            }
            return dQry;
        }
    }// getDetailQuery00002023002

    protected string getDetailQueryForMarkupFee
    {
        get
        {//detailQuery 
            string dQry = string.Empty;
            //if (isGetTotalVal)
            //    {
            //    dQry = "select d." + dStatementno + ",d." + dCardno + ",d." + dAccountno + ",d." + dDocno + ",d." + dTransdate + ",d." + dPostingdate + ",d." + dRefereneno + ",d." + dAccountcurrency + ",d." + dOrigtranamount + ",d." + dOrigtrancurrency + ",d." + dBilltranamount + ",d." + dBilltranamountsign + ",d." + dTrandescription + ",d." + dMerchant + ",m." + mCardprimary;// +",d." + dInstallmentData + ",d." + dOrigDocNo + ",d. dEntryNo +";
            //    dQry += " from " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " d ," + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " m "
            //  + " where d." + dStatementno + " = m." + mStatementno + " and d.BRANCH = " + curBranch + " and m.BRANCH = " + curBranch //,'011045350007301'
            //        //+ "and statementno in (104818,76304)"
            //  ;
            //    }
            //else
            //    {
            //    if (isFullFields)
            dQry = "select d.branch,d.statementno,d.statementnumber,d.contractno,d.accountno,d.accountcurrency,d.cardno,d.mbr,d.transdate,d.postingdate,d.trandescription,d.merchant,d.origtranamount,d.origtrancurrency,               case when d.TRANDESCRIPTION = 'CASH' then NVL(din.billtranamount, 0) + d.billtranamount else d.billtranamount end as billtranamount                          ,d.billtranamountsign,d.docno,d.refereneno,d.installmentdata,d.origdocno,d.installmentmerchant,d.installmentmerchantlocation,d.installmentstartdate,d.installmentloan,d.installmentcycleno,d.installmentcycles,d.installmentregrepayment,d.installmentenddate,d.installmentoutbalance,d.entryno";//,d.approvalcode";
            //else
            //    {
            //    dQry = "select d." + dStatementno + ",d." + dCardno + ",d." + dAccountno + ",d." + dDocno + ",d." + dTransdate + ",d." + dPostingdate + ",d." + dRefereneno + ",d." + dAccountcurrency + ",d." + dOrigtranamount + ",d." + dOrigtrancurrency + ",d." + dBilltranamount + ",d." + dBilltranamountsign + ",d." + dTrandescription + ",d." + dMerchant + ",d." + dInstallmentData + ",d." + dOrigDocNo + ",d." + dInstallmentMerchant + ",d." + dInstallmentMerchantLocation + ",d." + dEntryNo;
            //    }
            //  dQry += " from " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " d "
            //+ " where d.BRANCH = " + curBranch //,'011045350007301'

            if (frmStatementFile.Internal == true && !string.IsNullOrEmpty(frmStatementFile.internalAccNo))
            {
                dQry += " from " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " d "
                    + " left outer join " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " din on din.branch = d.branch and  din.statementno = d.statementno and  din.statementnumber = d.statementnumber and  din.docno = d.docno and  din.refereneno = d.refereneno and UPPER(Din.TRANDESCRIPTION) like '%MARK-UP%' "
          + " where d.BRANCH = " + curBranch + "and accountno ='" + frmStatementFile.internalAccNo + "'";
            }
            else
            {
                dQry += " from " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " d "
                    + " left outer join " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " din on din.branch = d.branch and  din.statementno = d.statementno and  din.statementnumber = d.statementnumber and  din.docno = d.docno and  din.refereneno = d.refereneno and UPPER(Din.TRANDESCRIPTION) like '%MARK-UP%' "
              + " where d.BRANCH = " + curBranch //,'011045350007301'

          ;
            }
            return dQry;
        }
    }




    protected int validateNoOfLines(ArrayList pFileLst, int pNoOfLines)
    {
        int rtrnVal = 0;
        clsValidatePageSize ValidatePageSize = new clsValidatePageSize();

        try
        {
            foreach (string fileNam in pFileLst)
            {
                if (fileNam.EndsWith("_Err.txt") || fileNam.EndsWith("_Err2.txt"))
                    continue;
                //streamWritMD5.WriteLine(clsBasFile.getFileFromPath(fileNam) + "  >>  " + clsBasFile.getFileMD5(fileNam) + "\n\r");// + "\n\r"
                rtrnVal += ValidatePageSize.ValidatePageSize(fileNam, pNoOfLines, "\u000C".ToString());
                //lblStatus.Text = ValidatePageSize.outMessage;
            }
            return rtrnVal;
        }
        catch (NotSupportedException ex)  //(Exception ex)  //
        {
            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        return rtrnVal;
    }



    protected void setContractTable(int pBrach, string pCurrency)
    {
        DSstatement.Tables.Remove("contractTable");
        string contractQuery = "SELECT DISTINCT x.accountcurrency,x.contractno FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and x.accountcurrency = '" + pCurrency + "'  and packagename = 'STTOTABLECORP'"
        //+ " and x.contractno = 'CORP_0000002621'"
        //+ " and x.accountcurrency = 'USD'"
        ;
        OracleDataAdapter contractDA = new OracleDataAdapter(contractQuery, conn);
        contractDA.Fill(DSstatement, "contractTable");

    }


    protected void FillStatementDataSet(int pBrach, string vip)//DataSet public static 
    {
        //vmBranch = 233;
        curBrnch = pBrach.ToString();
        masterQuery = getMasterQuery;
        detailQuery = getDetailQuery;

        getBasicData();
        detailQuery += " and d.postingdate between to_date('" + currStatementdatefrom.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') and to_date('" + currStatementdateto.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') ";
        if (curProduct.Length > 0)
        {
            masterQuery += " and " + curProduct;
            //masterQuery += " and cardproduct ='" + curProduct + "' ";
        }

        if (curCurrency.Length > 0)
        {
            masterQuery += " and m.cardaccountno like '%" + curCurrency + "' ";
            detailQuery += " and d.accountno like '%" + curCurrency + "' ";
        }

        if (strMainTableCond.Length > 0)
        {
            masterQuery += " and " + strMainTableCond;
        }

        if (strSubTableCond.Length > 0)
        {
            detailQuery += " and " + strSubTableCond;
        }


        //strOrder = " m.accountno,m.cardprimary desc,m.cardno ";//,m.cardstatus
        if (isCorporate)
        {
            masterQuery += " and packagename = 'STTOTABLECORP'";
            detailQuery += " and packagename = 'STTOTABLECORP'";
            //strOrder = " m.accountcurrency,m.contractno,m.customerno,m.accountno,m.accountcurrency,m.cardaccountno,m.cardprimary,m.cardstatus,m.cardno ";
            strOrder = " m.contractno,m.accountno,m.cardaccountno,m.cardprimary,m.cardcreationdate ";
            //strOrder = " m.contractno,m.accountno,m.cardaccountno,m.cardprimary,m.cardcreationdate ";

            //if (pBrach == 1)
            //    strOrder = " m.statementnumber ";


            //strOrder = " m.accountcurrency,m.contractno,m.customerno,m.accountno,m.cardaccountno,m.cardprimary,m.cardstatus,m.cardno ";
            masterQuery += " order by m.accountcurrency,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }
        else
        {
            masterQuery += " order by " + strOrder;
            if (curBranch == 29)
                masterQuery += " ,m.cardcreationdate";
            //if (curBranch == 1)//
            //  masterQuery += " order by " + strOrder; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 3)//AWB || curBranch == 3
            //  masterQuery += " order by m.CARDBRANCHPART,m.accountno,m.cardprimary,m.cardstatus,m.cardno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 5)//AIB
            //  masterQuery += " order by m.CARDBRANCHPART,m.externalno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            ////masterQuery += " order by m.CARDBRANCHPART,m.cardno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 6)//AAIB
            //  masterQuery += " order by substr(m.cardbranchpartname,6)," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 10)//BNP
            //  masterQuery += " order by m.cardproduct,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            //else  //
            //  masterQuery += " order by " + strOrderBy + "m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }

        /*    if(curBranch == 5 || curBranch == 3)//AIB
              detailQuery += " order by accountno,cardno,postingdate";
            else
              detailQuery += " order by accountno,cardno,postingdate ";
        */
        detailQuery += strSupHaving;
        //detailQuery += " order by d.accountno,d.cardno,d.postingdate";
        //detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno";
        if (pBrach == 41)
            detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno,d.entryno desc";
        else if (pBrach == 36)
            detailQuery += " order by d.statementno,d.postingdate,d.docno,d.entryno";
        else if (pBrach == 127)
            detailQuery += " order by d.transdate,d.statementno,d.postingdate,d.docno,d.entryno";
        //else if (pBrach == 1)
        //    detailQuery += " order by d.statementnumber,d.postingdate";
        else if (pBrach == 23)
            detailQuery += " order by d.cardno,d.transdate";
        else
            detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno,d.entryno";
        //detailQuery += " order by d.postingdate,d.accountno,d.cardno,d.docno,d.entryno";


        try
        { // change here for select only two to be gerated
            conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter MasterDA = new OracleDataAdapter(masterQuery, conn);
            OracleDataAdapter detailDA = new OracleDataAdapter(detailQuery, conn);

            conn.Open();
            //(new OracleCommand("alter session set optimizer_goal = RULE", conn)).ExecuteNonQuery();
            DSstatement = new DataSet("MasterDetailDS");
            MasterDA.Fill(DSstatement, "tStatementMasterTable");
            detailDA.Fill(DSstatement, "tStatementDetailTable");

            //DataView dv = DSstatement.Tables["tStatementMasterTable"].DefaultView;
            //dv.Sort = "postingdate ASC";
            //DataTable sortedDT = dv.ToTable();

            //clsBasFile.SetAccessRule(@"D:\log.txt");
            //clsBasFile.writeTxtFile(@"D:\log.txt", Environment.NewLine + masterQuery, System.IO.FileMode.OpenOrCreate);
            //clsBasFile.writeTxtFile(@"D:\log.txt", Environment.NewLine + detailQuery, System.IO.FileMode.OpenOrCreate);
            if (isCorporate)
            {
                //string contractQuery = "SELECT DISTINCT x.contractno FROM tstatementmastertable x where x.branch = " + curBranch + " order by x.accountcurrency";// + " and x.contractno = 'CORP_EGP_1069140001'"
                string contractQuery = "SELECT DISTINCT x.accountcurrency,x.contractno FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and packagename = 'STTOTABLECORP' order by x.accountcurrency";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter contractDA = new OracleDataAdapter(contractQuery, conn);
                contractDA.Fill(DSstatement, "contractTable");
                // create currency table
                string currencyQuery = "SELECT DISTINCT x.accountcurrency FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and x.accountcurrency is not null and packagename = 'STTOTABLECORP'";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter currencyDA = new OracleDataAdapter(currencyQuery, conn);
                currencyDA.Fill(DSstatement, "currencyTable");
            }
            if (isCredite)
            {
                //if (!isGetTotalVal)
                //{
                try
                {
                    /*
                     if (pBrach == 1 && isCorporate)
                        StatementNoDRel = DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementnumber], DSstatement.Tables["tStatementDetailTable"].Columns[dStatementnumber], false);
                    else
                        StatementNoDRel = DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementno], DSstatement.Tables["tStatementDetailTable"].Columns[dStatementno], false);
                    */
                    //if (pBrach == 128)
                    //{
                    //    StatementNoDRel = DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementno], DSstatement.Tables["tStatementDetailTable"].Columns[dStatementno], false);
                    //}
                    //else
                    //{
                    StatementNoDRel = DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementno], DSstatement.Tables["tStatementDetailTable"].Columns[dStatementno], false);
                    //}
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
                //}
            }
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        //return DSstatement;
    }

    protected void FillStatementDataSet_WithOverDueDays(int pBrach, string vip)//DataSet public static 
    {
        //vmBranch = 233;
        curBrnch = pBrach.ToString();
        masterQuery = getMasterQueryWithOverDueDays;
        detailQuery = getDetailQuery;

        getBasicData();
        detailQuery += " and d.postingdate between to_date('" + currStatementdatefrom.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') and to_date('" + currStatementdateto.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') ";
        if (curProduct.Length > 0)
        {
            masterQuery += " and " + curProduct;
            //masterQuery += " and cardproduct ='" + curProduct + "' ";
        }

        if (curCurrency.Length > 0)
        {
            masterQuery += " and m.cardaccountno like '%" + curCurrency + "' ";
            detailQuery += " and d.accountno like '%" + curCurrency + "' ";
        }

        if (strMainTableCond.Length > 0)
        {
            masterQuery += " and " + strMainTableCond;
        }

        if (strSubTableCond.Length > 0)
        {
            detailQuery += " and " + strSubTableCond;
        }


        //strOrder = " m.accountno,m.cardprimary desc,m.cardno ";//,m.cardstatus
        if (isCorporate)
        {
            masterQuery += " and packagename = 'STTOTABLECORP'";
            detailQuery += " and packagename = 'STTOTABLECORP'";
            //strOrder = " m.accountcurrency,m.contractno,m.customerno,m.accountno,m.accountcurrency,m.cardaccountno,m.cardprimary,m.cardstatus,m.cardno ";
            strOrder = " m.contractno,m.accountno,m.cardaccountno,m.cardprimary,m.cardcreationdate ";
            //strOrder = " m.contractno,m.accountno,m.cardaccountno,m.cardprimary,m.cardcreationdate ";

            //if (pBrach == 1)
            //    strOrder = " m.statementnumber ";


            //strOrder = " m.accountcurrency,m.contractno,m.customerno,m.accountno,m.cardaccountno,m.cardprimary,m.cardstatus,m.cardno ";
            masterQuery += " order by m.accountcurrency,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }
        else
        {
            if (curBranch == 6)
                masterQuery += " order by m.accountno,m.cardprimary desc,m.cardno ";
            else
                masterQuery += " order by " + strOrder;

            //if (curBranch == 1)//
            //  masterQuery += " order by " + strOrder; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 3)//AWB || curBranch == 3
            //  masterQuery += " order by m.CARDBRANCHPART,m.accountno,m.cardprimary,m.cardstatus,m.cardno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 5)//AIB
            //  masterQuery += " order by m.CARDBRANCHPART,m.externalno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            ////masterQuery += " order by m.CARDBRANCHPART,m.cardno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 6)//AAIB
            //  masterQuery += " order by substr(m.cardbranchpartname,6)," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 10)//BNP
            //  masterQuery += " order by m.cardproduct,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            //else  //
            //  masterQuery += " order by " + strOrderBy + "m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }

        /*    if(curBranch == 5 || curBranch == 3)//AIB
              detailQuery += " order by accountno,cardno,postingdate";
            else
              detailQuery += " order by accountno,cardno,postingdate ";
        */
        detailQuery += strSupHaving;
        //detailQuery += " order by d.accountno,d.cardno,d.postingdate";
        //detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno";
        if (pBrach == 41)
            detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno,d.entryno desc";
        else if (pBrach == 36)
            detailQuery += " order by d.statementno,d.postingdate,d.docno,d.entryno";
        //else if (pBrach == 1)
        //    detailQuery += " order by d.statementnumber,d.postingdate";
        else if (pBrach == 23)
            detailQuery += " order by d.cardno,d.transdate";
        else
            detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno,d.entryno";
        //detailQuery += " order by d.postingdate,d.accountno,d.cardno,d.docno,d.entryno";


        try
        {
            conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter MasterDA = new OracleDataAdapter(masterQuery, conn);
            OracleDataAdapter detailDA = new OracleDataAdapter(detailQuery, conn);

            conn.Open();
            //(new OracleCommand("alter session set optimizer_goal = RULE", conn)).ExecuteNonQuery();
            DSstatement = new DataSet("MasterDetailDS");
            MasterDA.Fill(DSstatement, "tStatementMasterTable");
            detailDA.Fill(DSstatement, "tStatementDetailTable");

            //DataView dv = DSstatement.Tables["tStatementMasterTable"].DefaultView;
            //dv.Sort = "postingdate ASC";
            //DataTable sortedDT = dv.ToTable();

            //clsBasFile.SetAccessRule(@"D:\log.txt");
            //clsBasFile.writeTxtFile(@"D:\log.txt", Environment.NewLine + masterQuery, System.IO.FileMode.OpenOrCreate);
            //clsBasFile.writeTxtFile(@"D:\log.txt", Environment.NewLine + detailQuery, System.IO.FileMode.OpenOrCreate);
            if (isCorporate)
            {
                //string contractQuery = "SELECT DISTINCT x.contractno FROM tstatementmastertable x where x.branch = " + curBranch + " order by x.accountcurrency";// + " and x.contractno = 'CORP_EGP_1069140001'"
                string contractQuery = "SELECT DISTINCT x.accountcurrency,x.contractno FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and packagename = 'STTOTABLECORP' order by x.accountcurrency";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter contractDA = new OracleDataAdapter(contractQuery, conn);
                contractDA.Fill(DSstatement, "contractTable");
                // create currency table
                string currencyQuery = "SELECT DISTINCT x.accountcurrency FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and x.accountcurrency is not null and packagename = 'STTOTABLECORP'";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter currencyDA = new OracleDataAdapter(currencyQuery, conn);
                currencyDA.Fill(DSstatement, "currencyTable");
            }
            if (isCredite)
            {
                //if (!isGetTotalVal)
                //{
                try
                {
                    /*
                     if (pBrach == 1 && isCorporate)
                        StatementNoDRel = DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementnumber], DSstatement.Tables["tStatementDetailTable"].Columns[dStatementnumber], false);
                    else
                        StatementNoDRel = DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementno], DSstatement.Tables["tStatementDetailTable"].Columns[dStatementno], false);
                    */
                    StatementNoDRel = DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementno], DSstatement.Tables["tStatementDetailTable"].Columns[dStatementno], false);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
                //}
            }
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        //return DSstatement;
    }

    protected void FillStatementDataSet_Exclude_VisaCards(int pBrach, string vip)//DataSet public static 
    {
        //vmBranch = 233;
        curBrnch = pBrach.ToString();
        masterQuery = getMasterQuery;
        detailQuery = getDetailQuery;

        getBasicData();
        detailQuery += " and d.postingdate between to_date('" + currStatementdatefrom.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') and to_date('" + currStatementdateto.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') ";
        if (curProduct.Length > 0)
        {
            masterQuery += " and " + curProduct;
            //masterQuery += " and cardproduct ='" + curProduct + "' ";
        }

        if (curCurrency.Length > 0)
        {
            masterQuery += " and m.cardaccountno like '%" + curCurrency + "' ";
            detailQuery += " and d.accountno like '%" + curCurrency + "' ";
        }

        masterQuery += " and m.cardno not like '4%'  ";

        if (strMainTableCond.Length > 0)
        {
            masterQuery += " and " + strMainTableCond;
        }

        detailQuery += " and d.cardno not like '4%'  ";

        if (strSubTableCond.Length > 0)
        {
            detailQuery += " and " + strSubTableCond;
        }


        //strOrder = " m.accountno,m.cardprimary desc,m.cardno ";//,m.cardstatus
        if (isCorporate)
        {
            masterQuery += " and packagename = 'STTOTABLECORP'";
            detailQuery += " and packagename = 'STTOTABLECORP'";
            //strOrder = " m.accountcurrency,m.contractno,m.customerno,m.accountno,m.accountcurrency,m.cardaccountno,m.cardprimary,m.cardstatus,m.cardno ";
            strOrder = " m.contractno,m.accountno,m.cardaccountno,m.cardprimary,m.cardcreationdate ";
            //strOrder = " m.contractno,m.accountno,m.cardaccountno,m.cardprimary,m.cardcreationdate ";

            //if (pBrach == 1)
            //    strOrder = " m.statementnumber ";


            //strOrder = " m.accountcurrency,m.contractno,m.customerno,m.accountno,m.cardaccountno,m.cardprimary,m.cardstatus,m.cardno ";
            masterQuery += " order by m.accountcurrency,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }
        else
        {
            masterQuery += " order by " + strOrder;
            //if (curBranch == 1)//
            //  masterQuery += " order by " + strOrder; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 3)//AWB || curBranch == 3
            //  masterQuery += " order by m.CARDBRANCHPART,m.accountno,m.cardprimary,m.cardstatus,m.cardno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 5)//AIB
            //  masterQuery += " order by m.CARDBRANCHPART,m.externalno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            ////masterQuery += " order by m.CARDBRANCHPART,m.cardno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 6)//AAIB
            //  masterQuery += " order by substr(m.cardbranchpartname,6)," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 10)//BNP
            //  masterQuery += " order by m.cardproduct,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            //else  //
            //  masterQuery += " order by " + strOrderBy + "m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }

        /*    if(curBranch == 5 || curBranch == 3)//AIB
              detailQuery += " order by accountno,cardno,postingdate";
            else
              detailQuery += " order by accountno,cardno,postingdate ";
        */
        detailQuery += strSupHaving;
        //detailQuery += " order by d.accountno,d.cardno,d.postingdate";
        //detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno";
        if (pBrach == 41)
            detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno,d.entryno desc";
        else if (pBrach == 36)
            detailQuery += " order by d.statementno,d.postingdate,d.docno,d.entryno";
        //else if (pBrach == 1)
        //    detailQuery += " order by d.statementnumber,d.postingdate";
        else if (pBrach == 23)
            detailQuery += " order by d.cardno,d.transdate";
        else
            detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno,d.entryno";
        //detailQuery += " order by d.postingdate,d.accountno,d.cardno,d.docno,d.entryno";


        try
        {
            conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter MasterDA = new OracleDataAdapter(masterQuery, conn);
            OracleDataAdapter detailDA = new OracleDataAdapter(detailQuery, conn);

            conn.Open();
            //(new OracleCommand("alter session set optimizer_goal = RULE", conn)).ExecuteNonQuery();
            DSstatement = new DataSet("MasterDetailDS");
            MasterDA.Fill(DSstatement, "tStatementMasterTable");
            detailDA.Fill(DSstatement, "tStatementDetailTable");

            //DataView dv = DSstatement.Tables["tStatementMasterTable"].DefaultView;
            //dv.Sort = "postingdate ASC";
            //DataTable sortedDT = dv.ToTable();

            //clsBasFile.SetAccessRule(@"D:\log.txt");
            //clsBasFile.writeTxtFile(@"D:\log.txt", Environment.NewLine + masterQuery, System.IO.FileMode.OpenOrCreate);
            //clsBasFile.writeTxtFile(@"D:\log.txt", Environment.NewLine + detailQuery, System.IO.FileMode.OpenOrCreate);
            if (isCorporate)
            {
                //string contractQuery = "SELECT DISTINCT x.contractno FROM tstatementmastertable x where x.branch = " + curBranch + " order by x.accountcurrency";// + " and x.contractno = 'CORP_EGP_1069140001'"
                string contractQuery = "SELECT DISTINCT x.accountcurrency,x.contractno FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and packagename = 'STTOTABLECORP' order by x.accountcurrency";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter contractDA = new OracleDataAdapter(contractQuery, conn);
                contractDA.Fill(DSstatement, "contractTable");
                // create currency table
                string currencyQuery = "SELECT DISTINCT x.accountcurrency FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and x.accountcurrency is not null and packagename = 'STTOTABLECORP'";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter currencyDA = new OracleDataAdapter(currencyQuery, conn);
                currencyDA.Fill(DSstatement, "currencyTable");
            }
            if (isCredite)
            {
                //if (!isGetTotalVal)
                //{
                try
                {
                    /*
                     if (pBrach == 1 && isCorporate)
                        StatementNoDRel = DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementnumber], DSstatement.Tables["tStatementDetailTable"].Columns[dStatementnumber], false);
                    else
                        StatementNoDRel = DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementno], DSstatement.Tables["tStatementDetailTable"].Columns[dStatementno], false);
                    */
                    StatementNoDRel = DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementno], DSstatement.Tables["tStatementDetailTable"].Columns[dStatementno], false);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
                //}
            }
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        //return DSstatement;
    }

    protected void FillStatementDataSet_SortCardPriority(int pBrach, string vip)//DataSet public static 
    {
        //vmBranch = 233;
        curBrnch = pBrach.ToString();
        masterQuery = getMasterQuery.Replace(" from ", " , ROW_NUMBER() OVER ( PARTITION BY  m.STATEMENTNUMBER ORDER BY m.STATEMENTNUMBER ASC, CASE WHEN m.cardstate = 'Given' THEN 1 WHEN m.cardstate = 'New Pin Generation Only' THEN 2 WHEN m.cardstate = 'New Pin Generated Only' THEN 3 WHEN m.cardstate = 'Embossed' THEN 4 WHEN m.cardstate = 'Embossing' THEN 5  WHEN m.cardstate = 'New' THEN 6 WHEN m.cardstate = 'Entered' THEN 7 WHEN m.cardstate = 'Lost Card' THEN 8 WHEN m.cardstate = 'Lost' THEN 9 WHEN m.cardstate = 'Stolen Card' THEN 10 WHEN m.cardstate = 'Stolen' THEN 11 WHEN m.cardstate = 'CLSB' THEN 12 WHEN m.cardstate = 'Compromised' THEN 13 WHEN m.cardstate = 'Last name change' THEN 14 WHEN m.cardstate = 'Damaged' THEN 15 WHEN m.cardstate = 'Cancelled' THEN 16 WHEN m.cardstate = 'Expired' THEN 17 ELSE 18 END ASC ) AS r_num  from   ");
        detailQuery = getDetailQuery;

        getBasicData();
        detailQuery += " and d.postingdate between to_date('" + currStatementdatefrom.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') and to_date('" + currStatementdateto.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') ";
        if (curProduct.Length > 0)
        {
            masterQuery += " and " + curProduct;
            //masterQuery += " and cardproduct ='" + curProduct + "' ";
        }

        if (curCurrency.Length > 0)
        {
            masterQuery += " and m.cardaccountno like '%" + curCurrency + "' ";
            detailQuery += " and d.accountno like '%" + curCurrency + "' ";
        }

        if (strMainTableCond.Length > 0)
        {
            masterQuery += " and " + strMainTableCond;
        }

        if (strSubTableCond.Length > 0)
        {
            detailQuery += " and " + strSubTableCond;
        }

        string strOrderCardPriority = " m.STATEMENTNUMBER ASC, CASE WHEN m.cardstate = 'Given' THEN 1 WHEN m.cardstate = 'New Pin Generation Only' THEN 2 WHEN m.cardstate = 'New Pin Generated Only' THEN 3 WHEN m.cardstate = 'Embossed' THEN 4 WHEN m.cardstate = 'Embossing' THEN 5  WHEN m.cardstate = 'New' THEN 6 WHEN m.cardstate = 'Entered' THEN 7 WHEN m.cardstate = 'Lost Card' THEN 8 WHEN m.cardstate = 'Lost' THEN 9 WHEN m.cardstate = 'Stolen Card' THEN 10 WHEN m.cardstate = 'Stolen' THEN 11 WHEN m.cardstate = 'CLSB' THEN 12 WHEN m.cardstate = 'Compromised' THEN 13 WHEN m.cardstate = 'Last name change' THEN 14 WHEN m.cardstate = 'Damaged' THEN 15 WHEN m.cardstate = 'Cancelled' THEN 16 WHEN m.cardstate = 'Expired' THEN 17 ELSE 18 END ASC  , CardPrimary DESC ,  ";

        //strOrder = " m.accountno,m.cardprimary desc,m.cardno ";//,m.cardstatus
        if (isCorporate)
        {
            masterQuery += " and packagename = 'STTOTABLECORP'";
            detailQuery += " and packagename = 'STTOTABLECORP'";
            //strOrder = " m.accountcurrency,m.contractno,m.customerno,m.accountno,m.accountcurrency,m.cardaccountno,m.cardprimary,m.cardstatus,m.cardno ";
            strOrder = " m.contractno,m.accountno,m.cardaccountno,m.cardprimary,m.cardcreationdate ";
            //strOrder = " m.accountcurrency,m.contractno,m.customerno,m.accountno,m.cardaccountno,m.cardprimary,m.cardstatus,m.cardno ";
            masterQuery += " order by m.accountcurrency,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }
        else
        {
            masterQuery += " order by " + strOrderCardPriority + strOrder;
            //if (curBranch == 1)//
            //  masterQuery += " order by " + strOrder; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 3)//AWB || curBranch == 3
            //  masterQuery += " order by m.CARDBRANCHPART,m.accountno,m.cardprimary,m.cardstatus,m.cardno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 5)//AIB
            //  masterQuery += " order by m.CARDBRANCHPART,m.externalno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            ////masterQuery += " order by m.CARDBRANCHPART,m.cardno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 6)//AAIB
            //  masterQuery += " order by substr(m.cardbranchpartname,6)," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 10)//BNP
            //  masterQuery += " order by m.cardproduct,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            //else  //
            //  masterQuery += " order by " + strOrderBy + "m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }

        /*    if(curBranch == 5 || curBranch == 3)//AIB
              detailQuery += " order by accountno,cardno,postingdate";
            else
              detailQuery += " order by accountno,cardno,postingdate ";
        */
        detailQuery += strSupHaving;
        //detailQuery += " order by d.accountno,d.cardno,d.postingdate";
        //detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno";
        if (pBrach == 41)
            detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno,d.entryno desc";
        else if (pBrach == 36)
            detailQuery += " order by d.statementno,d.postingdate,d.docno,d.entryno";
        else if (pBrach == 23)
            detailQuery += " order by d.cardno,d.transdate";
        else
            detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno,d.entryno";
        //detailQuery += " order by d.postingdate,d.accountno,d.cardno,d.docno,d.entryno";


        try
        {
            conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter MasterDA = new OracleDataAdapter(masterQuery, conn);
            OracleDataAdapter detailDA = new OracleDataAdapter(detailQuery, conn);

            conn.Open();
            //(new OracleCommand("alter session set optimizer_goal = RULE", conn)).ExecuteNonQuery();
            DSstatement = new DataSet("MasterDetailDS");
            MasterDA.Fill(DSstatement, "tStatementMasterTable");
            detailDA.Fill(DSstatement, "tStatementDetailTable");

            //DSstatement.Tables["tStatementMasterTable"]

            for (int i = DSstatement.Tables["tStatementMasterTable"].Rows.Count - 1; i >= 0; i--)
            {
                //DataRow dr = DSstatement.Tables["tStatementMasterTable"].Rows[i];
                string strr_num = DSstatement.Tables["tStatementMasterTable"].Rows[i]["r_num"].ToString().Trim();
                if (strr_num == "1")
                    continue;
                else
                    DSstatement.Tables["tStatementMasterTable"].Rows[i].Delete();
            }
            DSstatement.Tables["tStatementMasterTable"].AcceptChanges();

            DSstatement.Tables["tStatementMasterTable"].Columns.Remove("r_num");

            //clsBasFile.SetAccessRule(@"D:\log.txt");
            //clsBasFile.writeTxtFile(@"D:\log.txt", Environment.NewLine + masterQuery, System.IO.FileMode.OpenOrCreate);
            //clsBasFile.writeTxtFile(@"D:\log.txt", Environment.NewLine + detailQuery, System.IO.FileMode.OpenOrCreate);
            if (isCorporate)
            {
                //string contractQuery = "SELECT DISTINCT x.contractno FROM tstatementmastertable x where x.branch = " + curBranch + " order by x.accountcurrency";// + " and x.contractno = 'CORP_EGP_1069140001'"
                string contractQuery = "SELECT DISTINCT x.accountcurrency,x.contractno FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and packagename = 'STTOTABLECORP' order by x.accountcurrency";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter contractDA = new OracleDataAdapter(contractQuery, conn);
                contractDA.Fill(DSstatement, "contractTable");
                // create currency table
                string currencyQuery = "SELECT DISTINCT x.accountcurrency FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and x.accountcurrency is not null and packagename = 'STTOTABLECORP'";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter currencyDA = new OracleDataAdapter(currencyQuery, conn);
                currencyDA.Fill(DSstatement, "currencyTable");
            }
            if (isCredite)
            {
                //if (!isGetTotalVal)
                //{
                try
                {
                    StatementNoDRel = DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementnumber], DSstatement.Tables["tStatementDetailTable"].Columns[dStatementnumber], false);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
                //}
            }
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        //return DSstatement;
    }

    // Exclude all visa cards ---------------- Jira ALXB-3622
    protected void FillStatementDataSet_SortCardPriority_Exclude_VisaCards(int pBrach, string vip)//DataSet public static 
    {
        //vmBranch = 233;
        curBrnch = pBrach.ToString();
        masterQuery = getMasterQuery.Replace(" from ", " , ROW_NUMBER() OVER ( PARTITION BY  m.STATEMENTNUMBER ORDER BY m.STATEMENTNUMBER ASC, CASE WHEN m.cardstate = 'Given' THEN 1 WHEN m.cardstate = 'New Pin Generation Only' THEN 2 WHEN m.cardstate = 'New Pin Generated Only' THEN 3 WHEN m.cardstate = 'Embossed' THEN 4 WHEN m.cardstate = 'Embossing' THEN 5  WHEN m.cardstate = 'New' THEN 6 WHEN m.cardstate = 'Entered' THEN 7 WHEN m.cardstate = 'Lost Card' THEN 8 WHEN m.cardstate = 'Lost' THEN 9 WHEN m.cardstate = 'Stolen Card' THEN 10 WHEN m.cardstate = 'Stolen' THEN 11 WHEN m.cardstate = 'CLSB' THEN 12 WHEN m.cardstate = 'Compromised' THEN 13 WHEN m.cardstate = 'Last name change' THEN 14 WHEN m.cardstate = 'Damaged' THEN 15 WHEN m.cardstate = 'Cancelled' THEN 16 WHEN m.cardstate = 'Expired' THEN 17 ELSE 18 END ASC ) AS r_num  from   ");
        detailQuery = getDetailQuery;

        getBasicData();
        detailQuery += " and d.postingdate between to_date('" + currStatementdatefrom.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') and to_date('" + currStatementdateto.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') ";
        if (curProduct.Length > 0)
        {
            masterQuery += " and " + curProduct;
            //masterQuery += " and cardproduct ='" + curProduct + "' ";
        }

        if (curCurrency.Length > 0)
        {
            masterQuery += " and m.cardaccountno like '%" + curCurrency + "' ";
            detailQuery += " and d.accountno like '%" + curCurrency + "' ";
        }

        masterQuery += " and m.cardno not like '4%'  ";

        if (strMainTableCond.Length > 0)
        {
            masterQuery += " and " + strMainTableCond;
        }

        detailQuery += " and d.cardno not like '4%'  ";

        if (strSubTableCond.Length > 0)
        {
            detailQuery += " and " + strSubTableCond;
        }

        string strOrderCardPriority = " m.STATEMENTNUMBER ASC, CASE WHEN m.cardstate = 'Given' THEN 1 WHEN m.cardstate = 'New Pin Generation Only' THEN 2 WHEN m.cardstate = 'New Pin Generated Only' THEN 3 WHEN m.cardstate = 'Embossed' THEN 4 WHEN m.cardstate = 'Embossing' THEN 5  WHEN m.cardstate = 'New' THEN 6 WHEN m.cardstate = 'Entered' THEN 7 WHEN m.cardstate = 'Lost Card' THEN 8 WHEN m.cardstate = 'Lost' THEN 9 WHEN m.cardstate = 'Stolen Card' THEN 10 WHEN m.cardstate = 'Stolen' THEN 11 WHEN m.cardstate = 'CLSB' THEN 12 WHEN m.cardstate = 'Compromised' THEN 13 WHEN m.cardstate = 'Last name change' THEN 14 WHEN m.cardstate = 'Damaged' THEN 15 WHEN m.cardstate = 'Cancelled' THEN 16 WHEN m.cardstate = 'Expired' THEN 17 ELSE 18 END ASC  , CardPrimary DESC ,  ";

        //strOrder = " m.accountno,m.cardprimary desc,m.cardno ";//,m.cardstatus
        if (isCorporate)
        {
            masterQuery += " and packagename = 'STTOTABLECORP'";
            detailQuery += " and packagename = 'STTOTABLECORP'";
            //strOrder = " m.accountcurrency,m.contractno,m.customerno,m.accountno,m.accountcurrency,m.cardaccountno,m.cardprimary,m.cardstatus,m.cardno ";
            strOrder = " m.contractno,m.accountno,m.cardaccountno,m.cardprimary,m.cardcreationdate ";
            //strOrder = " m.accountcurrency,m.contractno,m.customerno,m.accountno,m.cardaccountno,m.cardprimary,m.cardstatus,m.cardno ";
            masterQuery += " order by m.accountcurrency,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }
        else
        {
            masterQuery += " order by " + strOrderCardPriority + strOrder;
            //if (curBranch == 1)//
            //  masterQuery += " order by " + strOrder; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 3)//AWB || curBranch == 3
            //  masterQuery += " order by m.CARDBRANCHPART,m.accountno,m.cardprimary,m.cardstatus,m.cardno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 5)//AIB
            //  masterQuery += " order by m.CARDBRANCHPART,m.externalno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            ////masterQuery += " order by m.CARDBRANCHPART,m.cardno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 6)//AAIB
            //  masterQuery += " order by substr(m.cardbranchpartname,6)," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 10)//BNP
            //  masterQuery += " order by m.cardproduct,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            //else  //
            //  masterQuery += " order by " + strOrderBy + "m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }

        /*    if(curBranch == 5 || curBranch == 3)//AIB
              detailQuery += " order by accountno,cardno,postingdate";
            else
              detailQuery += " order by accountno,cardno,postingdate ";
        */
        detailQuery += strSupHaving;
        //detailQuery += " order by d.accountno,d.cardno,d.postingdate";
        //detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno";
        if (pBrach == 41)
            detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno,d.entryno desc";
        else if (pBrach == 36)
            detailQuery += " order by d.statementno,d.postingdate,d.docno,d.entryno";
        else if (pBrach == 23)
            detailQuery += " order by d.cardno,d.transdate";
        else
            detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno,d.entryno";
        //detailQuery += " order by d.postingdate,d.accountno,d.cardno,d.docno,d.entryno";


        try
        {
            conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter MasterDA = new OracleDataAdapter(masterQuery, conn);
            OracleDataAdapter detailDA = new OracleDataAdapter(detailQuery, conn);

            conn.Open();
            //(new OracleCommand("alter session set optimizer_goal = RULE", conn)).ExecuteNonQuery();
            DSstatement = new DataSet("MasterDetailDS");
            MasterDA.Fill(DSstatement, "tStatementMasterTable");
            detailDA.Fill(DSstatement, "tStatementDetailTable");

            //DSstatement.Tables["tStatementMasterTable"]

            for (int i = DSstatement.Tables["tStatementMasterTable"].Rows.Count - 1; i >= 0; i--)
            {
                //DataRow dr = DSstatement.Tables["tStatementMasterTable"].Rows[i];
                string strr_num = DSstatement.Tables["tStatementMasterTable"].Rows[i]["r_num"].ToString().Trim();
                if (strr_num == "1")
                    continue;
                else
                    DSstatement.Tables["tStatementMasterTable"].Rows[i].Delete();
            }
            DSstatement.Tables["tStatementMasterTable"].AcceptChanges();

            DSstatement.Tables["tStatementMasterTable"].Columns.Remove("r_num");

            //clsBasFile.SetAccessRule(@"D:\log.txt");
            //clsBasFile.writeTxtFile(@"D:\log.txt", Environment.NewLine + masterQuery, System.IO.FileMode.OpenOrCreate);
            //clsBasFile.writeTxtFile(@"D:\log.txt", Environment.NewLine + detailQuery, System.IO.FileMode.OpenOrCreate);
            if (isCorporate)
            {
                //string contractQuery = "SELECT DISTINCT x.contractno FROM tstatementmastertable x where x.branch = " + curBranch + " order by x.accountcurrency";// + " and x.contractno = 'CORP_EGP_1069140001'"
                string contractQuery = "SELECT DISTINCT x.accountcurrency,x.contractno FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and packagename = 'STTOTABLECORP' order by x.accountcurrency";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter contractDA = new OracleDataAdapter(contractQuery, conn);
                contractDA.Fill(DSstatement, "contractTable");
                // create currency table
                string currencyQuery = "SELECT DISTINCT x.accountcurrency FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and x.accountcurrency is not null and packagename = 'STTOTABLECORP'";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter currencyDA = new OracleDataAdapter(currencyQuery, conn);
                currencyDA.Fill(DSstatement, "currencyTable");
            }
            if (isCredite)
            {
                //if (!isGetTotalVal)
                //{
                try
                {
                    StatementNoDRel = DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementnumber], DSstatement.Tables["tStatementDetailTable"].Columns[dStatementnumber], false);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
                //}
            }
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        //return DSstatement;
    }




    protected void FillStatementDataSet_UNBN(int pBrach, string vip)//DataSet public static 
    {
        //vmBranch = 233;
        curBrnch = pBrach.ToString();
        masterQuery = getMasterQuery;
        detailQuery = getDetailQuery;
        string masterQuery_1 = getMasterQuery;

        ////" SELECT TSTATEMENTMASTERTABLE.CLIENTID,TSTATEMENTMASTERTABLE.STATEMENTNO,TSTATEMENTMASTERTABLE.BRANCH, TSTATEMENTMASTERTABLE.CARDNO, TSTATEMENTMASTERTABLE.ACCOUNTNO, TSTATEMENTMASTERTABLE.CARDACCOUNTNO, TSTATEMENTMASTERTABLE.CONTRACTNO, TSTATEMENTMASTERTABLE.CARDPRIMARY, TSTATEMENTMASTERTABLE.CARDSTATUS, TSTATEMENTMASTERTABLE.ACCOUNTCURRENCY,  TSTATEMENTMASTERTABLE.CUSTOMERNAME, TSTATEMENTMASTERTABLE.CUSTOMERADDRESS1, TSTATEMENTMASTERTABLE.CUSTOMERADDRESS2, TSTATEMENTMASTERTABLE.CUSTOMERADDRESS3, TSTATEMENTMASTERTABLE.CUSTOMERREGION, TSTATEMENTMASTERTABLE.CUSTOMERCITY, TSTATEMENTMASTERTABLE.CARDPRODUCT, TSTATEMENTMASTERTABLE.CARDBRANCHPARTNAME, TSTATEMENTMASTERTABLE.EXTERNALNO, TSTATEMENTMASTERTABLE.STATEMENTDATETO, TSTATEMENTMASTERTABLE.ACCOUNTLIM, TSTATEMENTMASTERTABLE.ACCOUNTAVAILABLELIM, TSTATEMENTMASTERTABLE.MINDUEAMOUNT, TSTATEMENTMASTERTABLE.STETEMENTDUEDATE, TSTATEMENTMASTERTABLE.TOTALOVERDUEAMOUNT,  TSTATEMENTMASTERTABLE.OPENINGBALANCE, TSTATEMENTMASTERTABLE.CLOSINGBALANCE, TSTATEMENTMASTERTABLE.PRINARYCARDNO, TSTATEMENTMASTERTABLE.TOTALPAYMENTS, TSTATEMENTMASTERTABLE.TOTALPURCHASES, TSTATEMENTMASTERTABLE.TOTALCASHWITHDRAWAL, TSTATEMENTMASTERTABLE.TOTALINTEREST, TSTATEMENTMASTERTABLE.TOTALCHARGES, TSTATEMENTMASTERTABLE.CONTRACTTYPE , ( Select R.OPENINGBALANCE From STMT.TSTATEMENTMASTERCR_18051502 R where R.CONTRACTTYPE = 'Credit Card Rewards Program' AND R.CLIENTID = TSTATEMENTMASTERTABLE.CLIENTID AND R.BRANCH=TSTATEMENTMASTERTABLE.BRANCH  ) as REWARDOPENINGBALANCE, ( Select R.EARNEDBONUS From STMT.TSTATEMENTMASTERCR_18051502 R where R.CONTRACTTYPE = 'Credit Card Rewards Program' AND R.CLIENTID = TSTATEMENTMASTERTABLE.CLIENTID AND R.BRANCH=TSTATEMENTMASTERTABLE.BRANCH  ) as REWARDEARNEDBONUS, ( Select R.REDEEMEDBONUS From STMT.TSTATEMENTMASTERCR_18051502 R where R.CONTRACTTYPE = 'Credit Card Rewards Program' AND R.CLIENTID = TSTATEMENTMASTERTABLE.CLIENTID AND R.BRANCH=TSTATEMENTMASTERTABLE.BRANCH  ) as REWARDREDEEMEDBONUS, ( Select R.CLOSINGBALANCE From STMT.TSTATEMENTMASTERCR_18051502 R where R.CONTRACTTYPE = 'Credit Card Rewards Program' AND R.CLIENTID = TSTATEMENTMASTERTABLE.CLIENTID AND R.BRANCH=TSTATEMENTMASTERTABLE.BRANCH  ) as REWARDCLOSINGBALANCE , TSTATEMENTDETAILTABLE.TRANSDATE, TSTATEMENTDETAILTABLE.POSTINGDATE, TSTATEMENTDETAILTABLE.MERCHANT, TSTATEMENTDETAILTABLE.TRANDESCRIPTION, TSTATEMENTDETAILTABLE.ORIGTRANCURRENCY, TSTATEMENTDETAILTABLE.ORIGTRANAMOUNT, TSTATEMENTDETAILTABLE.BILLTRANAMOUNT, TSTATEMENTDETAILTABLE.BILLTRANAMOUNTSIGN, TSTATEMENTDETAILTABLE.BRANCH, TSTATEMENTDETAILTABLE.DOCNO, TSTATEMENTDETAILTABLE.STATEMENTNO "
        //  " SELECT * FROM   STMT.TSTATEMENTMASTERCR_18051502   "
        //        //+ "  Left OUTER JOIN STMT.TSTATEMENTDETAILCR_18051502 TSTATEMENTDETAILTABLE on TSTATEMENTMASTERTABLE.BRANCH = TSTATEMENTDETAILTABLE.BRANCH AND TSTATEMENTMASTERTABLE.STATEMENTNO = TSTATEMENTDETAILTABLE.STATEMENTNO "
        //+ " Where BRANCH=115  AND CONTRACTTYPE = 'Credit Card Rewards Program' ";

        masterQuery += " AND CONTRACTTYPE <> 'Credit Card Rewards Program' ";
        masterQuery_1 += " AND CONTRACTTYPE = 'Credit Card Rewards Program' ";


        getBasicData();
        detailQuery += " and d.postingdate between to_date('" + currStatementdatefrom.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') and to_date('" + currStatementdateto.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') ";
        if (curProduct.Length > 0)
        {
            masterQuery += " and " + curProduct;
            masterQuery_1 += " and " + curProduct;
            //masterQuery += " and cardproduct ='" + curProduct + "' ";
        }

        if (curCurrency.Length > 0)
        {
            masterQuery += " and m.cardaccountno like '%" + curCurrency + "' ";
            masterQuery_1 += " and m.cardaccountno like '%" + curCurrency + "' ";

            detailQuery += " and d.accountno like '%" + curCurrency + "' ";
        }

        if (strMainTableCond.Length > 0)
        {
            masterQuery += " and " + strMainTableCond;
            masterQuery_1 += " and " + strMainTableCond;
        }

        if (strSubTableCond.Length > 0)
        {
            detailQuery += " and " + strSubTableCond;
        }


        //strOrder = " m.accountno,m.cardprimary desc,m.cardno ";//,m.cardstatus
        if (isCorporate)
        {
            masterQuery += " and packagename = 'STTOTABLECORP'";
            masterQuery_1 += " and packagename = 'STTOTABLECORP'";

            detailQuery += " and packagename = 'STTOTABLECORP'";
            //strOrder = " m.accountcurrency,m.contractno,m.customerno,m.accountno,m.accountcurrency,m.cardaccountno,m.cardprimary,m.cardstatus,m.cardno ";
            strOrder = " m.contractno,m.accountno,m.cardaccountno,m.cardprimary,m.cardcreationdate ";
            //strOrder = " m.accountcurrency,m.contractno,m.customerno,m.accountno,m.cardaccountno,m.cardprimary,m.cardstatus,m.cardno ";
            masterQuery += " order by m.accountcurrency,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            masterQuery_1 += " order by m.accountcurrency,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }
        else
        {
            masterQuery += " order by " + strOrder;
            masterQuery_1 += " order by " + strOrder;
            //if (curBranch == 1)//
            //  masterQuery += " order by " + strOrder; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 3)//AWB || curBranch == 3
            //  masterQuery += " order by m.CARDBRANCHPART,m.accountno,m.cardprimary,m.cardstatus,m.cardno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 5)//AIB
            //  masterQuery += " order by m.CARDBRANCHPART,m.externalno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            ////masterQuery += " order by m.CARDBRANCHPART,m.cardno"; //     CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 6)//AAIB
            //  masterQuery += " order by substr(m.cardbranchpartname,6)," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            //else if (curBranch == 10)//BNP
            //  masterQuery += " order by m.cardproduct,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
            //else  //
            //  masterQuery += " order by " + strOrderBy + "m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }

        /*    if(curBranch == 5 || curBranch == 3)//AIB
              detailQuery += " order by accountno,cardno,postingdate";
            else
              detailQuery += " order by accountno,cardno,postingdate ";
        */
        detailQuery += strSupHaving;
        //detailQuery += " order by d.accountno,d.cardno,d.postingdate";
        //detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno";
        if (pBrach == 41)
            detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno,d.entryno desc";
        else if (pBrach == 36)
            detailQuery += " order by d.statementno,d.postingdate,d.docno,d.entryno";
        else if (pBrach == 23)
            detailQuery += " order by d.cardno,d.transdate";
        else
            detailQuery += " order by d.accountno,d.cardno,d.postingdate,d.docno,d.entryno";
        //detailQuery += " order by d.postingdate,d.accountno,d.cardno,d.docno,d.entryno";


        try
        {
            conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter MasterDA = new OracleDataAdapter(masterQuery, conn);
            OracleDataAdapter MasterDA_1 = new OracleDataAdapter(masterQuery_1, conn);
            OracleDataAdapter detailDA = new OracleDataAdapter(detailQuery, conn);

            conn.Open();
            //(new OracleCommand("alter session set optimizer_goal = RULE", conn)).ExecuteNonQuery();
            DSstatement = new DataSet("MasterDetailDS");
            MasterDA.Fill(DSstatement, "tStatementMasterTable");
            MasterDA_1.Fill(DSstatement, "tStatementMasterTable_1");
            detailDA.Fill(DSstatement, "tStatementDetailTable");
            //clsBasFile.SetAccessRule(@"D:\log.txt");
            //clsBasFile.writeTxtFile(@"D:\log.txt", Environment.NewLine + masterQuery, System.IO.FileMode.OpenOrCreate);
            //clsBasFile.writeTxtFile(@"D:\log.txt", Environment.NewLine + detailQuery, System.IO.FileMode.OpenOrCreate);
            if (isCorporate)
            {
                //string contractQuery = "SELECT DISTINCT x.contractno FROM tstatementmastertable x where x.branch = " + curBranch + " order by x.accountcurrency";// + " and x.contractno = 'CORP_EGP_1069140001'"
                string contractQuery = "SELECT DISTINCT x.accountcurrency,x.contractno FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and packagename = 'STTOTABLECORP' order by x.accountcurrency";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter contractDA = new OracleDataAdapter(contractQuery, conn);
                contractDA.Fill(DSstatement, "contractTable");
                // create currency table
                string currencyQuery = "SELECT DISTINCT x.accountcurrency FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and x.accountcurrency is not null and packagename = 'STTOTABLECORP'";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter currencyDA = new OracleDataAdapter(currencyQuery, conn);
                currencyDA.Fill(DSstatement, "currencyTable");
            }
            if (isCredite)
            {
                //if (!isGetTotalVal)
                //{
                try
                {
                    StatementNoDRel = DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementno], DSstatement.Tables["tStatementDetailTable"].Columns[dStatementno], false);
                    //StatementNoDRel = DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementno], DSstatement.Tables["tStatementMasterTable_1"].Columns[dStatementno], false);
                    //DSstatement.Relations.Add("StaementNoDR", DSstatement.Tables["tStatementMasterTable"].Columns[mStatementno], DSstatement.Tables["tStatementMasterTable_1"].Columns[dStatementno], false);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
                //}
            }
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
        //return DSstatement;
    }


    protected void FillStatementHistoryDataSet(int pBrach)//DataSet public static 
    {
        curBrnch = pBrach.ToString();
        masterQuery = getMasterQuery;
        detailQuery = getDetailQuery;

        getBasicData();
        detailQuery += " and d.postingdate between to_date('" + currStatementdatefrom.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') and to_date('" + currStatementdateto.ToString("dd/MM/yyyy") + "','dd/mm/yyyy') ";
        if (curProduct.Length > 0)
        {
            masterQuery += " and " + curProduct;
        }

        if (curCurrency.Length > 0)
        {
            masterQuery += " and m.cardaccountno like '%" + curCurrency + "' ";
            detailQuery += " and d.accountno like '%" + curCurrency + "' ";
        }

        if (isCorporate)
        {
            strOrder = " m.accountcurrency,m.contractno,m.customerno,m.accountno,m.cardaccountno,m.cardprimary,m.cardstatus,m.cardno ";
            masterQuery += " order by m.accountcurrency,m.CARDBRANCHPART," + strOrder; //CARDBRANCHPART,CARDBRANCHPART,;
        }
        else
        {
            masterQuery += " order by " + strOrder;
        }

        detailQuery += strSupHaving;
        detailQuery += " order by d.accountno,d.cardno,d.postingdate ";


        try
        {
            conn = new OracleConnection(clsDbCon.sConOracle);
            OracleDataAdapter MasterDA = new OracleDataAdapter(masterQuery, conn);
            OracleDataAdapter detailDA = new OracleDataAdapter(detailQuery, conn);

            conn.Open();
            DSstatementHist = new DataSet("MasterDetailDS");
            MasterDA.Fill(DSstatementHist, "tStatementMasterTable");
            detailDA.Fill(DSstatementHist, "tStatementDetailTable");
            if (isCorporate)
            {
                string contractQuery = "SELECT DISTINCT x.accountcurrency,x.contractno FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and packagename = 'STTOTABLECORP' order by x.accountcurrency";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter contractDA = new OracleDataAdapter(contractQuery, conn);
                contractDA.Fill(DSstatementHist, "contractTable");
                // create currency table
                string currencyQuery = "SELECT DISTINCT x.accountcurrency FROM " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + curBranch + " and x.accountcurrency is not null and packagename = 'STTOTABLECORP'";// + " and x.contractno = 'Corp_CB_CH000000005'" 
                OracleDataAdapter currencyDA = new OracleDataAdapter(currencyQuery, conn);
                currencyDA.Fill(DSstatementHist, "currencyTable");
            }
            if (isCredite)
            {
                try
                {
                    StatementNoDRel = DSstatementHist.Relations.Add("StaementNoDR", DSstatementHist.Tables["tStatementMasterTable"].Columns[mStatementno], DSstatementHist.Tables["tStatementDetailTable"].Columns[dStatementno], false);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
                //}
            }
            conn.Close();
        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex) 
        {
            clsBasErrors.catchError(ex);
        }
        finally
        {
        }
    }

    public clsBasStatement()
    {

    }
}

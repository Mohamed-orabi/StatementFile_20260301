using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;




public class clsStatHtmlFCMB : clsStatHtml
{
    private string strTopPicture = string.Empty;
    private string strDownPicture = string.Empty;
    private string strBankLogo = string.Empty;
    private string strLogo_FB = string.Empty;
    private string strLogo_IG = string.Empty;
    private string strLogo_LkdIn = string.Empty;
    private string strLogo_TW = string.Empty;
    private string strLogo_WA = string.Empty;
    private bool oddRow = true;
    private double balanceAtTransaction = new double();


    public clsStatHtmlFCMB()
    {
    }

    protected override void reSetupTheValues()
    {
        emailLabel = "FCMB Credit Card Statement - " + vCurDate.ToString("MMMM yyyy");
    }

    protected override void printHeader()
    {
        string newaddress1, newaddress2;
        clsMaintainData.fixAddress(curBranchVal, masterRow[mCustomeraddress1].ToString(), out newaddress1, out newaddress2);
        totPages++; endOfCustomer = "";
        extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
        if (extAccNum.Trim() == "")
            extAccNum = clsBasValid.validateStr(masterRow[accountNoName]);

        oddRow = true;
        balanceAtTransaction = Double.Parse(masterRow[mOpeningbalance].ToString());

        emailStr = new StringBuilder(""); // string builder for HTML mail body

        emailStr.Append(@"<html xmlns:v='urn:schemas-microsoft-com:vml' xmlns:o='urn:schemas-microsoft-com:office:office' lang='en'><head><meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>");
        emailStr.Append(@"<meta http-equiv='x-ua-compatible' content='ie=edge'><meta name='viewport' content='width=device-width, initial-scale=1'><meta name='x-apple-disable-message-reformatting'>");
        // Mail Title
        emailStr.Append(@"<title>FCMB Credit Card Statement - " + currStatementdateto.ToString("MM/yyyy") + @"</title>");
        // CSS Style
        emailStr.Append(@"<style type='text/css'>
            @import url('https://fonts.googleapis.com/css?family=Lato:300,400,700|Open+Sans');
            @media only screen {
              .col, td, th, div, p {font-family: 'Open Sans',-apple-system,system-ui,BlinkMacSystemFont,'Segoe UI','Roboto','Helvetica Neue',Arial,sans-serif;}
              .webfont {font-family: 'Lato',-apple-system,system-ui,BlinkMacSystemFont,'Segoe UI','Roboto','Helvetica Neue',Arial,sans-serif;}
            }

            a {text-decoration: none;}
            img {border: 0; line-height: 100%; max-width: 100%; vertical-align: middle;}
            #outlook a, .links-inherit-color a {padding: 0; color: inherit;}
            .col {font-size: 13px; line-height: 22px; vertical-align: top;}

            .hover-scale:hover {transform: scale(1.2);}
            .video {display: block; height: auto; object-fit: cover;}
            .star:hover a, .star:hover ~ .star a {color: #FFCF0F!important;}

            @media only screen and (max-width: 600px) {
              .video {width: 100%;}
              u ~ div .wrapper {min-width: 100vw;}
              .container {width: 100%!important; -webkit-text-size-adjust: 100%;}
            }

            @media only screen and (max-width: 480px) {
              .col {
                box-sizing: border-box;
                display: inline-block!important;
                line-height: 20px;
                width: 100%!important;
              }
              .col-sm-1 {max-width: 25%;}
              .col-sm-2 {max-width: 50%;}
              .col-sm-3 {max-width: 75%;}
              .col-sm-third {max-width: 33.33333%;}
              .col-sm-auto {width: auto!important;}
              .col-sm-push-1 {margin-left: 25%;}
              .col-sm-push-2 {margin-left: 50%;}
              .col-sm-push-3 {margin-left: 75%;}
              .col-sm-push-third {margin-left: 33.33333%;}

              .full-width-sm {display: table!important; width: 100%!important;}
              .stack-sm-first {display: table-header-group!important;}
              .stack-sm-last {display: table-footer-group!important;}
              .stack-sm-top {display: table-caption!important; max-width: 100%; padding-left: 0!important;}

              .toggle-content {
                max-height: 0;
                overflow: auto;
                transition: max-height .4s linear;
                -webkit-transition: max-height .4s linear;
              }
              .toggle-trigger:hover + .toggle-content,
              .toggle-content:hover {max-height: 999px!important;}

              .show-sm {
                display: inherit!important;
                font-size: inherit!important;
                line-height: inherit!important;
                max-height: none!important;
              }
              .hide-sm {display: none!important;}

              .align-sm-center {
                display: table!important;
                float: none;
                margin-left: auto!important;
                margin-right: auto!important;
              }
              .align-sm-left {float: left;}
              .align-sm-right {float: right;}

              .text-sm-center {text-align: center!important;}
              .text-sm-left {text-align: left!important;}
              .text-sm-right {text-align: right!important;}

              .nav-sm-vertical .nav-item {display: block!important;}
              .nav-sm-vertical .nav-item a {display: inline-block; padding: 5px 0!important;}

              .h1 {font-size: 32px !important;}
              .h2 {font-size: 24px !important;}
              .h3 {font-size: 16px !important;}

              .borderless-sm {border: none!important;}
              .height-sm-auto {height: auto!important;}
              .line-height-sm-0 {line-height: 0!important;}
              .overlay-sm-bg {background: #232323; background: rgba(0,0,0,0.4);}

              u ~ div .wrapper .toggle-trigger {display: none!important;}
              u ~ div .wrapper .toggle-content {max-height: none;}
              u ~ div .wrapper .nav-item {display: inline-block!important; padding: 0 10px!important;}
              u ~ div .wrapper .nav-sm-vertical .nav-item {display: block!important;}

              .p-sm-0 {padding: 0!important;}
              .p-sm-8 {padding: 8px!important;}
              .p-sm-16 {padding: 16px!important;}
              .p-sm-24 {padding: 24px!important;}
              .pt-sm-0 {padding-top: 0!important;}
              .pt-sm-8 {padding-top: 8px!important;}
              .pt-sm-16 {padding-top: 16px!important;}
              .pt-sm-24 {padding-top: 24px!important;}
              .pr-sm-0 {padding-right: 0!important;}
              .pr-sm-8 {padding-right: 8px!important;}
              .pr-sm-16 {padding-right: 16px!important;}
              .pr-sm-24 {padding-right: 24px!important;}
              .pb-sm-0 {padding-bottom: 0!important;}
              .pb-sm-8 {padding-bottom: 8px!important;}
              .pb-sm-16 {padding-bottom: 16px!important;}
              .pb-sm-24 {padding-bottom: 24px!important;}
              .pl-sm-0 {padding-left: 0!important;}
              .pl-sm-8 {padding-left: 8px!important;}
              .pl-sm-16 {padding-left: 16px!important;}
              .pl-sm-24 {padding-left: 24px!important;}
              .px-sm-0 {padding-right: 0!important; padding-left: 0!important;}
              .px-sm-8 {padding-right: 8px!important; padding-left: 8px!important;}
              .px-sm-16 {padding-right: 16px!important; padding-left: 16px!important;}
              .px-sm-24 {padding-right: 24px!important; padding-left: 24px!important;}
              .py-sm-0 {padding-top: 0!important; padding-bottom: 0!important;}
              .py-sm-8 {padding-top: 8px!important; padding-bottom: 8px!important;}
              .py-sm-16 {
            padding-top: 3px!important;
            padding-bottom: 3px!important;
        }
              .py-sm-24 {padding-top: 24px!important; padding-bottom: 24px!important;}
            }
          body,td,th {
            font-family: 'Open Sans', -apple-system, system-ui, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
        }
          a:link {
            color: #5E2785;
            text-decoration: none;
        }
          .px-sm-8 {
        }
          </style></head>");
        emailStr.Append(@"<body style='box-sizing:border-box;margin:0;padding:0;width:100%;word-break:break-word;-webkit-font-smoothing:antialiased;' link='#5E2785'>");
        emailStr.Append(@"<div style='display:none;font-size:0;line-height:0;'><!-- Add your inbox preview text here --></div>");
        emailStr.Append(@"<table class='wrapper' role='presentation' width='100%' cellspacing='0' cellpadding='0' font-size: 13px><tbody><tr>
        <td class='px-sm-16' bgcolor='#EEEEEE' align='center'><table class='container' role='presentation' width='600' cellspacing='0' cellpadding='0'><tbody><tr>");
        emailStr.Append(@"<td class='px-sm-8' style='padding: 0 24px;' align='left'><div class='spacer' style='line-height: 8px;'>‌</div>
        <table role='presentation' width='100%' cellspacing='0' cellpadding='0'><tbody><tr><td class='col' style='padding: 0 8px;' width='127'>
        <p style='color: #888888; font-size: 12px; margin: 0;'>");
        // Bank Logo
        emailStr.Append(@"<img src='cid:" + clsBasFile.getFileFromPath(strBankLogo) + @"' alt=''></p></td>");
        emailStr.Append(@"</tr></tbody></table><div class='spacer' style='line-height: 8px;'>‌</div></td></tr></tbody></table></td></tr></tbody></table>");
        emailStr.Append(@"<table class='wrapper' role='presentation' width='100%' cellspacing='0' cellpadding='0'><tbody><tr><td style='background: linear-gradient(to left, #5E2785, #A5388D);' height='300' bgcolor='#A5388D' align='center'>
        <table class='container' role='presentation' width='600' cellspacing='0' cellpadding='0' align='center'><tbody><tr><td class='py-sm-16'>
        <div class='spacer line-height-sm-0' style='line-height: 10px;'><br></div><table role='presentation' width='100%' height='380' cellspacing='0' cellpadding='0'>
        <tbody><tr><td style='padding: 0 20px;' height='380px' align='center'><h2 class='webfont h1' style='color: #FFFFFF; font-size: 52px; font-weight: 400; line-height: 100%; margin: 0;'>
        <span class='fullCenter Corbel' style='color: #5E2785; font-family: Corbel; font-weight: 700; vertical-align: top; font-size: 30px; text-align: left; line-height: 40px; text-transform: none;'>");
        // Top Banner 
        emailStr.Append(@"<img src='cid:" + clsBasFile.getFileFromPath(strTopPicture) + @"' alt='Credit Card Statement' style='border-radius: 10px;' width='640'></span>
        </h2></td></tr></tbody></table><div class='spacer line-height-sm-0' style='line-height: 10px;'><br></div></td></tr></tbody></table><!--[if gte mso 9]></div></v:textbox></v:rect><![endif]--></td></tr></tbody></table>");
        emailStr.Append(@"<table class='wrapper' role='presentation' width='100%' cellspacing='0' cellpadding='0'><tbody><tr>
        <td class='px-sm-16' bgcolor='#EEEEEE' align='center'><table class='container' role='presentation' width='640' cellspacing='0' cellpadding='0' align='center'>
        <tbody><tr><td class='px-sm-8' style='padding: 0 24px;' ali='' gn='left' width='638' bgcolor='#FFFFFF'><!--<div class='spacer line-height-sm-0 py-sm-8' style='line-height: 24px;'>‌</div>-->
        <table role='presentation' width='100%' cellspacing='0' cellpadding='0'  align='center'><tbody><tr><td class='col pb-sm-16' style='padding: 0 8px;'>
        <p><span class='webfont' style='color: #5E2785; font-size: 18px; font-weight: 700; margin: 0;'>Dear Customer,</span><br>&nbsp;");
        // Customer Name (If Requested)
        emailStr.Append(@"<br><strong>CREDIT CARD ACCOUNT STATEMENT (PRIVATE AND CONFIDENTIAL)</strong></p>");
        // Table 1 Start - Account Details
        emailStr.Append(@"<strong style='color: #5E2785;'>Account Details</strong><table cellspacing='0' cellpadding='3' font-size: 13px border='1'>");
        emailStr.Append(@"<tbody><tr bgcolor='#E9E8E8'><td bgc='' olor='#EDE3F9' width='300'>Account  name</td>");
        emailStr.Append(@"<td bgcol='' or='#EDE3F9' width='300' font-size: 18px>" + masterRow[mCustomername].ToString() + @"</td></tr>");
        emailStr.Append(@"<tr><td bgc='' olor='#EDE3F9' width='300'>Card type</td>");
        emailStr.Append(@"<td bgcol='' or='#EDE3F9' width='300'>" + masterRow[mCardproduct].ToString() + @"</td></tr>");
        emailStr.Append(@"<tr bgcolor='#E9E8E8'><td bgc='' olor='#EDE3F9' width='300'>Account  number</td>");
        emailStr.Append(@"<td bgcol='' or='#EDE3F9' width='300'>" + extAccNum + "</td></tr>");
        emailStr.Append(@"<tr><td bgc='' olor='#EDE3F9' width='300'>Credit limit</td>");
        emailStr.Append(@"<td bgcol='' or='#EDE3F9' width='300'>" + masterRow[mAccountlim].ToString() + @"</td></tr>");
        emailStr.Append(@"<tr bgcolor='#E9E8E8'><td bgc='' olor='#EDE3F9' width='300'>Repayment option</td>");
        emailStr.Append(@"<td bgcol='' or='#EDE3F9' width='300'>" + masterRow[mMinPayPercentage].ToString() + @"%</td>");
        emailStr.Append(@"<tr><td bgc='' olor='#EDE3F9' width='300'>Currency</td>");
        emailStr.Append(@"<td bgcol='' or='#EDE3F9' width='300'>" + masterRow[mAccountcurrency].ToString() + @"</td></tr></tbody></table><p></p>");
        // Table 2 Start - Payment Information
        emailStr.Append(@"<strong style='color: #5E2785;'>Payment Information For " + currStatementdateto.ToString("MMMM") + @"</strong><table cellspacing='0' cellpadding='3' font-size: 13px border='1'><tbody>");
        emailStr.Append(@"<tr bgcolor='#E9E8E8'><td width='300'>Current outstanding balance </td>");
        double closingBalance= Double.Parse(masterRow[mClosingbalance].ToString());
        double openingBalance = Double.Parse(masterRow[mOpeningbalance].ToString());
        double value;
        emailStr.Append(@"<td width='300'>" + String.Format("{0:n}",(closingBalance < 0 ? closingBalance : 0 )) + @"</td></tr>");
        emailStr.Append(@"<tr><td width='300'>Repayment due date  </td>");
        emailStr.Append(@"<td width='300'>" + ((DateTime)masterRow[mStetementduedate]).ToString("MMMM dd, yyyy") + @"</td></tr>");
        emailStr.Append(@"<tr bgcolor='#E9E8E8'><td width='300'>Minimum repayment amount due</td>");
        emailStr.Append(@"<td width='300'>" + String.Format("{0:n}",Double.TryParse(masterRow[mMindueamount].ToString(), out value) + @"</td></tr>"));
        emailStr.Append(@"<tr><td width='300'>Past due repayment amount</td>");
        emailStr.Append(@"<td width='300'>" + String.Format("{0:n}",Double.Parse(masterRow[mTotaloverdueamount].ToString())) + @"</td></tr>");
        emailStr.Append(@"<tr bgcolor='#E9E8E8'><td width='300'>Excess amount over credit limit</td>");
        emailStr.Append(@"<td width='300'>" + String.Format("{0:n}",Double.Parse(masterRow[mTotaloverdueamount].ToString())) + @"</td></tr></tbody></table><p></p>");
        // Table 3 Start - Account Summary
        emailStr.Append(@"<strong style='color: #5E2785;'>Account Summary</strong><table cellspacing='0' cellpadding='3' font-size: 13px border='1'><tbody>");
        emailStr.Append(@"<tr bgcolor='#E9E8E8'><td bgc='' olor='#EDE3F9' width='300'>Account Statement period </td>");
        emailStr.Append(@"<td bgcol='' or='#EDE3F9' width='300'>" + currStatementdatefrom.ToString("MMMM dd") + @" to " + currStatementdateto.ToString("MMMM yyyy") + @"</td></tr>");
        emailStr.Append(@"<tr bgc='' olor='#F0F8FE'><td bgc='' olor='#EDE3F9' width='300'>Outstanding  balance as at " + currStatementdatefrom.ToString("dd - MMMM") + @"</td>");
        emailStr.Append(@"<td bgcol='' or='#EDE3F9' width='300'>" + String.Format("{0:n}", (openingBalance < 0 ? openingBalance : 0)) + @"</td></tr>");
        emailStr.Append(@"<tr bgcolor='#E9E8E8'><td bgc='' olor='#EDE3F9' width='300'>Payment received</td>");
        emailStr.Append(@"<td bgcol='' or='#EDE3F9' width='300'>" + String.Format("{0:n}",Double.Parse(masterRow[mTotalpayments].ToString())) + @"</td></tr>");
        emailStr.Append(@"<tr bgc='' olor='#F0F8FE'><td bgc='' olor='#EDE3F9' width='300'>POS/ Web purchases</td>");
        emailStr.Append(@"<td bgcol='' or='#EDE3F9' width='300'>" + String.Format("{0:n}",Double.Parse(masterRow[mTotalpurchases].ToString())) + @"</td></tr>");
        emailStr.Append(@"<tr bgcolor='#E9E8E8'><td bgc='' olor='#EDE3F9' width='300'>ATM cash withdrawals</td>");
        emailStr.Append(@"<td bgcol='' or='#EDE3F9' width='300'>" + String.Format("{0:n}", Double.Parse(masterRow[mTotalcashwithdrawal].ToString())) + @"</td></tr>");
        emailStr.Append(@"<tr bgc='' olor='#F0F8FE'><td bgc='' olor='#EDE3F9' width='300'>Fees charged</td>");
        emailStr.Append(@"<td bgcol='' or='#EDE3F9' width='300'>" + String.Format("{0:n}",Double.Parse(masterRow[mTotalcharges].ToString())) + @"</td></tr>");
        emailStr.Append(@"<tr bgcolor='#E9E8E8'><td bgc='' olor='#EDE3F9' width='300'>Interest charged</td>");
        emailStr.Append(@"<td bgcol='' or='#EDE3F9' width='300'>" + String.Format("{0:n}",Double.Parse(masterRow[mTotalinterest].ToString())) + @"</td></tr>");
        emailStr.Append(@"<tr bgc='' olor='#F0F8FE'><td bgc='' olor='#EDE3F9' width='300'>Balance transfer from " + currStatementdatefrom.ToString("MMMM") + @"</td>");
        emailStr.Append(@"<td bgcol='' or='#EDE3F9' width='300'>" + String.Format("{0:n}",Double.Parse(masterRow[mTotaldebits].ToString())) + @"</td></tr>");
        emailStr.Append(@"<tr bgcolor='#E9E8E8'><td bgc='' olor='#EDE3F9' width='300'>New outstanding balance as at " + currStatementdateto.ToString("dd - MMMM") + @"</td>");
        emailStr.Append(@"<td bgcol='' or='#EDE3F9' width='300'>" + String.Format("{0:n}", (closingBalance < 0 ? closingBalance : 0)) + @"</td></tr>");
        emailStr.Append(@"<tr bgc='' olor='#F0F8FE'><td bgc='' olor='#EDE3F9' width='300'><strong>Available balance as at " + currStatementdateto.ToString("dd - MMMM") + @"</strong></td>");
        emailStr.Append(@"<td bgcol='' or='#EDE3F9' width='300'><strong>" + String.Format("{0:n}",Double.Parse(masterRow[mAccountavailablelim].ToString())) + @"</strong></td></tr></tbody></table><p></p>");
        // Detail Table (Table 4) header - Account Statement
        emailStr.Append(@"<strong style='color: #5E2785;'>Account Statement </strong><table cellspacing='0' cellpadding='3' font-size: 13px border='1'><tbody>");
        emailStr.Append(@"<tr bgcolor='#E9E8E8'>
                      <td width='300'>Trans. Date</td>");
        emailStr.Append(@"<td width='300'>Reference No.</td>");
        emailStr.Append(@"<td width='700'>Description</td>");
        emailStr.Append(@"<td width='300'>Deposit(" + masterRow[mAccountcurrency].ToString() + ")</td>");
        emailStr.Append(@"<td width='300'>Withdrawal(" + masterRow[mAccountcurrency].ToString() + ")</td>");
        emailStr.Append(@"<td width='300'>Balance</td></tr>");
        //// CR-FCMB-7869
        //emailStr.Append(@"<tr bgcolor='#E9E8E8'>
        //              <td width='350'>Trans. Date</td>");
        //emailStr.Append(@"<td width='800'>Description</td>");
        //emailStr.Append(@"<td width='350'>Deposit(" + masterRow[mAccountcurrency].ToString() + ")</td>");
        //emailStr.Append(@"<td width='350'>Withdrawal(" + masterRow[mAccountcurrency].ToString() + ")</td>");
        //emailStr.Append(@"<td width='350'>Balance</td></tr>");

        totalPages++;
        totNoOfPageStat++;
    }

    protected override void printDetail()
    {
        DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]), postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
        if (trnsDate > postingDate)
            trnsDate = postingDate;

        if (masterRow[mAccountcurrency].ToString().Trim() != detailRow[dOrigtrancurrency].ToString().Trim())
            strForeignCurr = basText.formatNumUnSign(detailRow[dOrigtranamount], "##,###,##0.00;(#,###,##0.00)", 13) + " " + basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   ");
        else
            strForeignCurr = basText.replicat(" ", 16);

        string trnsDesc = string.Empty; //= detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim()
        if (detailRow[dMerchant].ToString().Trim() == "")
            trnsDesc = detailRow[dTrandescription].ToString().Trim();
        else
            trnsDesc = detailRow[dMerchant].ToString().Trim();

        if (oddRow)
        {
            emailStr.Append(@"<tr>"); // Row bg color
            oddRow = false;
        }
        else
        {
            emailStr.Append(@"<tr bgcolor='#E9E8E8'>"); // Row bg color
            oddRow = true;
        }
        emailStr.Append(@"<td width='300'>" + ((DateTime)detailRow[dTransdate]).ToString("dd-MMM-yyyy") + @"</td>"); // Transaction Date
        emailStr.Append(@"<td width='300'>" + detailRow[dRefereneno].ToString() + @"</td>"); // Refrence No
        emailStr.Append(@"<td width='700'>" + detailRow[dTrandescription].ToString() + @"</td>"); // Description
        if (detailRow[dBilltranamountsign].ToString().Equals("CR"))
        {
            balanceAtTransaction += Double.Parse(detailRow[dBilltranamount].ToString());
            emailStr.Append(@"<td width='300'>" + String.Format("{0:n}", Double.Parse(detailRow[dBilltranamount].ToString())) + @"</td>"); // Deposit(account currency)
            emailStr.Append(@"<td width='300'></td>"); // Withdrawal(account currency)
        }
        else
        {
            balanceAtTransaction -= Double.Parse(detailRow[dBilltranamount].ToString());
            emailStr.Append(@"<td width='300'></td>"); // Deposit(account currency)
            emailStr.Append(@"<td width='300'>" + String.Format("{0:n}", Double.Parse(detailRow[dBilltranamount].ToString())) + @"</td>"); // Withdrawal(account currency)
        }
        emailStr.Append(@"<td width='300'>" + String.Format("{0:n}", balanceAtTransaction) + @"</td>"); // Balance  :: accumulative over previos transactions
        emailStr.Append(@"</tr>");

        //// CR: FCMB-7869 
        //emailStr.Append(@"<td width='350'>" + ((DateTime)detailRow[dTransdate]).ToString("dd-MMM-yyyy") + @"</td>"); // Transaction Date
        //emailStr.Append(@"<td width='800'>" + detailRow[dTrandescription].ToString() + @"</td>"); // Description
        //if (detailRow[dBilltranamountsign].ToString().Equals("CR"))
        //{
        //    balanceAtTransaction += Double.Parse(detailRow[dBilltranamount].ToString());
        //    emailStr.Append(@"<td width='350'>" + String.Format("{0:n}", Double.Parse(detailRow[dBilltranamount].ToString())) + @"</td>"); // Deposit(account currency)
        //    emailStr.Append(@"<td width='350'></td>"); // Withdrawal(account currency)
        //}
        //else
        //{
        //    balanceAtTransaction -= Double.Parse(detailRow[dBilltranamount].ToString());
        //    emailStr.Append(@"<td width='350'></td>"); // Deposit(account currency)
        //    emailStr.Append(@"<td width='350'>" + String.Format("{0:n}", Double.Parse(detailRow[dBilltranamount].ToString())) + @"</td>"); // Withdrawal(account currency)
        //}
        //emailStr.Append(@"<td width='350'>" + String.Format("{0:n}", balanceAtTransaction) + @"</td>"); // Balance  :: accumulative over previos transactions
        //emailStr.Append(@"</tr>");


        totNoOfTransactions++;
    }

    protected override void printCardFooter()
    {
        // Detail Table Closure
        emailStr.Append(@"</tbody></table>");
        // Footer Note ( Paragraph 1)
        emailStr.Append(@"<p><strong><em>Kindly note the following important information:</em></strong></p>
                  <ul>");
        emailStr.Append(@"<li><strong>Late payment fee:</strong> If your minimum repayment amount is not received on the due date stated above, a late payment fee of N2,000 monthly will apply.</li>");
        emailStr.Append(@"<li><strong>Minimum payment:</strong> If you make only the minimum payment each period, you will pay more interest and it will take you a longer period to pay off your current outstanding balance.</li></ul>");
        emailStr.Append(@"</ul><p> For enquiries, requests or complaints, kindly call our 24/7 Contact Centre on 07003290000 or send an email to <u>
        <a href='mailto:customerservice@fcmb.com'>customerservice@fcmb.com</a></u>.</p></td></tr>");
        // Main Wrapper Table Closure + Footer Paddings
        emailStr.Append(@"<tr><td class='col pb-sm-16' style='padding: 0 8px;'>&nbsp;</td></tr></tbody></table>
        <div class='spacer line-height-sm-0 py-sm-8' style='line-height: 24px;'>‌</div></td></tr></tbody></table></td></tr></tbody></table>");
        emailStr.Append(@"<table class='wrapper' role='presentation' width='100%' cellspacing='0' cellpadding='0'>
        <tbody><tr><td class='px-sm-16' bgcolor='#EEEEEE' align='center'><table class='container' role='presentation' width='600' cellspacing='0' cellpadding='0'>
        <tbody><tr><td><table role='presentation' width='100%' cellspacing='0' cellpadding='0'><tbody><tr><td class='col pb-sm-16' width='190'>
        <table class='wrapper' role='presentation' width='100%' cellspacing='0' cellpadding='0'><tbody><tr><td class='px-sm-16' height='22' bgcolor='#EEEEEE' align='center'>&nbsp;</td></tr></tbody></table>");
        emailStr.Append(@"<table class='wrapper' role='presentation' width='100%' cellspacing='0' cellpadding='0'><tbody><tr><td class='px-sm-16' bgcolor='#EEEEEE' align='center'>
        <table class='container' role='presentation' width='640' cellspacing='0' cellpadding='0'><tbody><tr><td class='px-sm-8' style='padding: 0 24px;' height='78' bgcolor='#FFFFFF' align='center'>");
        emailStr.Append(@"<span class='content-block' style='font-family: Corbel; vertical-align: top; padding-bottom: 10px; padding-top: 10px; font-size: 12px; color: #999999; text-align: center;'>");
        // Footer Disclaimer (Paragraph 2)
        emailStr.Append(@"<font color='#3a3a3a'>If you have reason to suspect any unauthorised activity on your account");
        emailStr.Append(@"<br style='font-family: Corbel;'>please contact us by sending an email to </font>");
        emailStr.Append(@"<font style='font-family: Corbel;' color='#ffb81c'>&nbsp;</font><a target='_blank' style='color: #5E2186;");
        emailStr.Append(@"text-decoration: underline; font-family: Corbel;' href='mailto:frauddesk@fcmb.com'>frauddesk@fcmb.com</a><br><br>");
        emailStr.Append(@"<font color='#5c068c'><font color='#3a3a3a'>For more information on our products and services, please call our 24/7 <br>");
        emailStr.Append(@"contact centre on</font> <font color='#5c068c'> 07003290000</font> or&nbsp;012798800 <font color='#3a3a3a'> or chat with us via <br>");
        emailStr.Append(@"Whatsapp on <font color='#5c068c'>(+234) 090 999 99814</font> <font color='#3a3a3a'> or</font> 
        <font color='#5c068c'> &nbsp;(+234) 090 999 99815.</font><font color='#efefef'><br></font>");
        emailStr.Append(@"<font color='#3a3a3a'>Alternatively send an email to</font> <a target='_blank' style='color: rgb(98, 35, 135);
        text-decoration: underline;' href='mailto:customerservice@fcmb.com'>customerservice@fcmb.com</a>.<br><br>");
        emailStr.Append(@"</font></font></span></td></tr><tr><tr><td class='px-sm-8' style='padding: 0 24px;' height='79' bgcolor='#FFFFFF' align='center'><table width='100%' cellspacing='5' cellpadding='5' border='0'><tbody>");
        // Footer Social Media Logos Block
        emailStr.Append(@"<tr><td style='color: #3A3A3A' align='center'><a href='https://www.facebook.com/fcmbmybank'>");
        emailStr.Append(@"<img src='cid:" + clsBasFile.getFileFromPath(strLogo_FB) + @"' alt='' width='52' height='52' border='0'>");
        emailStr.Append(@"<img src='cid:" + clsBasFile.getFileFromPath(strLogo_TW) + @"' alt='' width='52' height='52' border='0'>");
        emailStr.Append(@"<img src='cid:" + clsBasFile.getFileFromPath(strLogo_LkdIn) + @"' alt='' width='52' height='52' border='0'>");
        emailStr.Append(@"<img src='cid:" + clsBasFile.getFileFromPath(strLogo_IG) + @"' alt='' width='52' height='52' border='0'>");
        emailStr.Append(@"<img src='cid:" + clsBasFile.getFileFromPath(strLogo_WA) + @"' alt='' width='52' height='52' border='0'>");
        emailStr.Append(@"</a></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table>");
        // Footer Paddings
        emailStr.Append(@"<table class='wrapper' role='presentation' width='100%' cellspacing='0' cellpadding='0'><tbody><tr><td class='px-sm-16' bgcolor='#EEEEEE' align='center'>
        <table class='container' role='presentation' width='600' cellspacing='0' cellpadding='0'></table></td></tr></tbody></table>");
        emailStr.Append(@"<table class='wrapper' role='presentation' width='100%' cellspacing='0' cellpadding='0'><tbody><tr><td class='px-sm-16' bgcolor='#EEEEEE' align='center'>
        <table class='container' role='presentation' width='600' cellspacing='0' cellpadding='0'></table></td></tr></tbody></table>");
        emailStr.Append(@"<table class='wrapper' role='presentation' width='100%' cellspacing='0' cellpadding='0'><tbody><tr>
        <td class='px-sm-16' bgcolor='#EEEEEE' align='center'><table class='container' role='presentation' width='640' cellspacing='0' cellpadding='0'>
        <tbody><tr><td bgcolor='#FFFFFF' align='left'>&nbsp;</td></tr></tbody></table></td></tr></tbody></table>");
        emailStr.Append(@"<table class='wrapper' role='presentation' width='100%' cellspacing='0' cellpadding='0'><tbody><tr>
        <td class='px-sm-16' bgcolor='#EEEEEE' align='center'><table class='container' role='presentation' width='640' cellspacing='0' cellpadding='0'><tbody><tr>
        <td class='px-sm-8' style='padding: 0 24px;' bgcolor='#FFFFFF' align='left'><table role='presentation' width='100%' cellspacing='0' cellpadding='0'>
        <tbody><tr><td class='col text-sm-center' style='padding: 0 8px;' width='120' height='44' align='right'><p style='color: #888888;'>&nbsp;</p></td>");
        emailStr.Append(@"<td class='col text-sm-center' style='padding: 0 8px;' width='324' align='center'><span style='color: #888888; margin: 0;'>
        <span class='content-block' style='font-family: Corbel; vertical-align: top; padding-bottom: 10px; padding-top: 10px; font-size: 12px; color: #999999; text-align: center;'><br>");
        // Copyright statement
        emailStr.Append(@"<span style='color: #3a3a3a'>Copyright © 2020. First City Monument Bank</span></span></span></td>");
        // Footer Paddings and HTML Closures
        emailStr.Append(@"<td class='col text-sm-center' style='padding: 0 8px;' width='104' align='right'>&nbsp;</td></tr></tbody></table>
        <div class='spacer line-height-sm-0 py-sm-8' style='line-height: 48px;'>‌</div></td></tr></tbody></table></td></tr></tbody></table>");
        emailStr.Append(@"</td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></body></html>");
    }

    public string topPicture
    {
        set
        {
            strTopPicture = value;
            pLstAttachedPic.Add(strTopPicture);
        }
    }
    public string downPicture
    {
        set
        {
            strDownPicture = value;
            pLstAttachedPic.Add(strDownPicture);
        }
    }
    public string facebookLogo
    {
        set
        {
            strLogo_FB = value;
            pLstAttachedPic.Add(strLogo_FB);
        }
    }
    public string twitterLogo
    {
        set
        {
            strLogo_TW = value;
            pLstAttachedPic.Add(strLogo_TW);
        }
    }
    public string instagramLogo
    {
        set
        {
            strLogo_IG = value;
            pLstAttachedPic.Add(strLogo_IG);
        }
    }
    public string whatsappLogo
    {
        set
        {
            strLogo_WA = value;
            pLstAttachedPic.Add(strLogo_WA);
        }
    }
    public string linkinLogo
    {
        set
        {
            strLogo_LkdIn = value;
            pLstAttachedPic.Add(strLogo_LkdIn);
        }
    }

    public string bankLogo
    {
        set
        {
            strBankLogo = value;
            pLstAttachedPic.Add(strBankLogo);
        }
    }

}

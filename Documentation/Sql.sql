




--------------
--------------
--------------
delete from tstatementmastertable
delete from tstatementdetailtable

--------------
select ms.STATEMENTNO, ms.CUSTOMERNAME, ms.ACCOUNTNO, ms.CARDNO, ds.BILLTRANAMOUNT, ds.BILLTRANAMOUNTSIGN 
from tStatementMasterTable ms, tStatementdetailTable ds 
where ms.branch = seance.getbranch and 
ds.branch(+) = ms.branch and 
ds.statementno(+) = ms.STATEMENTNO 
--------------

select branch, STATEMENTNO , statementnumber, statementdatefrom , statementdateto, statementtype, statementsendtype  , stetementduedate , statementmessageline1, statementmessageline2, statementmessageline3 , contractno, contracttype, contractcreationdate  , contractstatus ,companycode, registrationnumber, contractlimit, customerno, customertitle, customername, customercreationdate, customertype, customerstatus, clientid, employer, dept, Position, employeenumber, customercountry, customerregion, customercity, customerzipcode, customeraddress1  , customeraddress2, customeraddress3, customeraddressparcode, accountno, externalno, accountcreationdate, accounttype, accountstatus, accountcurrency  , accountcountry, accountregion, accountcity, accountzipcode, accountaddress1, accountaddress2, accountaddress3, accountaddressparcode, accountlim, accountavailablelim, accountcashlim, accountavailablecashlim, cardno, cardtype, cardcreationdate, cardexpirydate, cardlastmodificationdate, cardproduct  , cardstate, cardstatus
, cardstatusdate, cardcurrency, cardbirthdate ,cardtitle ,cardvip,cardprimary ,prinarycardno ,cardbranchpart ,cardbranchpartname, cardaccountno ,cardclientname ,cardpaymentmethod ,cardlimit,cardavailablelimit ,cardcashlimit ,cardavailablecashlimit,cardcountry ,cardregion ,cardcity   , cardzipcode ,cardaddress1,cardaddress2 ,cardaddress3 ,cardaddressbarcode ,totaldueamount,mindueamount ,openingbalance ,closingbalance ,totalpurchases, totalcashwithdrawal ,totalinterest ,totalpayments ,totalcharges,totalcreditsandrefunds ,totaldebits ,totalcredits ,totalsuspends,totaloverdueamount ,totaloverlimitamount from tstatementmastertable  where BRANCH = 1 and STATEMENTNO > 6  order by accountno,cardprimary,cardno


select branch,STATEMENTNO ,statementnumber ,accountno ,cardno,transdate ,postingdate ,trandescription ,merchant ,origtranamount,origtrancurrency ,billtranamount ,billtranamountsign ,docno,refereneno from tstatementdetailtable  where BRANCH = 1 and STATEMENTNO > 6  order by accountno,cardno



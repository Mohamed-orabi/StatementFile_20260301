-- Create table
create table TSTATEMENTMASTERTABLE
(
  BRANCH                   NUMBER,
  STATEMENTNUMBER          NUMBER,
  STATEMENTDATEFROM        DATE,
  STATEMENTDATETO          DATE,
  STATEMENTTYPE            VARCHAR2(600),
  STATEMENTSENDTYPE        VARCHAR2(300),
  STETEMENTDUEDATE         DATE,
  STATEMENTMESSAGELINE1    VARCHAR2(1500),
  STATEMENTMESSAGELINE2    VARCHAR2(1500),
  STATEMENTMESSAGELINE3    VARCHAR2(1500),
  CONTRACTNO               VARCHAR2(60),
  CONTRACTTYPE             VARCHAR2(300),
  CONTRACTCREATIONDATE     DATE,
  CONTRACTSTATUS           VARCHAR2(60),
  COMPANYCODE              NUMBER,
  REGISTRATIONNUMBER       VARCHAR2(300),
  CONTRACTLIMIT            NUMBER,
  CUSTOMERNO               VARCHAR2(60),
  CUSTOMERTITLE            VARCHAR2(240),
  CUSTOMERNAME             VARCHAR2(750),
  CUSTOMERCREATIONDATE     DATE,
  CUSTOMERTYPE             VARCHAR2(300),
  CUSTOMERSTATUS           VARCHAR2(300),
  CLIENTID                 NUMBER,
  EMPLOYER                 VARCHAR2(900),
  DEPT                     VARCHAR2(300),
  POSITION                 VARCHAR2(300),
  EMPLOYEENUMBER           NUMBER,
  CUSTOMERCOUNTRY          VARCHAR2(480),
  CUSTOMERREGION           VARCHAR2(300),
  CUSTOMERCITY             VARCHAR2(300),
  CUSTOMERZIPCODE          VARCHAR2(300),
  CUSTOMERADDRESS1         VARCHAR2(150),
  CUSTOMERADDRESS2         VARCHAR2(150),
  CUSTOMERADDRESS3         VARCHAR2(150),
  CUSTOMERADDRESSPARCODE   VARCHAR2(90),
  ACCOUNTNO                VARCHAR2(120),
  EXTERNALNO               VARCHAR2(120),
  ACCOUNTCREATIONDATE      DATE,
  ACCOUNTTYPE              VARCHAR2(300),
  ACCOUNTSTATUS            VARCHAR2(300),
  ACCOUNTCURRENCY          VARCHAR2(9),
  ACCOUNTCOUNTRY           VARCHAR2(480),
  ACCOUNTREGION            VARCHAR2(300),
  ACCOUNTCITY              VARCHAR2(300),
  ACCOUNTZIPCODE           VARCHAR2(300),
  ACCOUNTADDRESS1          VARCHAR2(150),
  ACCOUNTADDRESS2          VARCHAR2(150),
  ACCOUNTADDRESS3          VARCHAR2(150),
  ACCOUNTADDRESSPARCODE    VARCHAR2(90),
  ACCOUNTLIM               NUMBER,
  ACCOUNTAVAILABLELIM      NUMBER,
  ACCOUNTCASHLIM           NUMBER,
  ACCOUNTAVAILABLECASHLIM  NUMBER,
  CARDNO                   VARCHAR2(150),
  CARDTYPE                 VARCHAR2(120),
  CARDCREATIONDATE         DATE,
  CARDEXPIRYDATE           DATE,
  CARDLASTMODIFICATIONDATE DATE,
  CARDPRODUCT              VARCHAR2(480),
  CARDSTATE                VARCHAR2(300),
  CARDSTATUS               VARCHAR2(300),
  CARDSTATUSDATE           DATE,
  CARDCURRENCY             VARCHAR2(30),
  CARDBIRTHDATE            DATE,
  CARDTITLE                VARCHAR2(450),
  CARDVIP                  VARCHAR2(6),
  CARDPRIMARY              VARCHAR2(6),
  PRINARYCARDNO            VARCHAR2(150),
  CARDBRANCHPART           NUMBER,
  CARDBRANCHPARTNAME       VARCHAR2(300),
  CARDACCOUNTNO            VARCHAR2(60),
  CARDCLIENTNAME           VARCHAR2(450),
  CARDPAYMENTMETHOD        VARCHAR2(300),
  CARDLIMIT                NUMBER,  «·„ »ﬁÏ ··«” Œœ«„ ›Ï «·ﬂ«— 
  CARDAVAILABLELIMIT       NUMBER,
  CARDCASHLIMIT            NUMBER,
  CARDAVAILABLECASHLIMIT   NUMBER,
  CARDCOUNTRY              VARCHAR2(600),
  CARDREGION               VARCHAR2(300),
  CARDCITY                 VARCHAR2(600),
  CARDZIPCODE              VARCHAR2(30),
  CARDADDRESS1             VARCHAR2(150),
  CARDADDRESS2             VARCHAR2(150),
  CARDADDRESS3             VARCHAR2(150),
  CARDADDRESSBARCODE       VARCHAR2(150),
  TOTALDUEAMOUNT           NUMBER,
  MINDUEAMOUNT             NUMBER,
  OPENINGBALANCE           NUMBER,
  CLOSINGBALANCE           NUMBER,
  TOTALPURCHASES           NUMBER,
  TOTALCASHWITHDRAWAL      NUMBER,
  TOTALINTEREST            NUMBER,
  TOTALPAYMENTS            NUMBER,
  TOTALCHARGES             NUMBER,
  TOTALCREDITSANDREFUNDS   NUMBER,
  TOTALDEBITS              NUMBER,
  TOTALCREDITS             NUMBER,
  TOTALSUSPENDS            NUMBER,
  TOTALOVERDUEAMOUNT       NUMBER,
  TOTALOVERLIMITAMOUNT     NUMBER
)
tablespace A4M
  pctfree 10
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );





select branch , statementnumber, statementdatefrom, statementdateto, statementtype, statementsendtype  , stetementduedate, statementmessageline1, statementmessageline2, statementmessageline3, contractno, contracttype, contractcreationdate  , contractstatus, companycode, registrationnumber, contractlimit, customerno, customertitle, customername, customercreationdate, customertype  , customerstatus, clientid, employer, dept, Position, employeenumber, customercountry, customerregion, customercity, customerzipcode, customeraddress1  , customeraddress2, customeraddress3, customeraddressparcode, accountno, externalno, accountcreationdate, accounttype, accountstatus, accountcurrency  , accountcountry, accountregion, accountcity, accountzipcode, accountaddress1, accountaddress2, accountaddress3, accountaddressparcode, accountlim  , accountavailablelim, accountcashlim, accountavailablecashlim, cardno, cardtype, cardcreationdate, cardexpirydate, cardlastmodificationdate, cardproduct  , cardstate, cardstatus, cardstatusdate, cardcurrency, cardbirthdate ,cardtitle ,cardvip ,cardprimary ,prinarycardno ,cardbranchpart ,cardbranchpartname   , cardaccountno ,cardclientname ,cardpaymentmethod ,cardlimit ,cardavailablelimit ,cardcashlimit ,cardavailablecashlimit ,cardcountry ,cardregion ,cardcity   , cardzipcode ,cardaddress1 ,cardaddress2 ,cardaddress3 ,cardaddressbarcode ,totaldueamount ,mindueamount ,openingbalance ,closingbalance ,totalpurchases   , totalcashwithdrawal ,totalinterest ,totalpayments ,totalcharges ,totalcreditsandrefunds ,totaldebits ,totalcredits ,totalsuspends ,totaloverdueamount ,totaloverlimitamount  
from tstatementmastertable 
where branch = 4
order by accountno,cardprimary,cardno 


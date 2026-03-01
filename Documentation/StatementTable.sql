












--------------------
--------------------
--------------------
--------------------
--------------------
--------------------
--------------------
--------------------
add Constrains:
ALTER TABLE r_h_Stmntmstr ADD CONSTRAINT pk_Stmntmstr_stmntmstrcode primary key (stmntmstrcode);
ALTER TABLE r_h_Stmntdtl ADD CONSTRAINT PK_STMNTDTL_STMNTMSTRCODE primary key (stmntmstrcode);

ALTER TABLE r_h_Stmntdtl ADD CONSTRAINT FK_STMNTDTL_STMNTMSTRCODE
FOREIGN KEY (stmntmstrcode)
REFERENCES r_h_Stmntmstr (stmntmstrcode)
ON DELETE CASCADE;

CREATE INDEX IDX_STMNTMSTR_Branch ON
R_H_STMNTMSTR (branch);

create index IDX_STMNTMSTR_BRANCHDATE on R_H_STMNTMSTR (BRANCH, STMNTTYPE, STATEMENTDATEFROM);
--------------------
-- Create table
create table r_h_Stmntdtl
(
  stmntdtlcode       NUMBER,
  stmntmstrcode      NUMBER,
  BRANCH             NUMBER,
  STATEMENTNO        NUMBER,
  STATEMENTNUMBER    NUMBER,
  CONTRACTNO         VARCHAR2(20),
  ACCOUNTNO          VARCHAR2(20),
  ACCOUNTCURRENCY    VARCHAR2(3),
  CARDNO             VARCHAR2(20),
  MBR                NUMBER,
  TRANSDATE          DATE,
  POSTINGDATE        DATE,
  TRANDESCRIPTION    VARCHAR2(250),
  MERCHANT           VARCHAR2(100),
  ORIGTRANAMOUNT     NUMBER,
  ORIGTRANCURRENCY   VARCHAR2(10),
  BILLTRANAMOUNT     NUMBER,
  BILLTRANAMOUNTSIGN VARCHAR2(2),
  DOCNO              NUMBER,
  REFERENENO         VARCHAR2(23)
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
--------------------
-- Create table
create table TSTATEMENTDETAILTABLE
(
  BRANCH             NUMBER,
  STATEMENTNO        NUMBER,
  STATEMENTNUMBER    NUMBER,
  CONTRACTNO         VARCHAR2(20),
  ACCOUNTNO          VARCHAR2(20),
  ACCOUNTCURRENCY    VARCHAR2(3),
  CARDNO             VARCHAR2(20),
  MBR                NUMBER,
  TRANSDATE          DATE,
  POSTINGDATE        DATE,
  TRANDESCRIPTION    VARCHAR2(250),
  MERCHANT           VARCHAR2(100),
  ORIGTRANAMOUNT     NUMBER,
  ORIGTRANCURRENCY   VARCHAR2(10),
  BILLTRANAMOUNT     NUMBER,
  BILLTRANAMOUNTSIGN VARCHAR2(2),
  DOCNO              NUMBER,
  REFERENENO         VARCHAR2(23)
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
-- Create/Recreate indexes 
create index ICARDNOTSTATEMENTDETAILTABLE on TSTATEMENTDETAILTABLE (CARDNO)
  tablespace A4M
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
create index ISTDETAILTABLE on TSTATEMENTDETAILTABLE (BRANCH, STATEMENTNO)
  tablespace A4M
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
--------------------
-- Create table
create table r_h_Stmntmstr
(
  stmntmstrcode            NUMBER,
  stmntType                VARCHAR2(1),-- N Normal statement C Corperate Statement 
  BRANCH                   NUMBER,
  STATEMENTNO              NUMBER,
  STATEMENTNUMBER          NUMBER,
  STATEMENTDATEFROM        DATE,
  STATEMENTDATETO          DATE,
  STATEMENTTYPE            VARCHAR2(200),
  STATEMENTSENDTYPE        VARCHAR2(100),
  STETEMENTDUEDATE         DATE,
  CONTRACTNO               VARCHAR2(20),
  CONTRACTTYPE             VARCHAR2(100),
  CONTRACTCREATIONDATE     DATE,
  CONTRACTSTATUS           VARCHAR2(100),
  COMPANYCODE              NUMBER,
  REGISTRATIONNUMBER       VARCHAR2(100),
  CONTRACTLIMIT            NUMBER,
  CUSTOMERNO               VARCHAR2(20),
  CUSTOMERCREATIONDATE     DATE,
  CUSTOMERTYPE             VARCHAR2(100),
  CUSTOMERSTATUS           VARCHAR2(100),
  CLIENTID                 NUMBER,
  ACCOUNTNO                VARCHAR2(40),
  EXTERNALNO               VARCHAR2(40),
  ACCOUNTCREATIONDATE      DATE,
  ACCOUNTTYPE              VARCHAR2(100),
  ACCOUNTSTATUS            VARCHAR2(100),
  ACCOUNTCURRENCY          VARCHAR2(3),
  ACCOUNTLIM               NUMBER,
  ACCOUNTAVAILABLELIM      NUMBER,
  ACCOUNTCASHLIM           NUMBER,
  ACCOUNTAVAILABLECASHLIM  NUMBER,
  CARDNO                   VARCHAR2(50),
  MBR                      NUMBER,
  CARDTYPE                 VARCHAR2(40),
  CARDCREATIONDATE         DATE,
  CARDEXPIRYDATE           DATE,
  CARDLASTMODIFICATIONDATE DATE,
  CARDPRODUCT              VARCHAR2(160),
  CARDSTATE                VARCHAR2(100),
  CARDSTATUS               VARCHAR2(100),
  CARDSTATUSDATE           DATE,
  CARDCURRENCY             VARCHAR2(10),
  CARDPRIMARY              VARCHAR2(2),
  PRINARYCARDNO            VARCHAR2(50),
  CARDBRANCHPART           NUMBER,
  CARDBRANCHPARTNAME       VARCHAR2(100),
  CARDACCOUNTNO            VARCHAR2(20),
  CARDCLIENTNAME           VARCHAR2(150),
  CARDPAYMENTMETHOD        VARCHAR2(100),
  CARDLIMIT                NUMBER,
  CARDAVAILABLELIMIT       NUMBER,
  CARDCASHLIMIT            NUMBER,
  CARDAVAILABLECASHLIMIT   NUMBER,
  CARDCOUNTRY              VARCHAR2(200),
  CARDREGION               VARCHAR2(100),
  CARDCITY                 VARCHAR2(200),
  CARDZIPCODE              VARCHAR2(10),
  CARDADDRESS1             VARCHAR2(50),
  CARDADDRESS2             VARCHAR2(50),
  CARDADDRESS3             VARCHAR2(50),
  CARDADDRESSBARCODE       VARCHAR2(50),
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
--------------------
-- Create table
create table TSTATEMENTMASTERTABLE
(
  BRANCH                   NUMBER,
  STATEMENTNO              NUMBER,
  STATEMENTNUMBER          NUMBER,
  STATEMENTDATEFROM        DATE,
  STATEMENTDATETO          DATE,
  STATEMENTTYPE            VARCHAR2(200),
  STATEMENTSENDTYPE        VARCHAR2(100),
  STETEMENTDUEDATE         DATE,
  STATEMENTMESSAGELINE1    VARCHAR2(500),
  STATEMENTMESSAGELINE2    VARCHAR2(500),
  STATEMENTMESSAGELINE3    VARCHAR2(500),
  CONTRACTNO               VARCHAR2(20),
  CONTRACTTYPE             VARCHAR2(100),
  CONTRACTCREATIONDATE     DATE,
  CONTRACTSTATUS           VARCHAR2(100),
  COMPANYCODE              NUMBER,
  REGISTRATIONNUMBER       VARCHAR2(100),
  CONTRACTLIMIT            NUMBER,
  CUSTOMERNO               VARCHAR2(20),
  CUSTOMERTITLE            VARCHAR2(80),
  CUSTOMERNAME             VARCHAR2(250),
  CUSTOMERCREATIONDATE     DATE,
  CUSTOMERTYPE             VARCHAR2(100),
  CUSTOMERSTATUS           VARCHAR2(100),
  CLIENTID                 NUMBER,
  EMPLOYER                 VARCHAR2(300),
  DEPT                     VARCHAR2(100),
  POSITION                 VARCHAR2(100),
  EMPLOYEENUMBER           NUMBER,
  CUSTOMERCOUNTRY          VARCHAR2(160),
  CUSTOMERREGION           VARCHAR2(100),
  CUSTOMERCITY             VARCHAR2(100),
  CUSTOMERZIPCODE          VARCHAR2(100),
  CUSTOMERADDRESS1         VARCHAR2(50),
  CUSTOMERADDRESS2         VARCHAR2(50),
  CUSTOMERADDRESS3         VARCHAR2(50),
  CUSTOMERADDRESSPARCODE   VARCHAR2(30),
  ACCOUNTNO                VARCHAR2(40),
  EXTERNALNO               VARCHAR2(40),
  ACCOUNTCREATIONDATE      DATE,
  ACCOUNTTYPE              VARCHAR2(100),
  ACCOUNTSTATUS            VARCHAR2(100),
  ACCOUNTCURRENCY          VARCHAR2(3),
  ACCOUNTCOUNTRY           VARCHAR2(160),
  ACCOUNTREGION            VARCHAR2(100),
  ACCOUNTCITY              VARCHAR2(100),
  ACCOUNTZIPCODE           VARCHAR2(100),
  ACCOUNTADDRESS1          VARCHAR2(50),
  ACCOUNTADDRESS2          VARCHAR2(50),
  ACCOUNTADDRESS3          VARCHAR2(50),
  ACCOUNTADDRESSPARCODE    VARCHAR2(30),
  ACCOUNTLIM               NUMBER,
  ACCOUNTAVAILABLELIM      NUMBER,
  ACCOUNTCASHLIM           NUMBER,
  ACCOUNTAVAILABLECASHLIM  NUMBER,
  CARDNO                   VARCHAR2(50),
  MBR                      NUMBER,
  CARDTYPE                 VARCHAR2(40),
  CARDCREATIONDATE         DATE,
  CARDEXPIRYDATE           DATE,
  CARDLASTMODIFICATIONDATE DATE,
  CARDPRODUCT              VARCHAR2(160),
  CARDSTATE                VARCHAR2(100),
  CARDSTATUS               VARCHAR2(100),
  CARDSTATUSDATE           DATE,
  CARDCURRENCY             VARCHAR2(10),
  CARDBIRTHDATE            DATE,
  CARDTITLE                VARCHAR2(150),
  CARDVIP                  VARCHAR2(2),
  CARDPRIMARY              VARCHAR2(2),
  PRINARYCARDNO            VARCHAR2(50),
  CARDBRANCHPART           NUMBER,
  CARDBRANCHPARTNAME       VARCHAR2(100),
  CARDACCOUNTNO            VARCHAR2(20),
  CARDCLIENTNAME           VARCHAR2(150),
  CARDPAYMENTMETHOD        VARCHAR2(100),
  CARDLIMIT                NUMBER,
  CARDAVAILABLELIMIT       NUMBER,
  CARDCASHLIMIT            NUMBER,
  CARDAVAILABLECASHLIMIT   NUMBER,
  CARDCOUNTRY              VARCHAR2(200),
  CARDREGION               VARCHAR2(100),
  CARDCITY                 VARCHAR2(200),
  CARDZIPCODE              VARCHAR2(10),
  CARDADDRESS1             VARCHAR2(50),
  CARDADDRESS2             VARCHAR2(50),
  CARDADDRESS3             VARCHAR2(50),
  CARDADDRESSBARCODE       VARCHAR2(50),
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
  TOTALOVERLIMITAMOUNT     NUMBER,
  HOLSTMT                  VARCHAR2(1) default 'N',
  BARCODE                  VARCHAR2(20),
  USERACTFIELD1            VARCHAR2(20),
  USERACTFIELD2            VARCHAR2(20),
  USERCUSTFIELD1           VARCHAR2(20),
  USERCUSTFIELD2           VARCHAR2(20),
  USERCARDFIELD1           VARCHAR2(20),
  USERCARDFIELD2           VARCHAR2(20)
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
-- Create/Recreate indexes 
create index ICARDNOTSTATEMENTMASTERTABLE on TSTATEMENTMASTERTABLE (CARDNO)
  tablespace A4M
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
create index ISTMASTERTABLE on TSTATEMENTMASTERTABLE (BRANCH, STATEMENTNO)
  tablespace A4M
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
--------------------





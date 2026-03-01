CREATE OR REPLACE PACKAGE BODY StToTableCorp AS
   /****************************************************************************\
     Corporate Contract Statement To Table
    ****************************************************************************
    Изменения:
    ----------------------------------------------------------------------------
    Дата        Автор            Описание
    ----------  ---------------  ---------------------------------------------
    20.04.2006  Аревкина Е.Н.    Разработка
    06.07.2006  Arevkina E.N.    For Non Card Transactions CardNo = Primary card
    11.07.2006  Arevkina E.N.    BillTranAmountSign ='DB' or 'CR'
                                 Was add field StatementNo
    12.12.2006  Arevkina E.N.    Исправлен расчет StatementNo [TWREP-634]
    16.01.2007  Gimalova A.R.    [TWREP-634]
    01.02.2007  Arevkina E.N.    исправлен расчет Externalno, Вывод TranDetail [MSCC-314]
    21.02.2007  Arevkina E.N.    добавлено поле AccountCurrency [MSCC-314]
    09.04.2007  Gimalova A.R.    Was modified definition of card limits
    22.05.2007  Arevkina E.N.    Account Available limit,  [MSCC-314]
    03.08.2007  Емельянова М.В.  Добавлена таблица tExCommon
    21.11.2007  Gimalova A.      Adding User Attributes
   \****************************************************************************/
   type tRecPan is record(Pan varchar2(20),
                          MBR  number,
                          idClient number,
                          Contract varchar2(40),
                          Status varchar2(1),
                          Cardtype varchar2(40),
                          CardCreationDate date,
                          CardExpiryDate date,
                          CardLastModificationDate date,
                          CardProduct    varchar2(160),
                          CardState      varchar2(100),
                          CardStatus     varchar2(100),
                          CardStatusdate date,
                          CardCurrency   varchar2(10),
                          CardBirthDate  date ,
                          CardTitle      varchar2(150),
                          CardVip        varchar2(2),
                          CardPrimary    varchar2(2) ,
                          PrimaryCardNo  varchar2(50),
                          CardBranchPart number,
                          CardBranchPartName varchar2(100),
                          CardAccountNo      varchar2(20),
                          CardClientName     varchar2(150),
                          CardPaymentMethod  varchar2(100),
                          CardLimit          number,
                          CardAvailableLimit number,
                          CardCashLimit      number,
                          CardAvailableCashLimit number);
                         -- );

   type tArrPan is table of tRecPan index by binary_integer;

  sPanArr  tArrPan;
  sPanCounter number;
  vPrinaryCardNo  varchar2(50);
  vPrimaryMbr     number;

  vCheckMake boolean;
  sBranch number;
  sBranchName varchar(200);
  sCounterCard number;
  vFromDate date;
  vToDate date;
  vPrizArh number;
  vNumberStr number;
  sStatementRecNumber number;
  vStatementType      varchar2(200);
  vStatementSendType  varchar2(100);
  vText1              varchar2(4000);
  vText2              varchar2(4000);
  --
  vContract           varchar(40);
  vContractType       number;
  vContractTypeName   varchar2(120);
  vCompany            number;
  vContractCreateDate date;
  vContractStatus     varchar2(100);
  vContractLim        number;
  vRegNumber          varchar2(40);
  --
  vIdClientCont       number;
  vCustomerNo         varchar2(20);
  vNameComp varchar(300);
  vCustomerCreateDate date;
  vCustomerType       varchar2(40);
  vCompanyName        varchar2(300);
  vCustomerStatus     varchar2(100);
  vOffice             varchar2(100);
  vEmployee           number;
  vPosition           varchar2(100);
  vZIPCont            varchar2(10);
  vCountryName        varchar(160);
  vRegionName         varchar(100);
  vCityName           varchar(100);
  vAddressCont        varchar(200);
  vAddressLine1       varchar2(54);
  vAddressLine2       varchar2(54);
  vAddressLine3       varchar2(54);
  vAddressBarCode     varchar2(20);
  --

  vDepositEGP varchar(40);
  vDepositUSD varchar(40);
  vDeposit varchar(40);
  --vDepositDOM varchar(40);
  --vDepositINT varchar(40);
  vChDeposit varchar(40);
  vAccountCurrency    varchar2(10);
  vIdClient  number;
  vExternal           varchar2(40);
  vAccountCreateDate  date;
  vAccountType        varchar2(100);
  vAccountStatus      varchar2(100);
  vAbbreviation varchar(10);
  vAccountLim         number;
  vAccAvailLim        number;
  vAccCashLim         number;
  vAccAvailCashLim    number;
  vChAccountLim       number;
  vChAccAvailLim      number;
  vChAccCashLim       number;
  vChAccAvailCashLim  number;
  ---------------------
  vOverdraft          number;
  vDateBalance        date;
  vTotalDueAmount     number;
  vMinDueAmount       number;
  vPrevious_balance   number;
  vNew_balance        number;
  vTotalMoney_Out     number;
  vTotalMoney_In      number;
  vTotalSuspends      number;
  vOverDueAmount      number;
  vTotalOverlimit     number;
  vTotalCreditsAndRefunds number;
  vMinimum_Payment number;
  vTotal_Payment    number;
  vWithdrawal      number;
  vPurchases       number;
  vCredit_Interest number;
  vOther           number;
  vPayment_Credits number;
  ---------------------
  /*
  vAddress1 varchar(200);
  vAddress2 varchar(200);
  vAddress3 varchar(200);
  vAddress4 varchar(200);
    */
  vStatementDate date;
  vAccount varchar(40);
  vTotalLimit number;
  vAvailable_Credit number;
  vDueDate date;
  --
  vContractClient number;
  vCardPan varchar2(50);
  vCardMBR number;
  vDocno        number;
  vTranDate     date;
  vValuedate    date;
  vDescription  tReferenceEntry.Name%TYPE;
  vTermLocation varchar2(100);
  vOrgValue     number;
  vOrgAbbr      varchar2(10);
  vTranRef      varchar2(23);
  --

  vPan varchar(50);
  vMbr number;
  vClientFio varchar2(40);
  vOpen_Balance number;
  vDebit number;
  vCredit number;
  vClose_Balance number;
  vMinPay number;
  vTotDebit number;
  vTotCredit number;
  vTotOpenBalance number;
  vTotCloseBalance number;
  vTotMinPay number;
  vMoney_Out number;
  vMoney_In number;
  vMoney_Out_Tran number;
  vMoney_In_Tran number;
  vRemain number;
  vChContract varchar(40);
  vPrizTitle number;
  --vIdent varchar(80);
  vInterest number;
  -- формирование постраничной выдачи
  vCardsCounter     number;
  vRecCounter       number;
  vPageNumber       number;
  --формирование адреса
  --vAddressCont varchar(200);
  vCityCont    tClientPersone.CityCont%TYPE;
  vRegionCont  varchar(40);
  vCountryCont number;
  signErrorDueDate boolean := false;
  sStatementId      number;
  vTotalCloseBalance number;

  sHoldStmt        varchar2(1);
  sBarCode         varchar2(20);
  sUserActField1   varchar2(20);
  sUserActField2   varchar2(20);
  sUserCardField1  varchar2(20);
  sUserCardField2  varchar2(20);
  sUserCustField1  varchar2(20);
  sUserCustField2  varchar2(20);

  PROCEDURE AddRec(pText in varchar);
  PROCEDURE GetBeginBalance;
  PROCEDURE GetOverdraft;
  Procedure GetAvailableLimit;
  PROCEDURE GetPan;
  PROCEDURE GetCardHolder;
  PROCEDURE PrintTitle;
  PROCEDURE PrintEndingTr;
  PROCEDURE PrintTransactions;
  PROCEDURE GetCompanyName;
  --PROCEDURE GetCurrency;
  PROCEDURE GetTransactions;
  PROCEDURE GetFio;
  PROCEDURE GetEntryGroup(pEntCode number,pMoney_Out number,pMoney_In number);
  PROCEDURE PrintProgon;
  PROCEDURE SetPageNumber;
  FUNCTION  PrintOpenCard return boolean;
  FUNCTION  PrintNotClosedCard return boolean;
  FUNCTION  PrintAnyCard return boolean;
  procedure GetPaymentDue;
  function  GetNumber(pStr in varchar) return boolean;
  PROCEDURE PrintAddress(pAddressCont in varchar);
  function  GetStatementSendType(pStatType in varchar) return varchar;
  procedure ClearTable;
  procedure SaveDate2MasterTable;
  procedure GetContractData(pContract in varchar);
  function  GetCardChangeState(pPan in varchar,pMbr in number) return date;
  procedure GetDailyLimit(pPan in varchar,pMbr in number);-- return number;
  function  GetPayMethod(pContract in varchar, pAccNo in varchar) return varchar;
  procedure GetAddressCard(pIdClient in number,pCardCountry out varchar,pCardRegion out varchar, pCardCity out varchar,
            pCardzipCode out varchar,pCardAddress1 out varchar,pCardAddress2 out varchar,pCardAddress3 out varchar);
  FUNCTION  GetPrimaryOpenCard return boolean;
  FUNCTION  GetPrimaryNotClosedCard return boolean;
  FUNCTION  GetPrimaryAnyCard return boolean;
  PROCEDURE GetOutstanding_Transactions;
  procedure SetWithDrawalForCard(pCardPan in varchar,pCardMBR in number,pWithdrawal in number);
  procedure SetUsageForCard(pCardPan in varchar,pCardMBR in number,pSum in number);
  function  GetTranRef(pTrancode in number) return varchar;--20.04.2006
  procedure Save2DetailTable;
  function  GetBarCode(pIdclient in number) return varchar;
  procedure GetCHOverdraft;
  function  GetStatId(pPan in varchar) return number;
  function  GetOldAccNo(pDepositAccount in varchar) return varchar;
  procedure GetChAccountLimits;
  procedure SetCardLimits(pCurrPanNo in number, pCardLim out number, pCardAvailLim out number, pCardCashLim out number, pCardAvailCashLim out number);
  procedure GetDueAmount(pAccNo in varchar,pDate in date, pDueAmount out number, pOverDueAmount out number);
  procedure GetAccUserAttributes   (pAcountNo  in varchar);
  procedure GetCardUserAttributes  (pPan in varchar, pMbr in number);
  procedure GetClientUserAttributes(pIdClient   in number);

  FUNCTION STATEMENT (pcheckmake boolean) return number is
  Begin
    s.say('---Statement',2);
    vPageNumber    := 0;
    vcheckmake     := pcheckmake;
    if vcheckmake then
      StatementDataPrepare.sNeedmake := true;
      return 0;
    end if;
    vPrinaryCardNo := '';
    vPrimaryMbr    := 0;
    vprizArh  :=0;
    vDueDate  := Null;
    signErrorDueDate := false;
    vFromDate := Statementdataprepare.sParameters.FromDate;
    vToDate   := Statementdataprepare.sParameters.ToDate;
    sStatementRecNumber := Statementdataprepare.sStatementNo;
    s.say('!!!!!!!!!!!!!!!!!!!!!!!!sStatementRecNumber '||sStatementRecNumber,2);
    if StatementDataPrepare.sParameters.UseArchives then
      vprizArh := 1;
    end if;
    if sStatementRecNumber = 1 then
      sStatementId := 1;
      ClearTable;
    else
      sStatementId := DataMember.GetNumber(Object.GetType('BRANCH'), sBranch, 'MSCCSTATEMENT_LASTID');
    end if;
    vStatementType := nvl(StatementType.GetCaptionById(Statementdataprepare.sParameters.StatementType),' ');
    vStatementSendType := GetStatementSendType(Statementdataprepare.sParameters.StatementType);
    vNumberStr := Statementdataprepare.sArrRecords.COUNT;
    vContract  := Statementdataprepare.sParameters.Contractno;
    s.say('!!!!!!!  vContract '||vContract||'   ',2);
    vContractType  := Contract.GetType(vContract);
    vContractTypeName := ContractType.GetName(vContractType);
    s.say('vContractType '||vContractTypeName,2);
    vDepositEGP := nvl(Contract.GetAccountNoByItemName(Statementdataprepare.sParameters.Contractno,'ItemDepositEGP'),' ');
    vDepositUSD := nvl(Contract.GetAccountNoByItemName(Statementdataprepare.sParameters.Contractno,'ItemDepositUSD'),' ');
    --vContract := Statementdataprepare.sParameters.Contractno;
    --vContractType := Contract.GetType(vContract);
    vText1 := ContractParams.LoadChar(ContractParams.cContracttype,vContractType,'PromText',false);
    vText2 := StatementDataPrepare.sParameters.PromotionalText;
    vStatementDate := vToDate;
    GetContractData(vContract);
    for nn in 1..2 loop
      vPayment_Credits := 0;
      vWithdrawal := 0;
      vPurchases  := 0;
      vOther      := 0;
      vPayment_Credits := 0;
      vTotDebit  := 0;
      vTotCredit := 0;
      vTotal_Payment := 0;
      vTotOpenBalance  := 0;
      vTotCloseBalance := 0;
      vCredit_Interest := 0;
      vTotMinPay       := 0;
      vTotalOverlimit := 0;
      if nn = 1 then
        vPrizTitle := 0;
        vDeposit := vDepositEGP;
      else
        vPrizTitle := 0;
        vDeposit := vDepositUSD;
      end if;
      if vDeposit != ' ' then
        GetCompanyName;
        vAccount := vDeposit;
        GetOverdraft;
        GetAvailableLimit;
        --GetCurrency;
        GetCardHolder;
        PrintEndingTr;
        --SaveDate2MasterTable;
      end if;
    end loop;
    SetPageNumber;
    return 0;
    exception when others then
    s.say('error sTatement '||sqlcode,1);
    return sqlcode;
  End;
function GetStatementSendType(pStatType in varchar) return varchar
  is
  vSendType varchar2(100);
  begin
    for i in (select * from tstatementtype where branch = sBranch and id = pStatType) loop
      vSendType := StatementPeriodRT.GetSendTypeText(i.sendtype);
    end loop;
    return vSendType;
  end;
procedure ClearTable
 is
 begin
   delete from tStatementMasterTable where branch = sBranch;
   delete from tStatementDetailTable where branch = sBranch;
   exception
   when others then
     s.say('StToTableContr.ClearTable Error:'||sqlErrm);
     null;
 end;
procedure GetContractData(pContract in varchar)
  is
  vStatus varchar2(40) := null;

  begin
    vContractCreateDate := null;
    vContractStatus     := null;
    vContractLim        := 0;
    begin
      select c.createdate, c.status
        into vContractCreateDate,vStatus
        from tcontract c
       where c.branch = sBranch
         and c.no = pContract
         and rownum = 1;
    exception
     when NO_DATA_FOUND then
       null;
    end;
    vContractStatus := Contract.GetStatusName(vStatus);
    exception
     when OTHERS then
       s.say('StToTableCorp.GetContractData Error:'||sqlErrm);
       null;
  end;
  PROCEDURE AddRec(pText in varchar) is
  Begin
    vNumberStr := vNumberStr + 1;
    vRecCounter := vRecCounter + 1;
    Statementdataprepare.sArrRecords(vNumberStr) := SUBSTR(pText, 1, 998) || CHR(13)||CHR(10);
  End;
  PROCEDURE SetPageNumber
  is
  vCurrNumber number := 0;
  vCurrRec varchar2(20000) := null;
  vCurrPageNo number := 0;
  begin

    for i in 0..vPageNumber-1 loop
     -- s.say(to_char( i*60 + 12),2);
      vCurrRec  := ltrim(rtrim(Statementdataprepare.sArrRecords( i*60 + 18)));
      vCurrRec  := translate( vCurrRec, chr(13)||chr(10), ' ');
      vCurrPageNo := to_number(ltrim(rtrim(vCurrRec)));
      vCurrRec  := vCurrRec||'of '||to_char(vPageNumber);
      if vCurrPageNo = vPageNumber then
        vCurrRec := vCurrRec||' END ';
      end if;
      Statementdataprepare.sArrRecords(vCurrNumber*60 + 18) := vCurrRec || CHR(13)||CHR(10);
      vCurrNumber := vCurrNumber + 1;
    end loop;
  end;
function GetBarCode(pIdclient in number) return varchar is
   vBarCode  varchar2(20);
   vPropList   APITypes.TypeUserPropList;
  begin
   vPropList := Client.GetUserAttributesList(pIdClient);
   for i in 1..vPropList.Count loop
     if vPropList(i).ID = 'BARCODE' then
       vBarCode := vPropList(i).ValueStr;
       exit;
     end if;
   end loop;
   return  vBarCode;
  end;
procedure GetAccUserAttributes(pAcountNo in varchar)
  is
  vPropList   APITypes.TypeUserPropList;
  begin
    sHoldStmt      := null;
    sBarCode       := null;
    sUserActField1 := null;
    sUserActField2 := null;
    vPropList := Account.GetUserAttributesList(pAcountNo);
    for i in 1..vPropList.Count loop
      if vPropList(i).ID = 'HOLDSTMT' then
        sHoldStmt := vPropList(i).ValueStr;
      elsif vPropList(i).ID = 'BARCODE' then
        sBarCode := vPropList(i).ValueStr;
      elsif vPropList(i).ID = 'ACTFIELD1' then
        sUserActField1:= vPropList(i).ValueStr;
      elsif vPropList(i).ID = 'ACTFIELD2' then
        sUserActField2:= vPropList(i).ValueStr;
      end if;
    end loop;
    sHoldStmt := nvl(sHoldStmt, 'N');
    exception
      when NO_DATA_FOUND then
        null;
  end;
procedure GetCardUserAttributes(pPan in varchar, pMbr in number)
  is
  vPropList   APITypes.TypeUserPropList;
  begin
    sUserCardField1 := null;
    sUserCardField2 := null;
    vPropList := Card.GetUserAttributesList(pPan, pMbr);
    for i in 1..vPropList.Count loop
      if vPropList(i).ID = 'CARDFIELD1' then
        sUserCardField1 := vPropList(i).ValueStr;
      elsif vPropList(i).ID = 'CARDFIELD2' then
        sUserCardField2 := vPropList(i).ValueStr;
      end if;
    end loop;
     exception
      when others then
        null;
  end;
procedure GetClientUserAttributes(pIdClient in number)
  is
  vPropList   APITypes.TypeUserPropList;
  begin
    sUserCustField1 := null;
    sUserCustField2 := null;
    vPropList := Client.GetUserAttributesList(pIdClient);
    for i in 1..vPropList.Count loop
      if vPropList(i).ID = 'CUSTFIELD1' then
        sUserCustField1 := vPropList(i).ValueStr;
      elsif vPropList(i).ID = 'CUSTFIELD2' then
        sUserCustField2 := vPropList(i).ValueStr;
      end if;
    end loop;
    exception
      when others then
        null;
  end;

procedure GetOverdraft
  is
  Begin
    vTotalLimit := 0;
    Select Overdraft
      into vTotalLimit
      from tAccount
      where branch  = sBranch and
        accountno = vDeposit and
        rownum = 1;
      exception
      when NO_DATA_FOUND then
      null;
    vTotalLimit := nvl(vTotalLimit,0);
  End;
procedure GetAvailableLimit
  is
  Begin
    vRemain  := 0;
    vAvailable_Credit := 0;
    Select
      sum(T.Remain) - sum(T.value)
    into vRemain
    from
      (
      select
        a.remain,
        0 value
        from tAccount a
        where
           a.branch = sBranch and
           a.Accountno = vDeposit
      union all
      select
        0 Remain,
        -e.Value
      from tEntry e, tDocument d
      where e.Branch   = sBranch and
        e.DebitAccount = vDeposit and
        d.Branch       = e.Branch and
        d.DocNo        = e.DocNo and
        d.OpDate      > vToDate and
        d.NewDocNo is null
      union all
      select
        0 Remain,
        e.Value
      from tEntry e, tDocument d
      where e.Branch    = sBranch and
        e.CreditAccount = vDeposit and
        d.Branch        = e.Branch and
        d.DocNo         = e.DocNo and
        d.OpDate       > vToDate and
        d.NewDocNo is null
      union all
      select
        0 Remain,
        -e.Value
      from tEntryarc e, tDocumentarc d
      where e.Branch   = sBranch and
        e.DebitAccount = vDeposit and
        d.Branch       = e.Branch and
        d.DocNo        = e.DocNo and
        d.OpDate      > vToDate and
        d.NewDocNo is null and
        vPrizArh = 1
      union all
      select
        0 Remain,
        e.Value
      from tEntryarc e, tDocumentarc d
      where e.Branch    = sBranch and
        e.CreditAccount = vDeposit and
        d.Branch        = e.Branch and
        d.DocNo         = e.DocNo and
        d.OpDate       > vToDate and
        d.NewDocNo is null and
        vPrizArh = 1
      ) T;
    vRemain := nvl(vRemain,0);
    if vRemain < 0 then
      vAvailable_Credit := vTotalLimit - Abs(vRemain);
    else
      vAvailable_Credit := vTotalLimit;
    end if;
    s.say('!!!!!!!!!!!              vRemain  '||vRemain||'  vTotalLimit '||vTotalLimit||'  vAvailable_Credit '||vAvailable_Credit||' '||vTotCloseBalance,2);
    --vAccAvailLim
  End;
function GetOldAccNo(pDepositAccount in varchar) return varchar
  is
  vOldAccount tACC2ACC.OldAccountno%TYPE := null;
  begin
    for i in (select * from tacc2acc where branch = sBranch and newaccountno = pDepositAccount )loop
      vOldAccount := i.OldAccountNo;
      exit;
    end loop;
    return vOldAccount;
  exception
    when others then
      s.say('StToTableContr.GetOldAccNo Error:'||sqlErrm);
      return null;
  end;
procedure GetCompanyName
  is
   -- vAddressCont varchar(200) := ' ';
    vCountryCont number := 0;
    vRegionCont varchar(200) := ' ';
    vCityCont number := '0';
    vAccType   number;
    vAccStat   varchar2(2);
  Begin
    vAddressCont := '';
    vAddressBarCode := '';
    vZIPCont     := '';
    vOffice      := null;
    vEmployee    := null;
    vPosition    := null;
    vIdClientCont := -1;
    vExternal           := null;
    vAbbreviation := '';
    vRegNumber := '';
    begin
      select
        cl.companycode,cl.Name,cl.AddressCont,cl.CountryCont,cl.RegionCont,cl.CityCont,
        cl.datecreate, 'Corporate Customer' Type, a.idclient,a.AccountType,
        cl.regno,cl.ZIPCont, r.Abbreviation,
        a.Acct_Stat, a.createdate
      into vCompany,vNameComp,vAddressCont,vCountryCont,vRegionCont,vCityCont,
        vCustomerCreateDate,vCustomerType,  vIdClientCont,vAccType,
        vRegNumber,vZIPCont,vAbbreviation,
        vAccStat,vAccountCreateDate
      from taccount a,tclientBank cl,treferencecurrency r
      WHERE a.branch = sbranch and
      a.accountno = vDeposit and
      cl.Branch = a.Branch and
      cl.Idclient = a.Idclient and
      r.Branch(+) = a.branch and
      r.Currency(+) = a.Currencyno and
      rownum = 1;
      exception
      when NO_DATA_FOUND then
      null;
    end;
    vAccountTYpe := AccountType.GetName(vAccType);
    vAccountStatus := ReferenceAcct_stat.GetName(vAccStat);
    vNameComp    := nvl(vNameComp,' ');
    vAddressCont := nvl(vAddressCont,' ');
    vAddressBarCode := GetBarCode(vIdClientCont);

    vCountryCont := nvl(vCountryCont,0);
    vRegionCont  := nvl(vRegionCont,' ');
    vCityCont    := nvl(vCityCont,'0');
    vIdClientCont := nvl(vIdClientCont,-1);
    vCompany := nvl(vCompany,0);
    vCustomerNo := vRegnumber;
    --
    vExternal := GetOldAccNo(vDeposit);
    GetAddressCard(vIdClientCont,vCountryName,vRegionName, vCityName,
    vZIPCont,vAddressLine1,vAddressLine2,vAddressLine3);
  End;
function GetEndBalance(pCardholderDeposit in varchar, pToDate in date) return number
  is
  vCHEndBalance number := 0;
  Begin
    select
      sum(T.RemDep) - sum(T.ValDep)
    into vCHEndBalance
    from
      (
      select
        a.Remain RemDep,
        0        ValDep
      from tAccount a
      where a.Branch    = sBranch and
        a.AccountNo = pCardholderDeposit
      union all
      select
        0 RemDep,
        -e.Value ValDep
      from tEntry e, tDocument d
      where e.Branch   = sBranch and
        e.DebitAccount = pCardholderDeposit and
        d.Branch       = e.Branch and
        d.DocNo        = e.DocNo and
        d.OpDate       > pToDate and
        d.NewDocNo is null
      union all
      select
        0 RemDep,
        e.Value ValDep
      from tEntry e, tDocument d
      where e.Branch    = sBranch and
        e.CreditAccount = pCardholderDeposit and
        d.Branch        = e.Branch and
        d.DocNo         = e.DocNo and
        d.OpDate        > pToDate and
        d.NewDocNo is null
      union all
      select
        0 RemDep,
        -e.Value ValDep
      from tEntryarc e, tDocumentarc d
      where e.Branch   = sBranch and
        e.DebitAccount = pCardholderDeposit and
        d.Branch       = e.Branch and
        d.DocNo        = e.DocNo and
        d.OpDate       > pToDate and
        d.NewDocNo is null and
        vprizArh = 1
      union all
      select
        0 RemDep,
        e.Value ValDep
      from tEntryarc e, tDocumentarc d
      where e.Branch    = sBranch and
        e.CreditAccount = pCardholderDeposit and
        d.Branch        = e.Branch and
        d.DocNo         = e.DocNo and
        d.OpDate        > pToDate and
        d.NewDocNo is null and
        vprizArh = 1
      ) T;
    vCHEndBalance := nvl(vCHEndBalance,0);
    return vCHEndBalance;
    exception
     when OTHERS then
       s.say('StToTableCorp. Error:'||sqlErrm);
       return 0;
  End;

procedure GetBeginBalance
  is
  Begin
    select
      sum(T.RemDep) - sum(T.ValDep)
    into vRemain
    from
      (
      select
        a.Remain RemDep,
        0        ValDep
      from tAccount a
      where a.Branch    = sBranch and
        a.AccountNo = vChDeposit
      union all
      select
        0 RemDep,
        -e.Value ValDep
      from tEntry e, tDocument d
      where e.Branch   = sBranch and
        e.DebitAccount = vChDeposit and
        d.Branch       = e.Branch and
        d.DocNo        = e.DocNo and
        d.OpDate      >= vFromDate and
        d.NewDocNo is null
      union all
      select
        0 RemDep,
        e.Value ValDep
      from tEntry e, tDocument d
      where e.Branch    = sBranch and
        e.CreditAccount = vChDeposit and
        d.Branch        = e.Branch and
        d.DocNo         = e.DocNo and
        d.OpDate       >= vFromDate and
        d.NewDocNo is null
      union all
      select
        0 RemDep,
        -e.Value ValDep
      from tEntryarc e, tDocumentarc d
      where e.Branch   = sBranch and
        e.DebitAccount = vChDeposit and
        d.Branch       = e.Branch and
        d.DocNo        = e.DocNo and
        d.OpDate      >= vFromDate and
        d.NewDocNo is null and
        vprizArh = 1
      union all
      select
        0 RemDep,
        e.Value ValDep
      from tEntryarc e, tDocumentarc d
      where e.Branch    = sBranch and
        e.CreditAccount = vChDeposit and
        d.Branch        = e.Branch and
        d.DocNo         = e.DocNo and
        d.OpDate       >= vFromDate and
        d.NewDocNo is null and
        vprizArh = 1
      ) T;
    vRemain := nvl(vRemain,0);
  End;
function GetTranRef(pTrancode in number) return varchar is --20.04.2006
   vTranARN  varchar2(40);
   vCBList   varchar2(100);
  begin
    vTranARN := ' ';
    vCBList := '+15+16+17+05+06+07+25+26+27+35+36+37+';   --CB
    begin
      select
        decode(instr(vCBList,'+'||ltrim(rtrim(to_char(v.tc,'00')))||'+'),0,decode(ltrim(rtrim(to_char(v.tc,'00'))),'40',substr(v.tcr0,36,23),substr(v.tcr0,20,23)),substr(v.tcr0,23,23))  arn
        into vTranARN
        from texvisatcr v
       where v.branch = sbranch and
         v.code = pTranCode;
    exception
     when NO_DATA_FOUND then
       null;
    end;
    return nvl(vTranARN,' ');
  end;
  procedure GetTransactions is
  Begin
  vTotalCreditsAndRefunds := 0;
  vWithdrawal      := 0;
  vPurchases       := 0;
  vCredit_Interest := 0;
  vOther           := 0;
  vPayment_Credits := 0;
    for i in (
      select
      Trans.Opdate,
      Trans.ValueDate,
      Trans.Docno,
      Trans.No,
      --nvl(c.Pan,vPan) Pan,
      Trans.Merchant,
      Trans.Description,
      Trans.Ident,
      Trans.DebitEntCode,
      decode(substr(Trans.Ident,1,5),'OPCOM',' ',
        decode( instr(Trans.Ident,'FEE'),0,Trans.Termlocation,' ')) Termlocation,
      Trans.Detail,
      Trans.OrgCurrency,
      Trans.AeDocno,
      Trans.eValueDate,
      Trans.TranCode,
      c.Pan CardPan,
      c.mbr CardMBR,
      sum(Trans.OrgValue) OrgValue,
      sum(Trans.Money_Out) Money_Out,
      sum(Trans.Money_In)  Money_In
      from(
      select
      d.Opdate,
      e.ValueDate,
      e.Docno,
      e.No,
      Re.Name Description,
      Re.Ident,
      e.debitentcode,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null,
      decode(ex.branch, null, ' ',
                             ex.TermLocation),
                             v.TermLocation),
                             p.TermLocation),
                             ae.TermLocation) Termlocation,
      decode( Ae.branch,null,
      decode( P.branch,null,
      decode( V.branch, null,
      decode(ex.branch, null, 0,
                             ex.Code),
                             V.TranCode),
                             P.TranCode),
                             Ae.TranCode) TranCode,
      decode(p.branch,null,
      decode(v.branch, null,
      decode(ex.branch, null, ' ',
                             decode(ex.device,'A',' ',null,' ',ex.Retailer)),
                             v.Retailer),
                             p.Retailer) Merchant,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null,
      decode(ex.branch, null, ' ',
                             ex.Pan),
                             v.Pan),
                             p.Pan),
                             ae.Pan) Pan,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null,
      decode(ex.branch, null, 0,
                             ex.MBR),
                             v.MBR),
                             p.MBR),
                             ae.MBR) MBR,
      re.name Detail,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null,
      decode(ex.branch, null, 0,
                             ex.OrgCurrency),
                             V.OrgCurrency),
                             P.OrgCurrency),
                             Ae.OrgCurrency) OrgCurrency,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null,
      decode(ex.branch, null, 0,
                             ex.OrgAmount),
                             V.OrgValue),
                             P.OrgValue),
                             Ae.OrgValue) OrgValue,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null,
      decode(ex.branch, null, 0,
                             ex.Docno),
                             V.Docno),
                             P.Docno),
                             Ae.Docno) AeDocno,
      e.valuedate eValuedate,
      e.value Money_Out,
      0       Money_In
      from tEntry e,tReferenceEntry re,tDocument d,tPoscheque p,tVoiceSlip v,tAtmExt ae, tExCommon ex
      where e.Branch       = sbranch and
            e.Debitaccount = vchDeposit and
            e.Value       != 0 and
            re.branch = e.branch and
            re.code = e.debitentcode and
            d.branch       = e.Branch and
            d.docno        = e.docno and
            d.NewDocNo     is null and
            d.OpDate       between vFromDate and vToDate and
            p.Branch(+)    = d.branch and
            p.DocNo(+)     = d.Docno  and
            v.Branch(+)    = d.branch and
            v.DocNo(+)     = d.Docno and
            ae.Branch(+)   = d.branch and
            ae.DocNo(+)    = d.Docno and
            ex.Branch(+)   = d.branch and
            ex.DocNo(+)    = d.Docno
      union all
      select
      d.Opdate,
      e.ValueDate,
      e.Docno,
      e.No,
      Re.Name Description,
      Re.Ident,
      e.debitentcode,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null,
      decode(ex.branch, null, ' ',
                             ex.TermLocation),
                             v.TermLocation),
                             p.TermLocation),
                             ae.TermLocation) Termlocation,
      decode( Ae.branch,null,
      decode( P.branch,null,
      decode( V.branch, null,
      decode(ex.branch, null, 0,
                             ex.Code),
                             V.TranCode),
                             P.TranCode),
                             Ae.TranCode) TranCode,
      decode(p.branch,null,
      decode(v.branch, null,
      decode(ex.branch, null, ' ',
                             decode(ex.device,'A',' ',null,' ',ex.Retailer)),
                             v.Retailer),
                             p.Retailer) Merchant,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null,
      decode(ex.branch, null, ' ',
                             ex.Pan),
                             v.Pan),
                             p.Pan),
                             ae.Pan) Pan,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null,
      decode(ex.branch, null, 0,
                             ex.MBR),
                             v.MBR),
                             p.MBR),
                             ae.MBR) MBR,
      re.name Detail,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null,
      decode(ex.branch, null, 0,
                             ex.OrgCurrency),
                             V.OrgCurrency),
                             P.OrgCurrency),
                             Ae.OrgCurrency) OrgCurrency,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null,
      decode(ex.branch, null, 0,
                             ex.OrgAmount),
                             V.OrgValue),
                             P.OrgValue),
                             Ae.OrgValue) OrgValue,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null,
      decode(ex.branch, null, 0,
                             ex.Docno),
                             V.Docno),
                             P.Docno),
                             Ae.Docno) AeDocno,
      e.valuedate eValuedate,
      0       Money_Out,
      e.value Money_In
      from tEntry e,tReferenceEntry re,tDocument d,tPoscheque p,tVoiceSlip v,tAtmExt ae, tExCommon ex
      where e.Branch       = sbranch and
            e.Creditaccount = vchDeposit and
            e.Value       != 0 and
            re.branch = e.branch and
            re.code = e.debitentcode and
            d.branch       = e.Branch and
            d.docno        = e.docno and
            d.NewDocNo     is null and
            d.OpDate       between vFromDate and vToDate and
            p.Branch(+)    = d.branch and
            p.DocNo(+)     = d.Docno  and
            v.Branch(+)    = d.branch and
            v.DocNo(+)     = d.Docno and
            ae.Branch(+)   = d.branch and
            ae.DocNo(+)    = d.Docno and
            ex.Branch(+)   = d.branch and
            ex.DocNo(+)    = d.Docno
      union all
      select
      d.Opdate,
      e.ValueDate,
      e.Docno,
      e.No,
      Re.Name Description,
      re.ident,
      e.debitentcode,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null,
      decode(ex.branch, null, ' ',
                             ex.TermLocation),
                             v.TermLocation),
                             p.TermLocation),
                             ae.TermLocation) Termlocation,
      decode( Ae.branch,null,
      decode( P.branch,null,
      decode( V.branch, null,
      decode(ex.branch, null, 0,
                             ex.Code),
                             V.TranCode),
                             P.TranCode),
                             Ae.TranCode) TranCode,
      decode(p.branch,null,
      decode(v.branch, null,
      decode(ex.branch, null, ' ',
                             decode(ex.device,'A',' ',null,' ',ex.Retailer)),
                             v.Retailer),
                             p.Retailer) Merchant,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null,
      decode(ex.branch, null, ' ',
                             ex.Pan),
                             v.Pan),
                             p.Pan),
                             ae.Pan) Pan,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null,
      decode(ex.branch, null, 0,
                             ex.MBR),
                             v.MBR),
                             p.MBR),
                             ae.MBR) MBR,
      re.name Detail,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null,
      decode(ex.branch, null, 0,
                             ex.OrgCurrency),
                             V.OrgCurrency),
                             P.OrgCurrency),
                             Ae.OrgCurrency) OrgCurrency,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null,
      decode(ex.branch, null, 0,
                             ex.OrgAmount),
                             V.OrgValue),
                             P.OrgValue),
                             Ae.OrgValue) OrgValue,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null,
      decode(ex.branch, null, 0,
                             ex.Docno),
                             V.Docno),
                             P.Docno),
                             Ae.Docno) AeDocno,
      e.valuedate eValuedate,
      e.value Money_Out,
      0       Money_In
      from tEntryarc e,tReferenceEntry re,tDocumentarc d,tPoschequearc p,
           tVoiceSliparc v,tAtmExtarc ae, tExCommon ex
      where e.Branch       = sbranch and
            e.Debitaccount = vchDeposit and
            e.Value       != 0 and
            re.branch = e.branch and
            re.code = e.debitentcode and
            d.branch       = e.Branch and
            d.docno        = e.docno and
            d.NewDocNo     is null and
            d.OpDate       between vFromDate and vToDate and
            p.Branch(+)    = d.branch and
            p.DocNo(+)     = d.Docno  and
            v.Branch(+)    = d.branch and
            v.DocNo(+)     = d.Docno and
            ae.Branch(+)   = d.branch and
            ae.DocNo(+)    = d.Docno and
            ex.Branch(+)   = d.branch and
            ex.DocNo(+)    = d.Docno and
            vprizArh = 1
      union all
      select
      d.Opdate,
      e.ValueDate,
      e.Docno,
      e.No,
      Re.Name Description,
      re.ident,
      e.debitentcode,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null,
      decode(ex.branch, null, ' ',
                             ex.TermLocation),
                             v.TermLocation),
                             p.TermLocation),
                             ae.TermLocation) Termlocation,
      decode( Ae.branch,null,
      decode( P.branch,null,
      decode( V.branch, null,
      decode(ex.branch, null, 0,
                             ex.Code),
                             V.TranCode),
                             P.TranCode),
                             Ae.TranCode) TranCode,
      decode(p.branch,null,
      decode(v.branch,null,
      decode(ex.branch, null, ' ',
                             decode(ex.device,'A',' ',null,' ',ex.Retailer)),
                             v.Retailer),
                             p.Retailer) Merchant,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null,
      decode(ex.branch, null, ' ',
                             ex.Pan),
                             v.Pan),
                             p.Pan),
                             ae.Pan) Pan,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null,
      decode(ex.branch, null, 0,
                             ex.MBR),
                             v.MBR),
                             p.MBR),
                             ae.MBR) MBR,
      re.name Detail,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null,
      decode(ex.branch, null, 0,
                             ex.OrgCurrency),
                             V.OrgCurrency),
                             P.OrgCurrency),
                             Ae.OrgCurrency) OrgCurrency,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null,
      decode(ex.branch, null, 0,
                             ex.OrgAmount),
                             V.OrgValue),
                             P.OrgValue),
                             Ae.OrgValue) OrgValue,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null,
      decode(ex.branch, null, 0,
                             ex.Docno),
                             V.Docno),
                             P.Docno),
                             Ae.Docno) AeDocno,
      e.valuedate eValuedate,
      0       Money_Out,
      e.value Money_In
      from tEntryarc e,tReferenceEntry re,tDocumentarc d,tPoschequearc p,
           tVoiceSliparc v,tAtmExtarc ae, tExCommon ex
      where e.Branch       = sbranch and
            e.Creditaccount = vchDeposit and
            e.Value       != 0 and
            re.branch = e.branch and
            re.code = e.debitentcode and
            d.branch       = e.Branch and
            d.docno        = e.docno and
            d.NewDocNo     is null and
            d.OpDate       between vFromDate and vToDate and
            p.Branch(+)    = d.branch and
            p.DocNo(+)     = d.Docno  and
            v.Branch(+)    = d.branch and
            v.DocNo(+)     = d.Docno and
            ae.Branch(+)   = d.branch and
            ae.DocNo(+)    = d.Docno and
            ex.Branch(+)   = d.branch and
            ex.DocNo(+)    = d.Docno and
            vprizArh = 1
      )Trans, tcard c
      where c.branch(+) = sbranch and
      c.pan(+) = Trans.pan and
      c.MBR(+) = Trans.mbr
      group by
      Trans.Opdate,
      Trans.ValueDate,
      Trans.Docno,
      Trans.No,
      Trans.Description,
      Trans.TranCode,
      Trans.Ident,
      Trans.Debitentcode,
      decode(substr(Trans.Ident,1,5),'OPCOM',' ',
        decode( instr(Trans.Ident,'FEE'),0,Trans.Termlocation,' ')),
      Trans.Merchant,
      c.Pan,
      c.mbr,
      Trans.Detail,
      Trans.OrgCurrency,
      Trans.AeDocno,
      Trans.eValueDate
      order by 1 desc,2 desc,3,4,5,6) loop
      vCardPan      := i.Cardpan;
      vCardMBR      := i.CardMBR;
      vDocno        := i.docno;
      vTranDate     := trunc(i.ValueDate);
      vValuedate    := i.Opdate;
      vDescription  := nvl(i.Description,' ');
      vTermLocation := nvl(i.TermLocation,' ');
      begin
        vOrgAbbr := ' ';
        select Abbreviation
          into vOrgAbbr
          from tReferenceCurrency
          where branch  = sBranch and
            currency = i.OrgCurrency and
            rownum = 1;
            exception
            when NO_DATA_FOUND then
            null;
        vOrgAbbr  := nvl(vOrgAbbr,' ');
      end;
      if instr(i.Ident,'FEE') != 0 or substr(i.Ident,1,5) = 'OPCOM' then
        vOrgValue     := null;
        vOrgAbbr      := ' ';
      else
        vOrgValue     := i.OrgValue;
      end if;
      vTranRef      := GetTranRef(i.Trancode);
      vMoney_Out    := vMoney_Out + i.Money_Out;
      vMoney_In     := vMoney_In + i.Money_In;
      vMoney_Out_Tran    := i.Money_Out;
      vMoney_In_Tran     := i.Money_In;
      if (substr(i.ident,1,21) != 'CHARGE_INTEREST_GROUP') then
        GetEntryGroup(i.debitentcode,i.Money_Out,i.Money_In);
      else
        vCredit_Interest := vCredit_Interest + i.Money_Out - i.Money_In;
      end if;
      SetUsageForCard(nvl(i.CardPan,' '),nvl(i.CardMBR,0),vMoney_Out_Tran - vMoney_In_Tran);
      Save2DetailTable;
    end loop;
  End;
  procedure Save2DetailTable
   is
   vSing         varchar2(2);
   vOrgTranValue number := null;
   vCardNo       varchar2(20) := null;
   vCardMbr1     number :=null;
   vStateId      number := 0;
   begin
    vSing := '';
    if vMoney_In_Tran > 0 then
      vSing := 'CR';
    end if;
    if vMoney_Out_Tran > 0 then
      vSing := 'DB';
    end if;
    if vOrgValue != 0 then
      vOrgTranValue := vOrgValue;
    end if;
    if nvl(vCardPan,' ') != ' ' then
      vCardNo := vCardPan;
      vCardMbr1 := vCardMbr;
    else
      vCardNo := vPrinaryCardNo;
      vCardMbr1 := vPrimaryMbr;
    end if;
      vStateId  := sStatementId + GetStatId(nvl(vCardNo,' '));

    insert into tStatementDetailTable values(sBranch, vStateId,sStatementRecNumber,vchContract, vchDeposit, vAccountCurrency,vCardNo,vCardMbr1,
    vTranDate,vValuedate,vDescription,vTermLocation,vOrgTranValue,vOrgAbbr,vMoney_Out_Tran+vMoney_In_Tran,vSing,
    vDocno,vTranRef);
   exception
     when others then
       s.say('StToTableCorp.Save2DetailTable Error:'||sqlErrm);
       null;
   end;
  procedure GetOutstanding_Transactions is
    vDetail varchar2(100);
  Begin
    vTotalSuspends := 0;
    for i in (
      select
      Operdate,
      Packno,
      TermLocation,
      Device,
      Pan,
      Mbr,
      Retailer,
      Trancode,
      Entcode,
      OrgCurrency,
      OrgValue,
      decode(ltrim(rtrim(device))||ltrim(rtrim(entcode)),'B22',WithDrawal,
                                       'B23',WithDrawal,
                                       'P21',WithDrawal,
                                       'P22',WithDrawal,
                                       'V21',WithDrawal,
                                       'V22',WithDrawal,
                                             Value) Money_Out,
      0  Money_In
      from tExtractDelay
    where branch = sBranch and
      Docno is Null  and
      trunc(Operdate) between vFromDate and vToDate and
      (CreditAccountNo  = vchDeposit or
      DebitAccountNo    = vchDeposit)
    order by 1) loop

    vDetail := ' ';
    begin
      select substr(Name,1,100) Name
        into vDetail
        from tReferenceExtractOperation
        where
          device  = i.Device and
          entcode = i.Entcode and
          rownum = 1;
          exception
          when NO_DATA_FOUND then
          null;
      end;
      vDetail       := nvl(vDetail,' ');
      vDocno        := null;--i.Packno;
      vCardPan      := i.Pan;
      vCardMBR      := i.MBR;
--      vMerchant     := i.Retailer;
      vTranRef      := GetTranRef(i.Trancode);--20.04.2006
      vOrgValue     := i.OrgValue;
      begin
        vOrgAbbr := ' ';
        select Abbreviation
          into vOrgAbbr
          from tReferenceCurrency
          where branch  = sBranch and
            currency = i.OrgCurrency and
            rownum = 1;
            exception
            when NO_DATA_FOUND then
            null;
        vOrgAbbr  := nvl(vOrgAbbr,' ');
      end;
      vValuedate    := null;
      vTranDate    := trunc(i.Operdate);
      vDescription  := vDetail;
      vtermlocation := i.termlocation;
      vMoney_Out_Tran    := i.Money_Out;
      vMoney_In_Tran     := i.Money_In;
      if i.Money_Out != 0 or i.Money_In != 0 then
        vTotalSuspends := vTotalSuspends+ i.Money_Out ;
        Save2DetailTable;
      end if;
    end loop;
  End;

PROCEDURE GetEntryGroup(pEntCode number,pMoney_Out number,pMoney_In number) is
  found boolean := false;
 begin
   found := false;
   for i in (
   select  nvl(cashAdvance, 3) cashAdvance
     from tcontractEntryGroupList cl,
          tContractEntryGroup cg
    where
      cl.branch = sbranch and
      cl.EntCode = pEntCode and
      cl.Groupid != 0 and
      cg.branch = cl.branch and
      cg.groupid = cl.groupid ) loop
    if i.CashAdvance = 1 then
      vWithdrawal := vWithdrawal + pMoney_out - pMoney_in;
     vTotalCreditsAndRefunds := vTotalCreditsAndRefunds + pMoney_in;
      SetWithDrawalForCard(vCardPan,vCardMBR,pMoney_out - pMoney_in);
    elsif i.CashAdvance = 0 then
      vPurchases := vPurchases + pMoney_out - pMoney_in;
      vTotalCreditsAndRefunds := vTotalCreditsAndRefunds + pMoney_in;
    elsif i.CashAdvance = 2 then
      vPayment_credits := vPayment_credits - pMoney_out + pMoney_in;
      vTotalCreditsAndRefunds := vTotalCreditsAndRefunds - pMoney_out;-- + pMoney_in;
    elsif i.CashAdvance = 3 then
      vOther := vOther + pMoney_out - pMoney_in;
      vTotalCreditsAndRefunds := vTotalCreditsAndRefunds + pMoney_in;
    end if;
  found := true;
  end loop;
 if not found then
    vOther := vOther + pMoney_out - pMoney_in;
    vTotalCreditsAndRefunds := vTotalCreditsAndRefunds + pMoney_in;
 end if;
 end;
procedure GetPan is
  vmdOpen varchar2(10) := null;
  vCardFlag  boolean;
  Begin
    vPan := ' ';
    vMbr := 0;
    vIdClient := null;
    sPanCounter := 0;
    vPrinaryCardNo := '';
    vPrimaryMbr    := 0;
    vmdOpen := referencecrd_stat.CARD_OPEN;
    for i in(
      select decode(c.idclient, vIdclient,0,1) decodePrim, decode(c.crd_stat,vmdOpen,' ',crd_stat) decodeOpen,
        c.Pan,c.mbr,c.idclient, c.createdate, c.canceldate,c.updatedate,c.cardproduct,
        c.Signstat,c.CRD_STAT, c.currencyno,c.parentpan, c.orderdate,
        c.nameoncard,decode(c.CRD_STAT,ReferenceCrd_STAT.CARD_VIP,'Y','N') VIP,
        decode(c.idclient,cont.idclient,'Y','N') PrimaryCard,
        cont.idclient ContractClient,
        c.branchPart,
        cl.FIO,
        'Credit' Cardtype
     -- select c.Pan,c.idclient,c.CRD_STAT, c.mbr
      from tContract cont, tContractItem ci, tCard c, tclientpersone cl
      where
        cont.branch = sBranch and
        cont.no    = vChContract and
        ci.branch  = cont.Branch and
        ci.no      = cont.no and
        ci.ItemType= '2' and
        c.Branch   = ci.Branch and
        c.Pan = substr(ci.key,1,(length(ci.key)-1)) and
        c.mbr = to_number(substr(ci.key,length(ci.key),1)) and
        cl.branch   = c.branch and
        cl.idclient = c.idclient
        order by 1, 2) loop
      ---If card is reissue
      vCardFlag := true;
      for j in 1..sPanCounter loop
        if i.pan = sPanArr(j).Pan then
          if i.ParentPan is not null then
            s.say ('Card is reissued',2);
            vContractClient := i.ContractClient;
            sPanArr(j).Pan := i.Pan;
            sPanArr(j).MBR := i.MBR;
            sPanArr(j).Contract := vChContract;
            sPanArr(j).Idclient := i.IdClient;
            sPanArr(j).Status   := i.Crd_Stat;
            sPanArr(j).Cardtype         := i.cardtype; --'Credit';
            sPanArr(j).CardCreationDate := i.createdate;
            sPanArr(j).CardExpiryDate   := i.canceldate;
            sPanArr(j).CardLastModificationDate := i.updatedate;
            sPanArr(j).CardProduct := ReferenceCardProduct.GetName(i.cardproduct);
            sPanArr(j).CardState   := ReferenceCardSign.GetName(i.Signstat);
            sPanArr(j).CardStatus  := ReferenceCrd_Stat.GetName(i.CRD_STAT);
            sPanArr(j).CardStatusdate := GetCardChangeState(i.pan,i.mbr);
            sPanArr(j).CardCurrency  := ReferenceCurrency.GetAbbreviation(i.currencyno);
            sPanArr(j).CardBirthDate := i.orderdate;
            sPanArr(j).CardTitle     := i.nameoncard; /* */
            sPanArr(j).CardVip        := i.VIP;
            sPanArr(j).CardPrimary    := i.PrimaryCard;
            sPanArr(j).CardBranchPart := i.branchPart;
            sPanArr(j).CardBranchPartName := ReferenceBranchPart.GetBranchPartName(i.branchPart);
            sPanArr(j).CardAccountNo  := vchDeposit;
            sPanArr(j).CardClientName  := i.FIO;
            sPanArr(j).CardPaymentMethod  := GetPayMethod(vChContract, vchDeposit);
            sPanArr(j).CardLimit := null;
            sPanArr(j).CardAvailableLimit := 0;
            sPanArr(j).CardCashLimit := null;
            sPanArr(j).CardAvailableCashLimit := 0;
            GetDailyLimit(i.Pan,i.Mbr);
          end if;
        vCardFlag := false;
        end if;
      end loop;
      ------------------------
      if vCardFlag = true then
        sPanCounter := sPanCounter + 1;
        vContractClient := i.ContractClient;
        sPanArr(sPanCounter).Pan := i.Pan;
        sPanArr(sPanCounter).MBR := i.MBR;
        sPanArr(sPanCounter).Contract := vChContract;
        sPanArr(sPanCounter).Idclient := i.IdClient;
        sPanArr(sPanCounter).Status   := i.Crd_Stat;
        sPanArr(sPanCounter).Cardtype         := i.cardtype; --'Credit';
        sPanArr(sPanCounter).CardCreationDate := i.createdate;
        sPanArr(sPanCounter).CardExpiryDate   := i.canceldate;
        sPanArr(sPanCounter).CardLastModificationDate := i.updatedate;
        sPanArr(sPanCounter).CardProduct := ReferenceCardProduct.GetName(i.cardproduct);
        sPanArr(sPanCounter).CardState   := ReferenceCardSign.GetName(i.Signstat);
        sPanArr(sPanCounter).CardStatus  := ReferenceCrd_Stat.GetName(i.CRD_STAT);
        sPanArr(sPanCounter).CardStatusdate := GetCardChangeState(i.pan,i.mbr);
        sPanArr(sPanCounter).CardCurrency  := ReferenceCurrency.GetAbbreviation(i.currencyno);
        sPanArr(sPanCounter).CardBirthDate := i.orderdate;
        sPanArr(sPanCounter).CardTitle     := i.nameoncard; /* */
        sPanArr(sPanCounter).CardVip        := i.VIP;
        sPanArr(sPanCounter).CardPrimary    := i.PrimaryCard;
        sPanArr(sPanCounter).CardBranchPart := i.branchPart;
        sPanArr(sPanCounter).CardBranchPartName := ReferenceBranchPart.GetBranchPartName(i.branchPart);
        sPanArr(sPanCounter).CardAccountNo  := vchDeposit;
        sPanArr(sPanCounter).CardClientName  := i.FIO;
        sPanArr(sPanCounter).CardPaymentMethod  := GetPayMethod(vChContract, vchDeposit);
        sPanArr(sPanCounter).CardLimit := null;
        sPanArr(sPanCounter).CardAvailableLimit := 0;
        sPanArr(sPanCounter).CardCashLimit := null;
        sPanArr(sPanCounter).CardAvailableCashLimit := 0;
        GetDailyLimit(i.Pan,i.Mbr);
      end if;
    end loop;
    if not GetPrimaryOpenCard then
     if not GetPrimaryNotClosedCard then
       if not GetPrimaryAnyCard then
          vPrinaryCardNo := '';
          vPrimaryMbr    := 0;
       end if;
     end if;
    end if;
  End;
function GetStatId(pPan in varchar) return number
  is
  vCurrPan number := 0;
  vPrimaryCard number := 0;
  begin
    for i in 1..sPanCounter loop
      if sPanArr(i).pan = pPan then
        return vCurrPan;
      end if;
      if sPanArr(i).IdClient = vIdClient and vPrimaryCard = 0 then
        vPrimaryCard := vCurrPan;
      end if;
      vCurrPan := vCurrPan + 1;
    end loop;
    if vCurrPan != 0 then
      s.say('UNKNOWN CARD',2);
      vCurrPan := vPrimaryCard;
    else
      s.say('NO CARD!!!',2);
    end if;
    return vCurrPan;
  end;
procedure GetDailyLimit(pPan in varchar,pMbr in number)-- return number
  is
  vListLim  APITypes.typeCardLimitList;
  begin
    vListLim := ReferenceLimit.GetCardLimitsApiTypes(pPan, pMBR);
    for i in 1..vListLim.Count loop
      if vListLim(i).Code = 3  then
        sPanArr(sPanCounter).CardLimit := vListLim(i).Value;
      end if;
      if vListLim(i).Code = 2 then
        sPanArr(sPanCounter).CardCashLimit := vListLim(i).Value;
      end if;
    end loop;
  end;
procedure SetWithDrawalForCard(pCardPan in varchar,pCardMBR in number,pWithdrawal in number)
  is
  begin
    for i in 1..sPanCounter loop
      if sPanArr(i).Pan = pCardPan and sPanArr(i).Mbr = pCardMBR then
        if sPanArr(i).CardAvailableCashLimit is not null then
          sPanArr(i).CardAvailableCashLimit := sPanArr(i).CardAvailableCashLimit - pWithdrawal;
        end if;
        exit;
      end if;
    end loop;
  end;
procedure SetUsageForCard(pCardPan in varchar,pCardMBR in number,pSum in number)
  is
  begin
    for i in 1..sPanCounter loop
      if sPanArr(i).Pan = pCardPan and sPanArr(i).Mbr = pCardMBR then
        if sPanArr(i).CardAvailableLimit is not null then
          sPanArr(i).CardAvailableLimit := sPanArr(i).CardAvailableLimit - pSum;
        end if;
        exit;
      end if;
    end loop;
  end;
FUNCTION GetPrimaryOpenCard return boolean
  is
  vFound boolean := false;
  vFistFound boolean := true;
  begin
    for i in 1 .. sPanCounter loop
      if sPanArr(i).Status = '1' and sPanArr(i).IdClient = vContractClient then
        vPan := sPanArr(i).Pan;
        vMBR := sPanArr(i).MBR;
        vPrinaryCardNo := vPan;
        vPrimaryMbr    := vMbr;
        vFound := true;
        exit;
      end if;
    end loop;
    return vFound;
  end;
FUNCTION GetPrimaryNotClosedCard return boolean
  is
  vFound boolean := false;
  vFistFound boolean := true;
  begin
    for i in 1 .. sPanCounter loop
      if sPanArr(i).Status != '9' and sPanArr(i).IdClient = vContractClient then
        vPan := sPanArr(i).Pan;
        vMBR := sPanArr(i).MBR;
        vPrinaryCardNo := vPan;
        vPrimaryMbr    := vMbr;
        vFound := true;
        exit;
      end if;
    end loop;
    return vFound;
  end;
function GetPrimaryAnyCard return boolean
  is
  vFistFound boolean := true;
  vFound boolean := false;
  begin
    for i in 1 .. sPanCounter loop
      if sPanArr(i).IdClient = vContractClient then
        vPan := sPanArr(i).Pan;
        vMBR := sPanArr(i).MBR;
        vPrinaryCardNo := vPan;
        vPrimaryMbr    := vMbr;
        vFound := true;
        exit;
      end if;
    end loop;
  return vFound;
  end;




function GetCardChangeState(pPan in varchar,pMbr in number) return date
  is
  vDate date := null;
  vType number := null;
  begin
    vType := Object.GetType(Card.OBJECT_NAME);
    for i in (Select /*+ ordered
                use_nl(O)
                INDEX_ASC(O IOBJECTLOG_KEYOBJECT)*/
            o.operdate,
            substr(O.REMARK, instr(O.REMARK, ' -> ')+4) State
     From   tObjectLog O
     Where  O.Branch     = sBranch                     and
        O.TypeObject =  vType                      and
        O.KeyObject   = pPan||ltrim(rtrim(to_char(pMBr))) and
        upper(O.Remark) like 'CHANGE STATE%'
      order by 1 desc) loop
      vDate := i.operdate;
      exit;
    end loop;
  /*
    for i in (select operdate, pan,mbr
               from tcardlog
              where branch = sbranch
                and pan = pPan
                and mbr = pMbr
                and eventtype = 6
           order by operdate desc)loop
      vDate := i.operdate;
      exit;
    end loop;
    */
    return vDate;
  end;
function GetPayMethod(pContract in varchar, pAccNo in varchar) return varchar
  is
   vPay boolean;
   vPayMethodHave boolean;
  begin
    vPayMethodHave := true;
    begin
      vPay := SCH_Customer.GetDAFSettings( pContract, pAccNo).Generate;
      exception
       when OTHERS then
         vPayMethodHave := false;
         null;
    end;
    if vPayMethodHave then
      if vPay then
        return 'Direct Debit';
      else
        return 'P.O.S';
      end if;
    else
      return 'Not Defined';
    end if;
    exception
     when OTHERS then
       s.say('GetPayMethod Error:'||sqlErrm,2);
       return null;
  end;
procedure GetFio is
  Begin
    vClientFio := ' ';
    begin
      select FIO
        into vClientFio
      from tClientPersone cl
      where
        cl.branch   = sBranch and
        cl.idClient = vIdClient and
        rownum  = 1;
      exception
      when NO_DATA_FOUND then
      null;
    end;
    vClientFio := nvl(vClientFio,' ');

  End;
function GetCurrencyByAccount(pAccount in varchar) return varchar
  is
  vCurrency varchar2(3) := null;
  begin
    for i in (select rc.ABBREVIATION
      from taccount a, treferencecurrency rc
      where a.branch = sBranch
        and a.accountno = pAccount
        and rc.branch = a.branch
        and rc.currency = a.currencyno)loop
      vCurrency := i.ABBREVIATION;
    end loop;
    return nvl(vCurrency,' ');
    exception
     when OTHERS then
       s.say('StToTableCorp.GetCurrencyByAccount Error:'||sqlErrm);
       return null;
  end;
function GetTotalCloseBalance return number is
  Cardholder varchar2(20) := null;
  vTotalBalance number := 0;
  vCardholderDeposit varchar2(20) := null;
  begin
      for i in (
        select cl.LinkNo Contractno
        from tContractLink cl, tContract c
        where cl.Branch = sBranch and
              cl.MainNo = vContract  and
              c.Branch = cl.Branch and
              c.No = cl.LinkNo    and
              (c.closedate is null or
               c.closedate <= vToDate )
        order by 1) loop
        if vDeposit = vDepositEGP then
          vCardholderDeposit := nvl(Contract.GetAccountNoByItemName(i.Contractno,'ItemDepositDOM'),' ');
        else
          vCardholderDeposit := nvl(Contract.GetAccountNoByItemName(i.Contractno,'ItemDepositINT'),' ');
        end if;
        vTotalBalance := vTotalBalance + GetEndBalance(vCardholderDeposit, vToDate);
      end loop;
      return vTotalBalance;
    exception
     when OTHERS then
       s.say('StToTableCorp.GetTotalCloseBalance Error:'||sqlErrm);
    return 0;
  end;

procedure GetCardHolder is
  vDueAmount number;
    Begin
      sCounterCard := 0;
      vCardsCounter    := 0;
      vRecCounter      := 0;
      vTotalCloseBalance := nvl(GetTotalCloseBalance,0);
      for i in (
        select cl.LinkNo Contractno
        from tContractLink cl, tContract c
        where cl.Branch = sBranch and
              cl.MainNo = vContract  and
              c.Branch = cl.Branch and
              c.No = cl.LinkNo    and
              (c.closedate is null or
               c.closedate <= vToDate )
        order by 1) loop
        vchContract := i.Contractno;
        vOverDueAmount := 0;
        if vDeposit = vDepositEGP then
          vChDeposit := nvl(Contract.GetAccountNoByItemName(i.Contractno,'ItemDepositDOM'),' ');
        else
          vChDeposit := nvl(Contract.GetAccountNoByItemName(i.Contractno,'ItemDepositINT'),' ');
        end if;
        vAccountCurrency := GetCurrencyByAccount(vChDeposit);
        if vChDeposit != ' ' then
          if vDueDate is Null then
            begin
              vDueDate := SCH_Customer.GetDueDate(vChContract,vToDate);
              exception
               when OTHERS then
                 signErrorDueDate := true;
                 s.say('SCH_Customer.GetDueDate. Error:'||sqlErrm);
                 null;
            end;
          end if;
          GetBeginBalance;
          vOpen_Balance := vRemain;
          vMoney_Out := 0;
          vMoney_In := 0;
          GetPan;
          s.say('     vPrinaryCardNo '||vPrinaryCardNo,2);
          GetTransactions;
          vDebit :=  vMoney_Out;
          vCredit := vMoney_In;
          vClose_Balance := vOpen_Balance + vCredit - vDebit;
          vTotDebit  := vTotDebit + vDebit;
          vTotCredit := vTotCredit + vCredit;
          vTotalMoney_Out :=  vDebit;
          vTotalMoney_In  :=  vCredit;
          vTotOpenBalance  := vTotOpenBalance + vOpen_Balance;
          vTotCloseBalance := vTotCloseBalance +  vClose_Balance;
          vPrevious_balance :=  vOpen_Balance;
          vNew_balance      :=  vClose_Balance;
          vOverDueAmount := 0;

          GetDueAmount(vChDeposit,vToDate, vDueAmount, vOverDueAmount);
          -----
           begin
              vMinPay := nvl(SCH_Customer.GetMinPayment(vChcontract,vChDeposit,vToDate),0);
              exception
               when OTHERS then
                 vMinPay := 0;
           end;
          vTotMinPay := vTotMinPay + vMinPay;
          if vPrizTitle = 0 then
            PrintTitle;
            sCounterCard := 0;
            vPrizTitle := 1;
          end if;
          vMinDueAmount :=  vMinPay;
          PrintTransactions;
          GetOutstanding_Transactions;
          SaveDate2MasterTable;
        end if;
      end loop;
    End;
procedure GetPaymentDue
 is
 vDepositCH varchar2(40) := null;
 vDueAmount number := null;
 --vContractCH varchar2(40) := null;
 begin
   vTotalOverlimit := 0;
   for i in (
     select cl.LinkNo Contractno
     from tContractLink cl, tContract c
     where cl.Branch = sBranch and
           cl.MainNo = vContract  and
           c.Branch = cl.Branch and
           c.No = cl.LinkNo    and
           c.closedate is null
     order by 1) loop
     --vContractCh := i.Contractno;
     if vDeposit = vDepositEGP then
       vDepositCh := nvl(Contract.GetAccountNoByItemName(i.Contractno,'ItemDepositDOM'),' ');
     else
       vDepositCh := nvl(Contract.GetAccountNoByItemName(i.Contractno,'ItemDepositINT'),' ');
     end if;
     if vDepositCh != ' ' then
       begin
         vMinimum_Payment := SCH_Customer.GetMinPayment(i.Contractno,vDepositch,vToDate);
         exception
          when OTHERS then
            vMinimum_Payment := 0;
       end;
       vTotal_Payment := vTotal_Payment + vMinimum_Payment;
       -----
       GetDueAmount(vDepositCh,vToDate, vDueAmount, vOverDueAmount);
       -----
     vTotalOverlimit := vTotalOverlimit + vOverDueAmount;
     end if;
   end loop;
 exception
   when others then
     s.say('GetPaymentDue Error:'||sqlErrm);
     null;
 end;
procedure PrintTitle
  is
  Begin
    GetPaymentDue;
    if vPageNumber > 0 then
      AddRec(chr(12)||'Company Name        '|| nvl(vNamecomp,' '));
    else
      AddRec('Company Name        '|| nvl(vNamecomp,' '));
    end if;

    AddRec(  'Contact Person Name '|| nvl( ' ',' '));
    PrintAddress(nvl( ltrim(rtrim(vAddressCont)),' '));
    AddRec( nvl(vCityName,' '));
    AddRec( nvl(vRegionName,' '));
    AddRec( nvl(vCountryName,' '));
    AddRec('Branch Code         '|| nvl(to_char(sBranch),' '));
    AddRec(' ');
    AddRec( nvl(vNamecomp,' '));
    AddRec( nvl(to_char(vStatementDate,'dd/mm/yy'),' '));
    AddRec( nvl(vAccount,' ')||' '||nvl(vAbbreviation,' '));
    AddRec( nvl(vAbbreviation,' ')||' '||nvl(to_char(vTotalLimit,'99999990.00'),'  '));
    vAccountLim := vTotalLimit;
    if vAvailable_Credit > 0 then
      AddRec( nvl(vAbbreviation,' ')||' '||nvl(to_char(vAvailable_Credit,'99999990.00'),'  ')); ---???
      --vAccAvailLim := vAvailable_Credit;
    else
      AddRec( nvl(vAbbreviation,' ')||' '||nvl(to_char(0,'99999990.00'),' '));
      --vAccAvailLim := 0;
    end if;
    AddRec(  nvl(vAbbreviation,' ')||' '||nvl(to_char(vTotal_Payment,'99999990.00'),'  '));
    vTotalDueAmount := vTotal_Payment;
    if not signErrorDueDate then
      AddRec( nvl(to_char(vDueDate,'dd/mm/yy'),' '));
    else
      AddRec('Contract Type Settings Is Not Defined');
    end if;
    vPageNumber := vPageNumber + 1;
    AddRec(to_char(vPageNumber));
    AddRec(' ');
  End;
function  GetNumber(pStr in varchar) return boolean
 is
  vNumber boolean := false;
  vPos    number;
  vKol    number;
 begin
   if substr(pStr,1,1) in ('1','2','3','4','5','6','7','8','9','0') then
     vNumber := true;
     vPos := 1;
     vKol := Instr(substr(pStr,vPos),' ');    --first space
     for i in 1..vKol - 1 loop
       if substr(pStr,i,1) not in ('1','2','3','4','5','6','7','8','9','0') then
         vNumber := false;
       end if;
     end loop;
   end if;
   return vNumber;
 exception
 when OTHERS then
   s.say('stdebit.GetNumber Error:'||sqlErrm);
   return false;
 end;--GetNumber
procedure  PrintAddress(pAddressCont in varchar)
  is
  vPos number;
  vPos1 number;
  vPos2 number;
  vKol1 number;
  vKol2 number;
  vAddressCont varchar2(300) := null ;
  vAddressLine varchar2(300) := null;
  vNumber boolean := true;
  vStr    varchar2(300) := null;
  vLengthKol number;
  vLengthKol2 number;
  vTranCount  number := 0;
  vFirstLine   boolean := true;
  vForNum number := 0;
  vSumForNum number := 0;
  vFirstLine2 boolean := true;
  begin
    -- 'Company Address     '
    -- '
    vLengthKol := 50;
    vAddressLine := pAddressCont;
    if substr(vAddressLine,1,1) in ('1','2','3','4','5','6','7','8','9','0')  then
      vPos1 := 1;
      vKol1 := Instr(substr(vAddressLine,vPos1),' ');    --first space
      vNumber := true;
      for i in 2..vKol1 - 1 loop
        if substr(vAddressLine,i,1) not in ('1','2','3','4','5','6','7','8','9','0') then
          vNumber := false;
        end if;
      end loop;
      if vNumber then
        vAddressLine := chr(55473)||chr(55682)||chr(55685) ||' '||vAddressLine;
      end if;
    end if;
    s.say(vAddressLine,2);
    if nvl(length(ltrim(rtrim(vAddressLine))),0) <= vLengthKol then
      AddRec('Company Address     '||Service.rightpad(vAddressLine,vLengthKol));
      vTranCount := vTranCount + 1;
      vAddressCont := vAddressLine;
      vFirstLine := false;
    else
      vPos     := 1;
      vPos1    := 1;
      vPos2    := 0;
      vLengthKol2 := length(ltrim(rtrim(vAddressLine)));
      vKol1    := Instr(substr(vAddressLine,vPos1),' ');
      vKol2   := Instr(substr(vAddressLine,vPos1),' ');
      for i in 1..100000 loop
        if Instr(substr(vAddressLine,vPos),' ') = 0 then
          if length(substr(vAddressLine,vPos1)) > (vLengthKol - vForNum) then
              vTranCount := vTranCount + 1;
              if vTranCount < 4 then
                 vAddressCont := substr(vAddressLine,vPos1,vKol1-vPos1);
                if GetNumber(substr(vAddressLine,vPos1)) then
                   if vFirstLine then
                      AddRec( 'Company Address     '||Service.rightpad(Service.rightpad(chr(55473)||chr(55682)||chr(55685) ||' '||substr(vAddressLine,vPos1,vKol1-vPos1),vLengthKol),vLengthKol));
                      vFirstLine := false;
                   else
                      AddRec( Service.LeftPad(' ', 20,' ')||Service.rightpad(Service.rightpad(chr(55473)||chr(55682)||chr(55685) ||' '||substr(vAddressLine,vPos1,vKol1-vPos1),vLengthKol),vLengthKol));
                   end if;
                else
                  if vFirstLine then
                    vFirstLine := false;
                    AddRec( 'Company Address     '||Service.rightpad(Service.rightpad(substr(vAddressLine,vPos1,vKol1-vPos1),vLengthKol),vLengthKol));
                  else
                    AddRec( Service.LeftPad(' ', 20,' ')||Service.rightpad(Service.rightpad(substr(vAddressLine,vPos1,vKol1-vPos1),vLengthKol),vLengthKol));
                  end if;
                end if;
                vFirstLine2 := true;
                vForNum := 0;
              end if;
            vPos1 := vPos;
          end if;
          exit;
        end if;
        if (vKol1-vPos1) >= (vLengthKol - vForNum) then
          if (vKol1-vPos1) > (vLengthKol - vForNum) then
             if vKol2 <> vKol1 then
                vKol2 := vPos -1;
             end if;
          else
                vKol2 := vKol1;
          end if;
            vTranCount := vTranCount + 1;
            if vTranCount < 4 then
                vAddressCont := substr(vAddressLine,vPos1,vKol2-vPos1);
                if GetNumber(substr(vAddressLine,vPos1)) then
                  if vFirstLine then
                    vFirstLine := false;
                    AddRec( 'Company Address     '||Service.rightpad(Service.rightpad(chr(55473)||chr(55682)||chr(55685) ||' '||substr(vAddressLine,vPos1,vKol2-vPos1),vLengthKol),vLengthKol));
                  else
                    AddRec( Service.LeftPad(' ',20 ,' ')||Service.rightpad(Service.rightpad(chr(55473)||chr(55682)||chr(55685) ||' '||substr(vAddressLine,vPos1,vKol2-vPos1),vLengthKol),vLengthKol));
                  end if;
                else
                  if vFirstLine then
                    AddRec( 'Company Address     '||Service.rightpad(Service.rightpad(substr(vAddressLine,vPos1,vKol2-vPos1),vLengthKol),vLengthKol));
                    vFirstLine := false;
                  else
                    AddRec( Service.LeftPad(' ',20 ,' ')||Service.rightpad(Service.rightpad(substr(vAddressLine,vPos1,vKol2-vPos1),vLengthKol),vLengthKol));
                  end if;
                end if;
              vFirstLine2 := true;
              vForNum := 0;
          if vFirstLine2 then
            vFirstLine2 := false;
            if GetNumber(substr(vAddressLine,vPos)) then
               vForNum := 4;
               vSumForNum := vSumForNum + 4;
            end if;
          end if;
            end if;
          vPos2 := vPos2 + length(substr(vAddressLine,vPos1,vKol2-vPos1))+1;
          if vPos2 >= vLengthKol2 + vSumForNum then
             exit;
          end if;
          vPos    := vKol2+1;
          vPos1   := vPos;
          vKol2  := vKol2+Instr(substr(vAddressLine,vPos),' ');
          vKol1   := vKol2;
        else
          if vFirstLine2 then
            vFirstLine2 := false;
            if GetNumber(substr(vAddressLine,vPos)) then
               vForNum := 4;
            end if;
          end if;
          vPos  := vKol1+1;
          vKol1 := vKol1+Instr(substr(vAddressLine,vPos),' ');
        end if;
      end loop;
        vTranCount := vTranCount + 1;
        if vTranCount < 4 then
          vAddressCont := substr(vAddressLine,vPos1);
          if GetNumber(ltrim(rtrim(substr(vAddressLine,vPos1)))) then
            if vFirstLine then
              AddRec('Company Address     '||Service.rightpad(chr(55473)||chr(55682)||chr(55685) ||' '||ltrim(rtrim(substr(vAddressLine,vPos1))),vLengthKol));
              vFirstLine := false;
            else
              AddRec(Service.LeftPad(' ', 20,' ')||Service.rightpad(chr(55473)||chr(55682)||chr(55685) ||' '||ltrim(rtrim(substr(vAddressLine,vPos1))),vLengthKol));
            end if;
          else
            if vFirstLine then
              vFirstLine := false;
              AddRec( 'Company Address     '||Service.rightpad(ltrim(rtrim(substr(vAddressLine,vPos1))),vLengthKol));
            else
              AddRec( Service.LeftPad(' ', 20,' ')||Service.rightpad(ltrim(rtrim(substr(vAddressLine,vPos1))),vLengthKol));
            end if;
          end if;
        end if;
    end if;
    for i in vTranCount..2 loop
      AddRec(' ');
    end loop;
  exception
   when OTHERS then
     s.say('NSGBCH.PrintAddress Error:'||sqlErrm);
  end;
procedure PrintProgon
  is
  begin
   -- vRecCounter := vRecCounter + 1;
    for i in vRecCounter+1..60 loop
      AddRec(' ');
    end loop;
  end;
procedure PrintTransactions
  is
  vRecCounter number := 0;
  Begin
   sCounterCard := sCounterCard + 1;
   if not PrintOpenCard then
    if not PrintNotClosedCard then
      if not PrintAnyCard then
        if vCardsCounter + 1 > 20 then
          Printprogon;
          vRecCounter := 0;
          vCardsCounter := 0;
          PrintTitle;
        end if;
        vCardsCounter := vCardsCounter + 1;
        AddRec(
        Service.RightPad(to_char(sCounterCard),10,' ')||
        Service.RightPad(' ',20,' ')||
        Service.RightPad(' ',40,' ')||
        Service.LeftPad( nvl(to_char(vClose_Balance,'9999999990.00'),'0.00 '),20,' ')
        );
      end if;
    end if;
   end if;
  End;
FUNCTION PrintOpenCard return boolean
  is
  vFound boolean := false;
  vFistFound boolean := true;
  vCountOpenCard number := 0;
  begin
    for i in 1 .. sPanCounter loop
      if sPanArr(i).Status = '1' then
        vCountOpenCard := vCountOpenCard + 1;
      end if;
    end loop;
    if vCardsCounter + vCountOpenCard >= 20 then
      PrintProgon;
      vRecCounter := 0;
      vCardsCounter := 0;
      PrintTitle;
    end if;
    vCardsCounter := vCardsCounter + vCountOpenCard;
    for i in 1 .. sPanCounter loop
      if sPanArr(i).Status = '1' then
        if vFistFound then
          vIdClient := sPanArr(i).IdClient;
          GetFio;
          AddRec(
          Service.RightPad(to_char(sCounterCard),8,' ')||
          Service.RightPad(nvl(sPanArr(i).PAN,' ')||'-'||to_char(nvl(sPanArr(i).MBR,0)),22,' ')||
          Service.RightPad(nvl(vClientFio,' '),40,' ')||
          Service.LeftPad( nvl(to_char(vClose_Balance,'9999999990.00'),'0.00 '),20,' ')
          );
          vFistFound := false;
        else
          vIdClient := sPanArr(i).IdClient;
          GetFio;
          AddRec(
          Service.RightPad(' ',8,' ')||
          Service.RightPad(nvl(sPanArr(i).PAN,' ')||'-'||to_char(nvl(sPanArr(i).MBR,0)),22,' ')||
          Service.RightPad(nvl(vClientFio,' '),40,' ')
          );
        end if;
        vFound := true;
      end if;
    end loop;
    return vFound;
  end;

FUNCTION PrintNotClosedCard return boolean
  is
  vFound boolean := false;
  vFistFound boolean := true;
  begin
    for i in 1 .. sPanCounter loop
      if sPanArr(i).Status != '9' then
        if vFistFound then
          if vCardsCounter + 1 >= 20 then
            Printprogon;
            vRecCounter := 0;
            vCardsCounter := 0;
            PrintTitle;
          end if;
          vCardsCounter := vCardsCounter + 1;
          vIdClient := sPanArr(i).IdClient;
          GetFio;
          AddRec(
          Service.RightPad(to_char(sCounterCard),8,' ')||
          Service.RightPad(nvl(sPanArr(i).PAN,' ')||'-'||to_char(nvl(sPanArr(i).MBR,0)),22,' ')||
          Service.RightPad(nvl(vClientFio,' '),40,' ')||
          Service.LeftPad( nvl(to_char(vClose_Balance,'9999999990.00'),'0.00 '),20,' ')
          );
          vFistFound := false;
        end if;
        vFound := true;
      end if;
    end loop;
    return vFound;
  end;
function PrintAnyCard return boolean
  is
  vFistFound boolean := true;
  vFound boolean := false;
  begin
    for i in 1 .. sPanCounter loop
      if vFistFound then
        if vCardsCounter + 1 >= 20 then
          Printprogon;
          vRecCounter := 0;
          vCardsCounter := 0;
          PrintTitle;
        end if;
        vCardsCounter := vCardsCounter + 1;
        vIdClient := sPanArr(i).IdClient;
        GetFio;
        AddRec(
        Service.RightPad(to_char(sCounterCard),8,' ')||
        Service.RightPad(nvl(sPanArr(i).PAN,' ')||'-'||to_char(nvl(sPanArr(i).MBR,0)),23,' ')||
        Service.RightPad(nvl(vClientFio,' '),40,' ')||
        Service.LeftPad( nvl(to_char(vClose_Balance,'9999999990.00'),'0.00 '),20,' ')
        );
        vFistFound := false;
        vFound := true;
      end if;
    end loop;
  return vFound;
  end;
procedure PrintEndingTr
  is
  vStrText varchar2(500) := '';
  Begin
    --vText :='Promotional Text';
    AddRec(
    Service.RightPad(' ',30,' ')||
    Service.RightPad(nvl('Total Cards Balance',' '),40,' ')||
    Service.LeftPad( nvl(to_char(vTotCloseBalance,'9999999990.00'),'0.00 '),20,' ')
    );
    while vCardsCounter < 20 loop
       vCardsCounter := vCardsCounter + 1;
       AddRec(' ');
    end loop;
    AddRec(' ');
    AddRec( nvl(to_char(vTotOpenBalance   ,'9999999990.00'),' '));
    AddRec( nvl(to_char(vWithdrawal       ,'9999999990.00'),' '));
    AddRec( nvl(to_char(vPurchases        ,'9999999990.00'),' '));
    AddRec( nvl(to_char(vCredit_Interest  ,'9999999990.00'),' '));
    AddRec( nvl(to_char(vOther            ,'9999999990.00'),' '));
    AddRec( nvl(to_char(vPayment_Credits  ,'9999999990.00'),' '));
    AddRec( nvl(to_char(vTotCloseBalance  ,'9999999990.00'),' '));

    AddRec(' ');
    for i in 1..nvl(Length(vText1),0) loop
      if substr(vText1,i,1) = chr(10) then
        AddRec(vStrText);
        vStrText := '';
      else
        vStrText := vStrText||substr(vText1,i,1);
      end if;
    end loop;
    if ltrim(rtrim(vStrText)) is not null then
        AddRec(vStrText);
    end if;
    vStrText := '';
    for i in 1..nvl(Length(vText2),0) loop
      if substr(vText2,i,1) = chr(10) then
        AddRec(vStrText);
        vStrText := '';
      else
        vStrText := vStrText||substr(vText2,i,1);
      end if;
    end loop;
    AddRec(nvl(vStrText,' '));

    PrintProgon;
  End;
procedure GetCHOverdraft
  is
  Begin
    vOverdraft := 0;
    begin
    Select Overdraft
      into vOverdraft
      from tAccount
      where branch  = sBranch and
        accountno = vChDeposit and
        rownum = 1;
      exception
      when NO_DATA_FOUND then
      null;
     end;
     vOverdraft  := nvl(vOverdraft,0);
     GetDueAmount( vChDeposit,vToDate, vTotalDueAmount, vOverDueAmount);
    vTotalOverLimit := 0;
    if vOverdraft + vNew_balance < 0  then
       vTotalOverLimit := abs(vOverdraft + vNew_balance);
    end if;
     --   vTotalDueAmount,vMinDueAmount, vOverDueAmount,vTotalOverlimit
      begin
        vMinDueAmount  := SCH_Customer.GetUnpaidMinPayment(vCHContract,vChDeposit,vToDate);
        exception
         when OTHERS then
           s.say('StToTableContr.GetUnpaidMinPayment Error:'||sqlErrm);
           vMinDueAmount := 0;
      end;
     exception
      when OTHERS then
        s.say('StToTableCorp.GetCHOverdraft Error:'||sqlErrm);
        null;
  End;

procedure GetDueAmount(pAccNo in varchar,pDate in date, pDueAmount out number, pOverDueAmount out number)
 is
  vDueAmount        number := null;
  vOverDueAmount    number := null;
 begin
   begin
      SCH_Customer.GetDueAmount(pAccNo,pDate, vDueAmount, vOverDueAmount);
      exception
       when OTHERS then
         null;
   end;
   pDueAmount := vDueAmount;
   pOverDueAmount := abs(vOverDueAmount);
 exception
   when others then
     s.say('GetDueAmount Error:'||sqlErrm);
     null;
 end;

procedure SaveDate2MasterTable
  is
  vPrimaryPan varchar2(50) := null;
  vCardCountry varchar2(200) := null;
  vCardRegion  varchar2(100) := null;
  vCardCity    varchar2(200) := null;
  vCardzipCode  varchar2(10) := null;
  vCardAddress1 varchar2(54) := null;
  vCardAddress2 varchar2(54) := null;
  vCardAddress3 varchar2(54) := null;
  vCardAvaiLim     number := 0;
  vCardAvaiCashLim number := 0;
  vCardBarCode     varchar2(20);
  vStateId number;
  vRet number := 0;
  vCardAccountNo varchar2(20);
  vCardAccountClient number;
  begin
    vStateId  := sStatementId;
    vAccAvailLim := vAccountLim + vTotalCloseBalance;
    if vAccAvailLim <= 0 then
      vAccAvailLim := 0;
    end if;
      GetChOverdraft;
      GetChAccountLimits;
    for i in 1.. sPanCounter loop
      if sPanArr(i).CardPrimary != 'Y' then
        vPrimaryPan := vPrinaryCardNo;
      else
        vPrimaryPan := null;
      end if;
    ----
    ----
      SetCardLimits(i,sPanArr(i).CardLimit, vCardAvaiLim, sPanArr(i).CardCashLimit, vCardAvaiCashLim);
      GetAddressCard(sPanArr(i).IdClient,vCardCountry,vCardRegion,vCardCity,vCardzipCode,vCardAddress1,vCardAddress2,vCardAddress3);
      vCardBarCode := GetBarCode(sPanArr(i).IdClient);

      vCardAccountNo     := sPanArr(i).CardAccountNo;
      vCardAccountClient := null;
      begin
        select idclient
          into vCardAccountClient
          from taccount
        where branch = sBranch
          and accountno = vCardAccountNo
          and rownum = 1;
      exception
      when NO_DATA_FOUND then
        null;
      end;
      GetAccUserAttributes(vCardAccountNo);       --новые реквизиты
      GetClientUserAttributes(vCardAccountClient);
      GetCardUserAttributes(sPanArr(i).Pan,sPanArr(i).Mbr);

      insert into tStatementMasterTable values(
      sBranch,vStateId,sStatementRecNumber,vFromDate,vToDate,
      vStatementType,vStatementSendType,vDueDate,vText1,vText2,null,
      -- Contract
      vContract,vContractTypeName,vContractCreateDate,vContractStatus,vCompany,vRegNumber,vTotalLimit,
      -- Client
      vCustomerNo,null,vNameComp,vCustomerCreateDate,vCustomerType,vCustomerStatus,vIdClientCont,
      vCompanyName,null,null,null,vCountryName,vRegionName,vCityName,vZIPCont,
      vAddressLine1,vAddressLine2,vAddressLine3,vAddressBarCode,
      -- Account
      vDeposit,vExternal,vAccountCreateDate,vAccountType,vAccountStatus,vAbbreviation,
      vCountryName,vRegionName,vCityName,vZIPCont,
      vAddressLine1,vAddressLine2,vAddressLine3,vAddressBarCode,
      vAccountLim,vAccAvailLim,vAccCashLim,vAccAvailCashLim,
      -- Card
      sPanArr(i).Pan,sPanArr(i).Mbr,sPanArr(i).Cardtype,
      sPanArr(i).CardCreationDate,sPanArr(i).CardExpiryDate,sPanArr(i).CardLastModificationDate,
      sPanArr(i).CardProduct, sPanArr(i).CardState, sPanArr(i).CardStatus,sPanArr(i).CardStatusDate,
      sPanArr(i).CardCurrency,sPanArr(i).CardBirthDate,sPanArr(i).CardTitle,sPanArr(i).CardVip,sPanArr(i).CardPrimary,
      vPrimaryPan,sPanArr(i).CardBranchPart,sPanArr(i).CardBranchPartName,
      sPanArr(i).CardAccountNo,sPanArr(i).CardClientName,sPanArr(i).CardPaymentMethod,
      sPanArr(i).CardLimit,vCardAvaiLim,sPanArr(i).CardCashLimit,vCardAvaiCashLim,
      vCardCountry,vCardRegion,vCardCity,vCardzipCode,vCardAddress1,vCardAddress2,vCardAddress3,vCardBarCode,
      -- Totals
      vTotalDueAmount,vMinDueAmount,vPrevious_balance,vNew_balance,vPurchases,vWithdrawal,vCredit_Interest,
      vPayment_Credits,vOther,vTotalCreditsAndRefunds ,vTotalMoney_Out,vTotalMoney_In,vTotalSuspends,vOverDueAmount,
      vTotalOverlimit,
      -- Attributes
      sHoldStmt, sBarCode, sUserActField1, sUserActField2, sUserCustField1, sUserCustField2, sUserCardField1, sUserCardField2,
      null, null, null, null
      );
      vStateId := vStateId + 1;
    end loop;

    if sPanCounter =  0 then
      insert into tStatementMasterTable values(
      sBranch,vStateId,sStatementRecNumber,vFromDate,vToDate,
      vStatementType,vStatementSendType,vDueDate,vText1,vText2,null,
      -- Contract
      vContract,vContractTypeName,vContractCreateDate,vContractStatus,vCompany,vRegNumber,vTotalLimit,
      -- Client
      vCustomerNo,null,vNameComp,vCustomerCreateDate,vCustomerType,vCustomerStatus,vIdClientCont,
      vCompanyName,vOffice,vPosition,vEmployee,vCountryName,vRegionName,vCityName,vZIPCont,
      vAddressLine1,vAddressLine2,vAddressLine3,vAddressBarCode,
      -- Account
      vDeposit,vExternal,vAccountCreateDate,vAccountType,vAccountStatus,vAbbreviation,
      vCountryName,vRegionName,vCityName,vZIPCont,
      vAddressLine1,vAddressLine2,vAddressLine3,vAddressBarCode,
      vAccountLim,vAccAvailLim,vAccCashLim,vAccAvailCashLim,
      -- Card
      null,null,null,
      null,null,null,
      null,null,null,null,
      null,null,null,null,null,
      null,null,null,
      null,null,null,
      null,null,null,null,
      null,null,null,null,null,null,null,null,
      -- Totals
      vTotalDueAmount,vMinDueAmount,vPrevious_balance,vNew_balance,vPurchases,vWithdrawal,vCredit_Interest,
      vPayment_Credits,vOther,vTotalCreditsAndRefunds ,vTotalMoney_Out,vTotalMoney_In,vTotalSuspends,vOverDueAmount,
      vTotalOverlimit,
      -- Attributes
      sHoldStmt, sBarCode, sUserActField1, sUserActField2, sUserCustField1, sUserCustField2, null, null,
      null, null, null, null
      );
      vStateId := vStateId +1;
    end if;
    sStatementId := vStateId;
    vRet := DataMember.SetNumber(Object.GetType('BRANCH'), sBranch, 'MSCCSTATEMENT_LASTID', vStateId, null);
 --commit;
 exception
   when others then
     s.say('StToTableCorp.SaveDate2MasterTable Error:'||sqlErrm);
     null;
 end;
procedure GetAddressCard(
  pIdClient in number,
  pCardCountry  out varchar,
  pCardRegion   out varchar,
  pCardCity     out varchar,
  pCardzipCode  out varchar,
  pCardAddress1 out varchar,
  pCardAddress2 out varchar,
  pCardAddress3 out varchar)
 is
  vCardCountry   varchar2(200) := null;
  vCardRegion    varchar2(100) := null;
  vCardCity      varchar2(200) := null;
  vCardzipCode   varchar2(54) := null;
  vAddress       varchar2(200):= null;
  vCardAddress1  varchar2(54) := null;
  vCardAddress2  varchar2(54) := null;
  vCardAddress3  varchar2(54) := null;
  vLengthKol number;
  vPos number;
  vPos1 number;
  vPos2 number;
  vLengthKol2 number;
  vKol1 number;
  vKol2 number;
  vPrizn number;
  vRate number := 0;
  vKolCount number := 0;
  vTranCount number := 0;
  vFirstLine boolean := true;
  vcnt number := 0;
  vAddressCont varchar2(300) := null ;
  vAddressLine varchar2(300) := null;
  vNumber boolean := true;
  vStr    varchar2(300) := null;
  vForNum number := 0;
  vSumForNum number := 0;
  vWord varchar2(5) := null;
 begin
   for i in (select nvl(cl.countryCont,cb.countryCont) countryCont,
      nvl(cl.regioncont,cb.regioncont) regioncont,
      nvl(cl.citycont,cb.citycont) citycont,
      nvl(cl.zipcont,cb.zipcont) zipcont,
      nvl(cl.AddressCont, cb.AddressCont) AddressCont
     from tclient c,
          tclientpersone cl,
          tclientBank cb
     where c.branch = sBranch
       and c.idclient = pIdClient
       and cl.branch(+) = c.branch
       and cl.idclient(+) = c.idclient
       and cb.branch(+) = c.branch
       and cb.idclient(+) = c.idclient)loop
     vCardCountry := ReferenceCountry.getname(i.countryCont);
     vCardRegion  := ReferenceRegion.GetName(i.countryCont,i.regioncont);
     vCardCity    := ReferenceCity.getname(i.countrycont,i.regioncont,i.citycont);
     vCardzipCode := i.zipcont;
     vAddress := i.AddressCont;
     exit;
   end loop;
  -----------
    vLengthKol := 50;
    vTranCount := 0;
    vAddressLine := vAddress;
    if substr(vAddressLine,1,1) in ('1','2','3','4','5','6','7','8','9','0')  then
      vPos1 := 1;
      vKol1 := Instr(substr(vAddressLine,vPos1),' ');    --first space
      vNumber := true;
      for i in 2..vKol1 - 1 loop
        if substr(vAddressLine,i,1) not in ('1','2','3','4','5','6','7','8','9','0') then
          vNumber := false;
        end if;
      end loop;
      if vNumber then
        vAddressLine := chr(55473)||chr(55682)||chr(55685) ||' '||vAddressLine;
      end if;
    end if;
    if nvl(length(ltrim(rtrim(vAddressLine))),0) <= vLengthKol then
      vTranCount := vTranCount + 1;
      vCardAddress1 := ltrim(rtrim(vAddressLine));
    else
      vPos     := 1;
      vPos1    := 1;
      vPos2    := 0;
      vLengthKol2 := length(ltrim(rtrim(vAddressLine)));
      vKol1    := Instr(substr(vAddressLine,vPos1),' ');
      vKol2   := Instr(substr(vAddressLine,vPos1),' ');
      for i in 1..100000 loop
        if Instr(substr(vAddressLine,vPos),' ') = 0 then
          if length(substr(vAddressLine,vPos1)) > (vLengthKol - vForNum) then
              if vTranCount < 3 then
                 vTranCount := vTranCount + 1;
                 --vAddressCont := substr(vAddressLine,vPos1,vKol1-vPos1);
                if GetNumber(substr(vAddressLine,vPos1)) then
                   vWord := chr(55473)||chr(55682)||chr(55685) ||' ';
                else
                   vWord := '';
                end if;
                if vTranCount = 1 then
                  vCardAddress1 := ltrim(rtrim(vWord||substr(vAddressLine,vPos1,vKol1-vPos1)));
                elsif vTranCount = 2 then
                  vCardAddress2 := ltrim(rtrim(vWord||substr(vAddressLine,vPos1,vKol1-vPos1)));
                else
                  vCardAddress3 := ltrim(rtrim(vWord||substr(vAddressLine,vPos1,vKol1-vPos1)));
                end if;
                vFirstLine := true;
                vForNum := 0;
              end if;
            vPos1 := vPos;
          end if;
          exit;
        end if;
        if (vKol1-vPos1) >= (vLengthKol - vForNum) then
          if (vKol1-vPos1) > (vLengthKol - vForNum) then
             if vKol2 <> vKol1 then
                vKol2 := vPos -1;
             end if;
          else
                vKol2 := vKol1;
          end if;
            if vTranCount < 3 then
                vTranCount := vTranCount + 1;
                if GetNumber(substr(vAddressLine,vPos1)) then
                  vWord := chr(55473)||chr(55682)||chr(55685) ||' ';
                else
                  vWord := '';
                end if;
                if vTranCount = 1 then
                  vCardAddress1 := ltrim(rtrim(vWord||Service.rightpad(substr(vAddressLine,vPos1,vKol2-vPos1),vLengthKol)));
                elsif vTranCount = 2 then
                  vCardAddress2 := ltrim(rtrim(vWord||Service.rightpad(substr(vAddressLine,vPos1,vKol2-vPos1),vLengthKol)));
                else
                  vCardAddress3 := ltrim(rtrim(vWord||Service.rightpad(substr(vAddressLine,vPos1,vKol2-vPos1),vLengthKol)));
                end if;
              vFirstLine := true;
              vForNum := 0;
          if vFirstLine then
            vFirstLine := false;
            if GetNumber(substr(vAddressLine,vPos)) then
               vForNum := 4;
               vSumForNum := vSumForNum + 4;
            end if;
          end if;
            end if;
          vPos2 := vPos2 + length(substr(vAddressLine,vPos1,vKol2-vPos1))+1;
          if vPos2 >= vLengthKol2 + vSumForNum then
             exit;
          end if;
          vPos    := vKol2+1;
          vPos1   := vPos;
          vKol2  := vKol2+Instr(substr(vAddressLine,vPos),' ');
          vKol1   := vKol2;
        else
          if vFirstLine then
            vFirstLine := false;
            if GetNumber(substr(vAddressLine,vPos)) then
               vForNum := 4;
            end if;
          end if;
          vPos  := vKol1+1;
          vKol1 := vKol1+Instr(substr(vAddressLine,vPos),' ');
        end if;
      end loop;
        if vTranCount < 3 then
          vTranCount := vTranCount + 1;
          vAddressCont := substr(vAddressLine,vPos1);
          if GetNumber(ltrim(rtrim(substr(vAddressLine,vPos1)))) then
            vWord := chr(55473)||chr(55682)||chr(55685) ||' ';
          else
            vWord := '';
          end if;
          if vTranCount = 1 then
            vCardAddress1 := ltrim(rtrim(vWord||Service.rightpad(ltrim(rtrim(substr(vAddressLine,vPos1))),vLengthKol)));
          elsif vTranCount = 2 then
            vCardAddress2 := ltrim(rtrim(vWord||Service.rightpad(ltrim(rtrim(substr(vAddressLine,vPos1))),vLengthKol)));
          else
            vCardAddress3 := ltrim(rtrim(vWord||Service.rightpad(ltrim(rtrim(substr(vAddressLine,vPos1))),vLengthKol)));
          end if;
        end if;
    end if;
    pCardCountry  := vCardCountry;
    pCardRegion   := vCardRegion;
    pCardCity     := vCardCity;
    pCardzipCode  := vCardzipCode;
    pCardAddress1 := vCardAddress1;
    pCardAddress2 := vCardAddress2;
    pCardAddress3 := vCardAddress3;
  -----------
 exception
   when others then
     s.say('StToTableCorp.GetAddressCard Error:'||sqlErrm);
     null;
 end;

procedure GetChAccountLimits
  is
  vListLim  APITypes.TypeAccountLimitList;
  begin
    --AccCashLim
    vListLim := ReferenceLimit.GetAccountLimitsApiTypes(vChDeposit);
    for i in 1..vListLim.Count loop
      if vListLim(i).Code = 1002  then
        vChAccCashLim := vListLim(i).value;
      end if;
    end loop;
    --AccountLim
    vChAccountLim := vOverdraft;
    --AccAvailLim
    if vNew_Balance < 0 and vChAccountLim is not null then
      vChAccAvailLim     := vChAccountLim  - abs(vNew_Balance);
    else
      vChAccAvailLim :=  vChAccountLim + vNew_Balance;
    end if;
    --AccCashLim
    if vChAccCashLim > vChAccountLim then         --09.04.2007
      vChAccCashLim := vChAccountLim;
    end if;
    --AccAvailCashLim
    if vWithdrawal != 0 and vChAccCashLim is not null then
      vChAccAvailCashLim :=vChAccCashLim - vWithdrawal;
    else
      vChAccAvailCashLim :=vChAccCashLim;
    end if;
    if vChAccAvailCashLim > vChAccAvailLim then
      vChAccAvailCashLim := vChAccAvailLim;
    end if;
    if vChAccAvailLim < 0 then
      vChAccAvailLim := 0;
    end if;
    if vChAccAvailCashLim < 0 then
      vChAccAvailCashLim := 0;
    end if;
    exception
     when OTHERS then
       s.say('StToTableCorp.GetChAccountLimits Error:'||sqlErrm);
       null;
  end;
procedure SetCardLimits(pCurrPanNo in number, pCardLim out number, pCardAvailLim out number, pCardCashLim out number, pCardAvailCashLim out number)
  is
   vCardLim           number := null;
   vCardAvailLim      number := null;
   vCardCashLim       number := null;
   vCardAvailCashLim  number := null;
  begin
    --CardLimit
    if sPanArr(pCurrPanNo).CardLimit is null or sPanArr(pCurrPanNo).CardLimit > vChAccountLim then
      sPanArr(pCurrPanNo).CardLimit := vChAccountLim;
    end if;
    vCardLim := sPanArr(pCurrPanNo).CardLimit;
    --CardAvailLimit
    vCardAvailLim := sPanArr(pCurrPanNo).CardAvailableLimit +  vCardLim;
    if vCardAvailLim < 0 then
      vCardAvailLim := 0;
    elsif vCardAvailLim > vCardLim then
      vCardAvailLim := vCardLim;
    end if;
    if vCardLim is null then
      vCardAvailLim := null;
    end if;
    --CardCashLimit
    if sPanArr(pCurrPanNo).CardCashLimit > vCardLim then
      sPanArr(pCurrPanNo).CardCashLimit := vCardLim;
    end if;
    vCardCashLim := sPanArr(pCurrPanNo).CardCashLimit;
    if vCardCashLim is null or vCardCashLim > vChAccCashLim then
      vCardCashLim := vChAccCashLim;
    end if;
    --CardAvailCashLimit
    vCardAvailCashLim := sPanArr(pCurrPanNo).CardAvailableCashLimit + vCardCashLim;
    if vCardAvailCashLim < 0 then
      vCardAvailCashLim := 0;
    elsif vCardAvailCashLim > vCardCashLim then
     vCardAvailCashLim := vCardCashLim;
    end if;
    if vCardCashLim is null then
      vCardAvailCashLim := null;
    end if;
    if vCardAvailCashLim > vCardAvailLim then
      vCardAvailCashLim := vCardAvailLim;
    end if;
    if vCardAvailLim > vChAccAvailLim then
      vCardAvailLim := vChAccAvailLim;
    end if;
    if vCardAvailCashLim > vChAccAvailCashLim then
      vCardAvailCashLim := vChAccAvailCashLim;
    end if;
    pCardLim          := vCardLim;
    pCardAvailLim     := vCardAvailLim;
    pCardCashLim      := vCardCashLim;
    pCardAvailCashLim := vCardAvailCashLim;
  exception
    when others then
      s.say('StToTableCorp.SetCardLimits Error:'||sqlErrm);
      null;
  end;

FUNCTION  StatementEML(pcheckmake boolean) RETURN NUMBER IS
  vaRecords Types.ArrStr1000;
  vClientType NUMBER;
  vCompanyRec Apitypes.typeCompanyRecord;
  vEML VARCHAR2(100);
  vFIO VARCHAR2(200);
  vParameters Statementdataprepare.typeParametersRec;
  vPersoneRec Apitypes.typePersoneRecord;
  vRet NUMBER;
  sNum NUMBER:=0;
  v_email    VARCHAR2(40);
    PROCEDURE fAddRec(pText IN VARCHAR)  IS
      BEGIN
        sNum := sNum + 1;
        Statementdataprepare.sArrRecords(sNum) := SUBSTR(pText, 1, 1000);
      END;
      BEGIN
      vRet := STATEMENT(pcheckmake);
      IF vRet != 0 THEN
        RETURN vRet;
      END IF;
      if pcheckmake then
        return 0;
      end if;
      vParameters := Statementdataprepare.sParameters;
      vaRecords := Statementdataprepare.sArrRecords;
      Statementdataprepare.sArrRecords.DELETE;
      fAddRec('STNO=1' || CHR(10));
      fAddRec('ST=' || vParameters.Statementtype ||
      Statementdataprepare.GetLngSuffix(vParameters.LngCode) || CHR(10));
      fAddRec('IDCLIENT=' || vParameters.IdClient || CHR(10));
      IF vParameters.IdClient IS NOT NULL THEN
        vClientType := Client.GetClientType(vParameters.IdClient);
        IF vClientType = Client.CT_PERSONE THEN
          vPersoneRec := Client.GetPersoneRecord(vParameters.IdClient);
          vEML := vPersoneRec.eMail;
          vFIO := vPersoneRec.FIO;
        ELSE
          vCompanyRec := Client.GetCompanyRecord(vParameters.IdClient);
          vEML := vCompanyRec.eMailCont;
          vFIO := vCompanyRec.Name;
        END IF;
      END IF;
    BEGIN
      SELECT email INTO v_email FROM A4m.TSTATEMENT WHERE branch = sBranch AND
        Contractno = vParameters.Contractno AND email IS NOT NULL AND ROWNUM = 1;
          IF NVL(v_email,' ') != ' ' THEN
            vEML := RTRIM(v_email);
          END IF;
      EXCEPTION
      WHEN OTHERS THEN NULL;
    END;
  --    fAddRec('EML=' || nvl(vParameters.EMail, vEML) || Chr(10));
    fAddRec('EML=' || vEML || CHR(10));
    fAddRec('F=' || vFIO || CHR(10));
    fAddRec('StatementData=');
    fAddRec('<!-- sendto: '||vEML||' -->');
    FOR i IN 1..vaRecords.COUNT LOOP
      --fAddRec(REPLACE(REPLACE(vaRecords(i), CHR(13)||CHR(10), ' '), CHR(12)));
      fAddRec(REPLACE(vaRecords(i), CHR(13)||CHR(10), ''));
    END LOOP;
    fAddRec(CHR(10));
    RETURN 0;
  EXCEPTION WHEN OTHERS THEN
    S.Err('StatementEML');
    RETURN SQLCODE;
END;
  Begin
    sBranch := Seance.GetBranch;
    sBranchName := Seance.GetBranchName;
  End;
/
show errors package body StToTableCorp;

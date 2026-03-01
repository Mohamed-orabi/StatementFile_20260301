CREATE OR REPLACE PACKAGE BODY sttotabledebit AS
   /****************************************************************************\
     Debit Statement To Table
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
   \****************************************************************************/
   type tRecPan is record(Pan varchar2(20),
                          MBR number,
                          idClient number,
                          Status varchar2(1),
                          Printed number,
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
                          CardBranchPart number,
                          CardBranchPartName varchar2(100),
                          CardAccountNo      varchar2(20),
                          CardClientName     varchar2(250),
                          CardPaymentMethod  varchar2(100),
                          CardLimit          number,
                          CardAvailableLimit number,
                          CardCashLimit      number,
                          CardAvailableCashLimit number);
   type tArrPan is table of tRecPan index by binary_integer;

  sPanArr  tArrPan;
  sPanCounter number;
  vPrinaryCardNo  varchar2(50);
  vPrimaryMbr     number;
  --
  sBranch number;
  sBranchName varchar(200);
  sFiid varchar(10);
  sCurrency  number;
  vCheckMake boolean;
  vFromDate date;
  vToDate date;
  vDate date;
  vPrizArh number;
  vNumberStr number;
  vDueDate            date;
  sStatementRecNumber number;
  vStatementType      varchar2(200);
  vStatementSendType  varchar2(100);
  vText1              varchar(4000);
  vText2              varchar(4000);
  --
  vCompany            number;
  -------
  vIdClient           number;
  vCustomerNo         varchar2(20);
  vCustomerTitle      varchar2(80);
  vFio                varchar(250);
  vCustomerCreateDate date;
  vCustomerType       varchar2(40);
  vCompanyName        varchar2(300);
  vCustomerStatus     varchar2(100);
  vOffice             varchar2(100);
  vEmployee           number;
  vPosition           varchar2(100);
  vCountryCont        number;
  vRegionCont         varchar(6);
  vCityCont           number;
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
  vDeposit varchar(40);
  --vIdClient number;
  vExternal           varchar2(40);
  vAccountCreateDate  date;
  vAccountType        varchar2(100);
  vAccountStatus      varchar2(100);
  vAbbreviation varchar(10);
  vAccountLim         number;
  vAccAvailLim        number;
  vAccCashLim         number;
  vAccAvailCashLim    number;
  vOverdraft number;
  --
  vPrintTransaction  boolean;
  vTotalSuspends number;
  vPayment_Credits number;
  vAll number;
  vWithdrawal number;
  vPurchases number;
  vOther number;
  vTotalCreditsAndRefunds number;
  --
  vValuedate date;
  vTranDate date;
  vDescription varchar(250);
  vTermLocation varchar(250);
  vDetail varchar(250);
  vMerchant   varchar2(100);
  vTranRef    varchar2(23);
  vCardPan    varchar2(50);
  vCardMBR    number;
  vMoney_Out number;
  vMoney_In number;
  --vFio varchar(80);
  vIndex  varchar2(10);
  vOrgAbbr varchar(10);
  vPan varchar(50);
  vMBR number;
  vNumber number;
  --vAddressCont varchar(300);
  --vCityCont    number;
  --vRegionCont  varchar(3);
  --vCountryCont number;
  vPrevious_balance number;
  vNew_balance number;
  vRemain number;
  vCredit_Interest number;
  vInterest number;
  vIdent varchar(80);
  vAvailable_Credit number;
  --vOverdraft number;
  vOrgCurrency number;
  vOrgValue number;
  vEntcode varchar(10);
  vTotalMoney_Out number;
  vTotalMoney_In number;
  vmkol number;
  vCnt number;
  vContract  varchar(40);
  vContractType number;
  vDocno       number;
  vAeDocno     number;
  veValuedate  date;
  vdevice      varchar(10);
  vstr         varchar(200);
  --vText        varchar(4000);
  -- формирование постраничной выдачи
  vPrintPreBalance  boolean;
  vTransCounter     number;
  vRecCounter       number;
  vPageNumber       number;
  sStatementId      number;
  PROCEDURE AddRec(pText in varchar);
  PROCEDURE FormMassiv;
  PROCEDURE GetBeginBalance;
  PROCEDURE GetFio;
  PROCEDURE GetTransactions;
  PROCEDURE PrintTitle;
  PROCEDURE PrintEndingTr;
  PROCEDURE PrintTransactions;
  PROCEDURE PrintProgon;
  PROCEDURE SetPageNumber;
  PROCEDURE SetPanPrinted(pPan in varchar);
  PROCEDURE PrintUnPrintedCards(pPrimary in number);
  PROCEDURE PrintHeader(pPrimary in number, pPan in varchar, pMBR in number);
  FUNCTION  PrintOpenCard return boolean;
  FUNCTION  PrintNotClosedCard return boolean;
  FUNCTION  PrintAnyCard return boolean;
  FUNCTION  GetDailyLimit(pPan in varchar,pMbr in number) return number;
  FUNCTION  GetCashLimit(pPan in varchar,pMbr in number) return number;
  FUNCTION  GetAddressLine(pAddressCont in varchar) return varchar;
  FUNCTION  GetNumber(pStr in varchar) return boolean;
  function GetStatementSendType(pStatType in varchar) return varchar;
  procedure ClearTable;
  function GetCardChangeState(pPan in varchar,pMbr in number) return date;
  function GetAccountNoForCard(pPan in varchar,pMbr in number) return varchar;
  procedure SaveDate2MasterTable;
  function GetTranRef(pTrancode in number) return varchar;--20.04.2006
  procedure GetAddressCard(pIdClient in number,pCardCountry out varchar,pCardRegion out varchar, pCardCity out varchar,
    pCardzipCode out varchar,pCardAddress1 out varchar,pCardAddress2 out varchar,pCardAddress3 out varchar);
  procedure SetWithDrawalForCard(pCardPan in varchar,pCardMBR in number,pWithdrawal in number);
  procedure SetUsageForCard(pCardPan in varchar,pCardMBR in number,pSum in number);
  PROCEDURE GetOutstanding_Transactions;
  function GetBarCode(pIdclient in number) return varchar;
  function GetStatId(pPan in varchar) return number;
---
  FUNCTION STATEMENT (pcheckmake boolean) return number is
  Begin
    vprizArh       :=0;
    vcheckmake     := pcheckmake;
    if vcheckmake then
      StatementDataPrepare.sNeedmake := true;
      return 0;
    end if;
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
    vPrinaryCardNo := '';
    vPrimaryMbr    :=0;
    vNumberStr := Statementdataprepare.sArrRecords.COUNT;
    vDeposit := Statementdataprepare.sParameters.Accountno;
    vText1    := StatementDataPrepare.sParameters.PromotionalText;
    vStatementType := nvl(StatementType.GetCaptionById(Statementdataprepare.sParameters.StatementType),' ');
    vStatementSendType := GetStatementSendType(Statementdataprepare.sParameters.StatementType);
    vDueDate := vToDate;
    FormMassiv;
    SetPageNumber;
    if vcheckmake and Statementdataprepare.sNeedMake then
      return 0;
    end if;
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
  PROCEDURE AddRec(pText in varchar) is
  Begin
    vNumberStr  := vNumberStr + 1;
    vRecCounter := vRecCounter + 1;
    Statementdataprepare.sArrRecords(vNumberStr) := SUBSTR(pText, 1, 998) || CHR(13)||CHR(10);
  End;
  PROCEDURE SetPageNumber
  is
  vCurrNumber number := 0;
  vCurrRec varchar2(20000) := null;
  vCurrPageNo number := 0;
  begin
    s.saY(Statementdataprepare.sArrRecords(14),2);
    for i in 0..vPageNumber-1 loop
     -- s.say(to_char( i*60 + 12),2);
      vCurrRec  := ltrim(rtrim(Statementdataprepare.sArrRecords( i*60 + 14)));
      vCurrRec  := translate( vCurrRec, chr(13)||chr(10), ' ');
      vCurrPageNo := to_number(ltrim(rtrim(vCurrRec)));
      vCurrRec  := vCurrRec||'of '||to_char(vPageNumber);
      if vCurrPageNo = vPageNumber then
        vCurrRec := vCurrRec||' END ';
      end if;
      Statementdataprepare.sArrRecords(vCurrNumber*60 + 14) := vCurrRec || CHR(13)||CHR(10);
      vCurrNumber := vCurrNumber + 1;
    end loop;
  end;
  PROCEDURE FormMassiv
  is
  Begin
    vPrintPreBalance := true;
    vTransCounter    := 0;
    vPageNumber      := 0;
    vRecCounter      := 0;
    if not vcheckmake then
      GetFio;
      vdate := vFromDate;
      GetBeginBalance;
      vPrevious_balance := vRemain;
      vdate := vToDate+1;
      GetBeginBalance;
      vNew_balance := vRemain;
      vAvailable_Credit := vRemain;
      PrintTitle;
    end if;
    GetTransactions;
    PrintEndingTr;
    GetOutstanding_Transactions;
    SaveDate2MasterTable;
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
        a.AccountNo = vDeposit
      union all
      select
        0 RemDep,
        -e.Value ValDep
      from tEntry e, tDocument d
      where e.Branch   = sBranch and
        e.DebitAccount = vDeposit and
        d.Branch       = e.Branch and
        d.DocNo        = e.DocNo and
        d.OpDate      >= vDate and
        d.NewDocNo is null
      union all
      select
        0 RemDep,
        e.Value ValDep
      from tEntry e, tDocument d
      where e.Branch    = sBranch and
        e.CreditAccount = vDeposit and
        d.Branch        = e.Branch and
        d.DocNo         = e.DocNo and
        d.OpDate       >= vDate and
        d.NewDocNo is null
      union all
      select
        0 RemDep,
        -e.Value ValDep
      from tEntryarc e, tDocumentarc d
      where e.Branch   = sBranch and
        e.DebitAccount = vDeposit and
        d.Branch       = e.Branch and
        d.DocNo        = e.DocNo and
        d.OpDate      >= vDate and
        d.NewDocNo is null and
        vPrizArh = 1
      union all
      select
        0 RemDep,
        e.Value ValDep
      from tEntryarc e, tDocumentarc d
      where e.Branch    = sBranch and
        e.CreditAccount = vDeposit and
        d.Branch        = e.Branch and
        d.DocNo         = e.DocNo and
        d.OpDate       >= vDate and
        d.NewDocNo is null and
        vPrizArh = 1
      ) T;
    vRemain := nvl(vRemain,0);
  End;
procedure GetFio
  is
  vFirstLine boolean := true;
  vInsider   varchar2(10) := null;
  vVIP       varchar2(10) := null;
  vBlocked   varchar2(10) := null;
  vRow ApiTypes.TypeClientAddPropRecord;
  vAccType   number;
  vAccStat   varchar2(2);
  vCompDepatment number := null;
  vvTitle        varchar2(80):=null;
  vPrivateCustomer  boolean  := true;
  vmdOpen varchar2(10) := null;
  vCardFlag  boolean;
  Begin
    vFio         := ' ';
    vAddressBarCode := '';
    vAbbreviation:= ' ';
    vIdClient    := -1;
    vCityCont    := -1;
    vRegionCont  := ' ';
    vCountryCont := -1;
    vIndex       := ' ';
    vCustomerTitle := '';
    vCustomerCreateDate := null;
    vCustomerType       := null;
    vCompanyName        := null;
    vCustomerStatus     := null;
    vOffice             := null;
    vEmployee           := null;
    vPosition           := null;
    vZIPCont            := null;
    vCountryName        := null;
    vRegionName         := null;
    vCityName           := null;
    vAddressCont        := null;
    vCountryCont        := -1;
    vRegionCont         := null;
    vCityCont           := -1;
    vCustomerNo         := '';
    Begin
      select
      cl.company,cl.Fio,cl.AddressCont,a.AccountType,r.Abbreviation,
      cl.Title,cl.CountryCont,cl.RegionCont,cl.CityCont,cl.ZIPCont,a.idclient ,
      cl.datecreate, 'Private Customer' Type,
      cl.JobTitle,cl.office,cl.TabNo, cl.Insider, cl.VIP, cl.ObjRestricted,
      a.Acct_Stat, a.createdate,pl.external,cl.personalcode
      into vCompany,vFio,vAddressCont,vAccType,vAbbreviation,
      vvTitle,vCountryCont,vRegionCont,vCityCont,vZIPCont,vIdClient,
      vCustomerCreateDate,vCustomerType,
      vPosition,vCompDepatment,vEmployee, vInsider, vVIP,vBlocked,
      vAccStat,vAccountCreateDate,vExternal,vCustomerNo
      from taccount a,tclientpersone cl,treferencecurrency r,tplanaccount pl
      WHERE a.branch = sbranch and
      a.accountno = vDeposit and
      cl.Branch = a.Branch and
      cl.Idclient = a.Idclient and
      r.Branch(+) = a.branch and
      r.Currency(+) = a.Currencyno and
      pl.branch = a.branch and
      pl.no = a.noplan and
      rownum = 1;
      exception
      when NO_DATA_FOUND then
      null;
    End;
    if vIdClient = -1 then --Corporate
      Begin
        select
        cl.companycode,cl.name,cl.AddressCont,a.AccountType,r.Abbreviation,
        '',cl.CountryCont,cl.RegionCont,cl.CityCont,cl.ZIPCont,a.idclient ,
        cl.datecreate, 'Corporate Customer' Type,
        '',null,null, null, '', '',
        a.Acct_Stat, a.createdate,pl.external,cl.regno
        into vCompany,vFio,vAddressCont,vAccType,vAbbreviation,
        vvTitle,vCountryCont,vRegionCont,vCityCont,vZIPCont,vIdClient,
        vCustomerCreateDate,vCustomerType,
        vPosition,vCompDepatment,vEmployee, vInsider, vVIP,vBlocked,
        vAccStat,vAccountCreateDate,vExternal,vCustomerNo
        from taccount a,tclientbank cl,treferencecurrency r,tplanaccount pl
        WHERE a.branch = sbranch and
        a.accountno = vDeposit and
        cl.Branch = a.Branch and
        cl.Idclient = a.Idclient and
        r.Branch(+) = a.branch and
        r.Currency(+) = a.Currencyno and
        pl.branch = a.branch and
        pl.no = a.noplan and
        rownum = 1;
        exception
        when NO_DATA_FOUND then
        null;
      End;
      vPrivateCustomer  := false;
    end if;
    vAccountTYpe := AccountType.GetName(vAccType);
    vAccountStatus := ReferenceAcct_stat.GetName(vAccStat);
    /*if vPrivateCustomer then
    end if;         */
    vRow := Client.GetAddPropByID(vidclient, 'ACTIONER');
    if vinsider = '1' then
       vCustomerStatus := 'Insider';
    elsif nvl(vRow.ValueNum,0) = 1 and vinsider != '1' then
       vCustomerStatus := 'Shareholder';
    else
       vCustomerStatus := 'Other';
    end if;
    if vVip = '1' then
       vCustomerStatus := 'VIP';
    end if;
    if vBlocked = '1' then
       vCustomerStatus := 'Blocked';
    end if;
    vcompany     := nvl(vcompany,0);
    vFio         := nvl(vFio,' ');
    vAddressCont := nvl(vAddressCont,' ');
    vAbbreviation:= nvl(vAbbreviation,' ');
    vIdCLient    := nvl(vIdClient,-1);
    vAddressBarCode := GetBarCode(vIdClient);
    vCityCont    := nvl(vCityCont,-1);
    vRegionCont  := nvl(vRegionCont,' ');
    vCountryCont := nvl(vCountryCont,-1);
    vCountryName := ReferenceCity.GetName(vCountryCont,vRegionCont,vCityCont);
    vRegionName  := ReferenceRegion.GetName(vCountryCont,vRegionCont);
    vCityName    := ReferenceCountry.GetName(vCountryCont);
    --
    vCompanyName := ' ';
    begin
     select c.Name,co.name
       into vCompanyName,vOffice
       from tIdclient2company ic,tclientbank c,tcompanyoffice co
       where ic.Branch  = sBranch and
             ic.code    = vcompany and
             c.branch   = ic.branch and
             c.idclient = ic.idclient and
             co.branch(+)  = c.branch and
             co.idclient(+) = c.idclient and
             co.code(+) = vCompDepatment and
             rownum  = 1;
       exception
       when NO_DATA_FOUND then
       null;
    end;
    vCompanyName := nvl(vCompanyName,' ');
    --s.say(''||,2);
    vPan := ' ';
    vMBR := 0;
    sPanCounter := 0;
    vmdOpen := referencecrd_stat.CARD_OPEN;
    --begin
    for i in (
      select decode(c.idclient, vIdclient,0,1) decodePrim, decode(c.crd_stat,vmdOpen,' ',crd_stat) decodeOpen,
        c.Pan,c.mbr,c.idclient, c.createdate, c.canceldate,c.updatedate,c.cardproduct,
        c.Signstat,c.CRD_STAT, c.currencyno,c.parentpan, c.orderdate,
        c.nameoncard,decode(c.CRD_STAT,ReferenceCrd_STAT.CARD_VIP,'Y','N') VIP,
        decode(c.idclient,vIdclient,'Y','N') PrimaryCard,
        c.branchPart,
        nvl(cl.FIO,cb.Name) FIO,
        decode(cl.branch,null,'Corporate','Debit') Cardtype
        from taccount a,tAcc2Card ac, tCard c, tclientpersone cl, tclientbank cb
       where a.branch  = sbranch and
        a.accountno    = vDeposit and
        ac.Branch      = a.Branch and
        ac.accountno   = a.accountno and
        c.branch       = ac.branch and
        c.pan          = ac.pan and
        c.mbr          = ac.mbr and
        cl.branch(+)   = c.branch and
        cl.idclient(+) = c.idclient and
        cb.branch(+)   = c.branch and
        cb.idclient(+) = c.idclient
        order by 1, 2) loop
      ---If card is reissue
      vCardFlag := true;
      for j in 1..sPanCounter loop
        s.say('---------------------------------------------');
        s.say('sPanCounter= '||sPanCounter,2);
        s.say('i.Pan= '||i.pan,2);
        s.say('i.mbr= '||i.mbr,2);
        s.say('sPanArr(j).Pan= '||sPanArr(j).Pan,2);
        s.say('i.ParentPan= '||i.ParentPan ,2);
        s.say('---------------------------------------------');
        if i.pan = sPanArr(j).Pan then
          if i.ParentPan is not null then
            s.say ('Card is reissued',2);
            sPanArr(j).Pan      := i.Pan;
            sPanArr(j).MBR      := i.MBR;
            sPanArr(j).Idclient := i.IdClient;
            sPanArr(j).Status   := i.Crd_Stat;
            sPanArr(j).Printed  := 0;
            sPanArr(j).Cardtype         := i.cardtype; --'Debit';
            sPanArr(j).CardCreationDate := i.createdate;
            sPanArr(j).CardExpiryDate   := i.canceldate;
            sPanArr(j).CardLastModificationDate := i.updatedate;
            sPanArr(j).CardProduct := ReferenceCardProduct.GetName(i.cardproduct);
            sPanArr(j).CardState   := ReferenceCardSign.GetName(i.Signstat);
            sPanArr(j).CardStatus  := ReferenceCrd_Stat.GetName(i.CRD_STAT);
            sPanArr(j).CardStatusdate := GetCardChangeState(i.pan,i.mbr);
            sPanArr(j).CardCurrency  := ReferenceCurrency.GetAbbreviation(i.currencyno);
            sPanArr(j).CardBirthDate := i.orderdate;
            sPanArr(j).CardTitle     := i.nameoncard;
            sPanArr(j).CardVip        := i.VIP;
            sPanArr(j).CardPrimary    := i.PrimaryCard;
            sPanArr(j).CardBranchPart := i.branchPart;
            sPanArr(j).CardBranchPartName := ReferenceBranchPart.GetBranchPartName(i.branchPart);
            sPanArr(j).CardAccountNo  := vDeposit;
            sPanArr(j).CardClientName  := i.FIO;
            sPanArr(j).CardPaymentMethod  := '';
            sPanArr(j).CardLimit := GetDailyLimit(i.pan,i.mbr);
            sPanArr(j).CardAvailableLimit := sPanArr(sPanCounter).CardLimit;
            sPanArr(j).CardCashLimit :=GetCashLimit(i.pan,i.mbr);
            sPanArr(j).CardAvailableCashLimit := sPanArr(sPanCounter).CardCashLimit;
             -- sPanArr(j).Printed  := 0;
          end if;
          vCardFlag := false;
        end if;
      end loop;
      ------------------------
      if vCardflag = true then
        sPanCounter := sPanCounter + 1;
        sPanArr(sPanCounter).Pan      := i.Pan;
        sPanArr(sPanCounter).MBR      := i.MBR;
        if i.IdClient = vIdClient and vFirstLine then
          vPan := i.Pan;
          vMBR := i.Mbr;
          vFirstLine := false;
        end if;
        sPanArr(sPanCounter).Idclient := i.IdClient;
        sPanArr(sPanCounter).Status   := i.Crd_Stat;
        sPanArr(sPanCounter).Printed  := 0;
        sPanArr(sPanCounter).Cardtype         := i.cardtype; --'Debit';
        sPanArr(sPanCounter).CardCreationDate := i.createdate;
        sPanArr(sPanCounter).CardExpiryDate   := i.canceldate;
        sPanArr(sPanCounter).CardLastModificationDate := i.updatedate;
        sPanArr(sPanCounter).CardProduct := ReferenceCardProduct.GetName(i.cardproduct);
        sPanArr(sPanCounter).CardState   := ReferenceCardSign.GetName(i.Signstat);
        sPanArr(sPanCounter).CardStatus  := ReferenceCrd_Stat.GetName(i.CRD_STAT);
        sPanArr(sPanCounter).CardStatusdate := GetCardChangeState(i.pan,i.mbr);
        sPanArr(sPanCounter).CardCurrency  := ReferenceCurrency.GetAbbreviation(i.currencyno);
        sPanArr(sPanCounter).CardBirthDate := i.orderdate;
        sPanArr(sPanCounter).CardTitle     := i.nameoncard;
        sPanArr(sPanCounter).CardVip        := i.VIP;
        sPanArr(sPanCounter).CardPrimary    := i.PrimaryCard;
        sPanArr(sPanCounter).CardBranchPart := i.branchPart;
        sPanArr(sPanCounter).CardBranchPartName := ReferenceBranchPart.GetBranchPartName(i.branchPart);
        sPanArr(sPanCounter).CardAccountNo  := vDeposit;
        sPanArr(sPanCounter).CardClientName  := i.FIO;
        sPanArr(sPanCounter).CardPaymentMethod  := '';
        sPanArr(sPanCounter).CardLimit := GetDailyLimit(i.pan,i.mbr);
        sPanArr(sPanCounter).CardAvailableLimit := sPanArr(sPanCounter).CardLimit;
        sPanArr(sPanCounter).CardCashLimit :=GetCashLimit(i.pan,i.mbr);
        sPanArr(sPanCounter).CardAvailableCashLimit := sPanArr(sPanCounter).CardCashLimit;
      -- sPanArr(sPanCounter).Printed  := 0;
      end if;
    end loop;
    --
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
function GetAccountNoForCard(pPan in varchar,pMbr in number) return varchar
  is
  vAccountno varchar2(20) := null;
  begin
    for i in (select ac.accountno
        from tacc2card ac
        where ac.branch = sbranch and
          ac.pan = pPan and
          ac.mbr = pMBR and
          ac.acct_stat = ReferenceACCT_STAT.STAT_PRIMARY)loop
      vAccountno := i.accountno;
      exit;
    end loop;
    return vAccountno;
  end;
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
    vNumber_One number;
    vFirstline  boolean := true;
    vPreviosPan varchar2(20) := null;
    vPrimaryCard number := 0;
  Begin
    vTotalMoney_Out := 0;
    vTotalMoney_In  := 0;
    vNumber_One := 1;
    vWithdrawal := 0;
    vPurchases  := 0;
    vCredit_Interest := 0;
    vOther := 0;
    vPayment_Credits := 0;
    vTotalCreditsAndRefunds :=0;
    vPrintTransaction := true;
    vAll := 0;
    for i in (
      select
      decode(c.idclient,vIdClient,1,
              decode(c.idclient,null,1,-1)) PrimaryCard,
      nvl(c.Pan,vPan) Pan,
      Trans.Opdate,
      Trans.ValueDate,
      Trans.Docno,
      Trans.No,
      --nvl(c.Pan,vPan) Pan,
      nvl(c.MBR,vMbr) MBR,
      Trans.Merchant,
      Trans.Description,
      Trans.Ident,
      Trans.DebitEntCode,
      Trans.Termlocation,
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
      decode(v.branch, null, ' ',
                             v.TermLocation),
                             p.TermLocation),
                             ae.TermLocation) Termlocation,
      decode( Ae.branch,null,
      decode( P.branch,null,
      decode( V.branch, null, 0,
                             V.TranCode),
                             P.TranCode),
                             Ae.TranCode) TranCode,
      decode(p.branch,null,
      decode(v.branch, null, ' ',
                             v.Retailer),
                             p.Retailer) Merchant,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null, ' ',
                             v.Pan),
                             p.Pan),
                             ae.Pan) Pan,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null, 0,
                             v.MBR),
                             p.MBR),
                             ae.MBR) MBR,
      re.name Detail,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null, 0,
                             V.OrgCurrency),
                             P.OrgCurrency),
                             Ae.OrgCurrency) OrgCurrency,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null, 0,
                             V.OrgValue),
                             P.OrgValue),
                             Ae.OrgValue) OrgValue,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null, 0,
                             V.Docno),
                             P.Docno),
                             Ae.Docno) AeDocno,
      e.valuedate eValuedate,
      e.value Money_Out,
      0       Money_In
      from tEntry e,tReferenceEntry re,tDocument d,tPoscheque p,tVoiceSlip v,tAtmExt ae
      where e.Branch       = sbranch and
            e.Debitaccount = vDeposit and
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
            ae.DocNo(+)    = d.Docno
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
      decode(v.branch, null, ' ',
                             v.TermLocation),
                             p.TermLocation),
                             ae.TermLocation) Termlocation,
      decode( Ae.branch,null,
      decode( P.branch,null,
      decode( V.branch, null, 0,
                             V.TranCode),
                             P.TranCode),
                             Ae.TranCode) TranCode,
      decode(p.branch,null,
      decode(v.branch, null, ' ',
                             v.Retailer),
                             p.Retailer) Merchant,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null, ' ',
                             v.Pan),
                             p.Pan),
                             ae.Pan) Pan,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null, 0,
                             v.MBR),
                             p.MBR),
                             ae.MBR) MBR,
      re.name Detail,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null, 0,
                             V.OrgCurrency),
                             P.OrgCurrency),
                             Ae.OrgCurrency) OrgCurrency,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null, 0,
                             V.OrgValue),
                             P.OrgValue),
                             Ae.OrgValue) OrgValue,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null, 0,
                             V.Docno),
                             P.Docno),
                             Ae.Docno) AeDocno,
      e.valuedate eValuedate,
      0       Money_Out,
      e.value Money_In
      from tEntry e,tReferenceEntry re,tDocument d,tPoscheque p,tVoiceSlip v,tAtmExt ae
      where e.Branch       = sbranch and
            e.Creditaccount = vDeposit and
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
            ae.DocNo(+)    = d.Docno
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
      decode(v.branch, null, ' ',
                             v.TermLocation),
                             p.TermLocation),
                             ae.TermLocation) Termlocation,
      decode( Ae.branch,null,
      decode( P.branch,null,
      decode( V.branch, null, 0,
                             V.TranCode),
                             P.TranCode),
                             Ae.TranCode) TranCode,
      decode(p.branch,null,
      decode(v.branch, null, ' ',
                             v.Retailer),
                             p.Retailer) Merchant,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null, ' ',
                             v.Pan),
                             p.Pan),
                             ae.Pan) Pan,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null, 0,
                             v.MBR),
                             p.MBR),
                             ae.MBR) MBR,
      re.name Detail,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null, 0,
                             V.OrgCurrency),
                             P.OrgCurrency),
                             Ae.OrgCurrency) OrgCurrency,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null, 0,
                             V.OrgValue),
                             P.OrgValue),
                             Ae.OrgValue) OrgValue,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null, 0,
                             V.Docno),
                             P.Docno),
                             Ae.Docno) AeDocno,
      e.valuedate eValuedate,
      e.value Money_Out,
      0       Money_In
      from tEntryarc e,tReferenceEntry re,tDocumentarc d,tPoschequearc p,
           tVoiceSliparc v,tAtmExtarc ae
      where e.Branch       = sbranch and
            e.Debitaccount = vDeposit and
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
      decode(v.branch, null, ' ',
                             v.TermLocation),
                             p.TermLocation),
                             ae.TermLocation) Termlocation,
      decode( Ae.branch,null,
      decode( P.branch,null,
      decode( V.branch, null, 0,
                             V.TranCode),
                             P.TranCode),
                             Ae.TranCode) TranCode,
      decode(p.branch,null,
      decode(v.branch,null, ' ',
                             v.Retailer),
                             p.Retailer) Merchant,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null, ' ',
                             v.Pan),
                             p.Pan),
                             ae.Pan) Pan,
      decode(ae.branch,null,
      decode(p.branch,null,
      decode(v.branch, null, 0,
                             v.MBR),
                             p.MBR),
                             ae.MBR) MBR,
      re.name Detail,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null, 0,
                             V.OrgCurrency),
                             P.OrgCurrency),
                             Ae.OrgCurrency) OrgCurrency,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null, 0,
                             V.OrgValue),
                             P.OrgValue),
                             Ae.OrgValue) OrgValue,
      decode(Ae.branch,null,
      decode(P.branch,null,
      decode(V.branch, null, 0,
                             V.Docno),
                             P.Docno),
                             Ae.Docno) AeDocno,
      e.valuedate eValuedate,
      0       Money_Out,
      e.value Money_In
      from tEntryarc e,tReferenceEntry re,tDocumentarc d,tPoschequearc p,
           tVoiceSliparc v,tAtmExtarc ae
      where e.Branch       = sbranch and
            e.Creditaccount = vDeposit and
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
      Trans.TermLocation,
      Trans.Merchant,
      nvl(c.Pan,vpan),
      nvl(c.MBR,vMbr),
      c.Pan,
      c.mbr,
      Trans.Detail,
      Trans.OrgCurrency,
      Trans.AeDocno,
      decode(c.idclient,vIdClient,1,
              decode(c.idclient,null,1,-1)),
      Trans.eValueDate
      order by 1 desc,2 desc,3,4,5,6) loop
      if vNumber_One = 1 then
        if vcheckmake then
          StatementDataPrepare.sNeedmake := true;
          return;
        end if;
        vNumber_One := 0;
      end if;
      begin
        vOrgCurrency := i.OrgCurrency;
        vOrgAbbr := ' ';
        select Abbreviation
          into vOrgAbbr
          from tReferenceCurrency
          where branch  = sBranch and
            currency = vOrgCurrency and
            rownum = 1;
            exception
            when NO_DATA_FOUND then
            null;
        vOrgAbbr  := nvl(vOrgAbbr,' ');
      end;
      vOrgValue     := i.OrgValue;
      vValuedate    := i.Opdate;
      vTranDate     := trunc(i.ValueDate);
      vDescription  := nvl(i.Description,' ');
      vTermLocation := nvl(i.TermLocation,' ');
      vDocno        := i.Docno;
      vDetail       := nvl(i.Detail,' ');
      vAeDocno      := i.AeDocno;
      veValuedate   := i.eValueDate;
      vMoney_Out    := i.Money_Out;
      vMoney_In     := i.Money_In;
      vCardPan      := i.CardPan;
      vCardMBR      := i.CardMBR;
      vMerchant     := i.Merchant;
      vTranRef      := GetTranRef(i.Trancode);--20.04.2006
      if vMoney_Out != 0 or vMoney_In != 0 then
        vPayment_Credits := vPayment_Credits + vMoney_In;
        vAll := vAll + vMoney_Out;
        if i.Ident in ('OPEXT_A10','OPEXT_B15','OPEXT_B23','OPEXT_V15','OPEXT_V23','ACC_ADD','ACC_SUB') /*and
          vMoney_In = 0*/ then
          vWithdrawal := vWithdrawal + vMoney_Out;
          SetWithDrawalForCard(vCardPan,vCardMBR,vMoney_out - vMoney_In);
        elsif (instr(i.Ident,'OPCOM') != 0 or instr(i.Ident,'FEE') != 0) and
          vMoney_In = 0 then
          vOther := vOther+ vMoney_Out;
        end if;
        if vFirstline then
          vPreviosPan  := nvl(i.Pan,' ');
          vPrimaryCard := i.PrimaryCard;
          vFirstline := false;
          if nvl(vPan,' ') != nvl(i.Pan,' ') then
            PrintHeader(i.PrimaryCard,i.Pan,i.MBR);
            vTransCounter := vTransCounter + 1;
          end if;
        end if;
        if vPreviosPan != nvl(i.Pan,' ') then
          PrintHeader(i.PrimaryCard,i.Pan,i.MBR);
          vTransCounter := vTransCounter + 1;
        end if;
        PrintTransactions;
        vPreviosPan  := nvl(i.Pan,' ');
        vPrimaryCard := i.PrimaryCard;
        vTotalMoney_Out := vTotalMoney_Out + vMoney_Out;
        vTotalMoney_In  := vTotalMoney_In  + vMoney_In;
      end if;
      SetUsageForCard(vCardPan,vCardMBR,vMoney_Out - vMoney_In);
    end loop;
    vPurchases := vAll-vWithdrawal-vOther;
  End;
  procedure GetOutstanding_Transactions is
  Begin
    vTotalSuspends := 0;
    vPrintTransaction := false;
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
      (CreditAccountNo  = vDeposit or
      DebitAccountNo    = vDeposit)
    order by 1) loop
    vDetail := ' ';
    begin
      vEntcode      := i.Entcode;
      vDevice       := i.Device;
      select substr(Name,1,100) Name
        into vDetail
        from tReferenceExtractOperation
        where
          device  = vDevice and
          entcode = vEntcode and
          rownum = 1;
          exception
          when NO_DATA_FOUND then
          null;
      end;
      vDetail       := nvl(vDetail,' ');
      vDocno        := null;--i.Packno;
      vCardPan      := i.Pan;
      vCardMBR      := i.MBR;
      vMerchant     := i.Retailer;
      vTranRef      := GetTranRef(i.Trancode);--20.04.2006
      vOrgValue     := i.OrgValue;
      begin
        vOrgCurrency := i.OrgCurrency;
        vOrgAbbr := ' ';
        select Abbreviation
          into vOrgAbbr
          from tReferenceCurrency
          where branch  = sBranch and
            currency = vOrgCurrency and
            rownum = 1;
            exception
            when NO_DATA_FOUND then
            null;
        vOrgAbbr  := nvl(vOrgAbbr,' ');
      end;
      vValuedate    := null;
      vTranDate     := trunc(i.Operdate);
      vDescription  := vDetail;
      vtermlocation := i.termlocation;
      vMoney_Out    := i.Money_Out;
      vMoney_In     := i.Money_In;
      if vMoney_Out != 0 or vMoney_In != 0 then
        PrintTransactions;
        vTotalSuspends := vTotalSuspends+ vMoney_Out;
      end if;
    end loop;
  End;

PROCEDURE PrintHeader(pPrimary in number, pPan in varchar, pMBR in number)
  is
  begin
    if vTransCounter >= 20 then
       PrintProgon;
       vTransCounter := 0;
       vRecCounter := 0;
       PrintTitle;
    end if;
    if pPan is null then
      if pPrimary = 1 then
        AddRec('Transaction for Primary Card no. ');
      else
        PrintUnPrintedCards(1);
        AddRec('Transaction for Supplementary Card no. ');
      end if;
    else
      SetPanPrinted(pPan);
      if pPrimary = 1 then
        AddRec('Transaction for Primary Card no. '||nvl(pPAN,' ')||'-'||ltrim(rtrim(to_char(nvl(pMBR,0)))));
      else
        PrintUnPrintedCards(1);
        AddRec('Transaction for Supplementary Card no. '||nvl(pPAN,' ')||'-'||ltrim(rtrim(to_char(nvl(pMBR,0)))));
      end if;

    end if;
  end;

FUNCTION PrintOpenCard return boolean
  is
  vFound boolean := false;
  vFistFound boolean := true;
  vCardList varchar2(3000) := null;
  begin
    vCardList := '';
    for i in 1 .. sPanCounter loop
      if sPanArr(i).Status = '1' and sPanArr(i).idClient = vIdClient then
        if vFistFound = true then
          vCardList := vCardList||nvl(sPanArr(i).PAN,' ')||'-'||ltrim(rtrim(to_char(nvl(sPanArr(i).MBR,0))));
          vFistFound := false;
          vPan := sPanArr(i).PAN;
          vMBR := sPanArr(i).MBR;
          vPrinaryCardNo := vPan;
          vPrimaryMbr    := vMbr;
        else
          vCardList := vCardList||', '||nvl(sPanArr(i).PAN,' ')||'-'||ltrim(rtrim(to_char(nvl(sPanArr(i).MBR,0))));
          vPrinaryCardNo := sPanArr(i).PAN;
          vPrimaryMbr    := sPanArr(i).MBR;
        end if;
        vFound := true;
      end if;
    end loop;
    if ltrim(rtrim(vCardList)) is not null then
      AddRec(vCardList);
    end if;
    return vFound;
  end;
FUNCTION PrintNotClosedCard return boolean
  is
  vFound boolean := false;
  vFistFound boolean := true;
  begin
    for i in 1 .. sPanCounter loop
      if sPanArr(i).Status != '9' and sPanArr(i).idClient = vIdClient then
        if vFistFound then
          AddRec(nvl(sPanArr(i).PAN,' ')||'-'||ltrim(rtrim(to_char(nvl(sPanArr(i).MBR,0)))));
          vFistFound := false;
          vPan := sPanArr(i).PAN;
          vMBR := sPanArr(i).MBR;
          vPrinaryCardNo := vPan;
          vPrimaryMbr    := vMbr;
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
      if sPanArr(i).idClient = vIdClient then
        if vFistFound then
          AddRec(nvl(sPanArr(i).PAN,' ')||'-'||ltrim(rtrim(to_char(nvl(sPanArr(i).MBR,0)))));
          vFistFound := false;
          vFound := true;
          vPan := sPanArr(i).PAN;
          vMBR := sPanArr(i).MBR;
          vPrinaryCardNo := vPan;
          vPrimaryMbr    := vMbr;
          exit;
        end if;
      end if;
    end loop;
  return vFound;
  end;
function GetDailyLimit(pPan in varchar,pMbr in number) return number
  is
  vLimit  number;
  vListLim  APITypes.typeCardLimitList;
  begin
    vLimit := null;
    vListLim := ReferenceLimit.GetCardLimitsApiTypes(pPan, pMBR);
    for i in 1..vListLim.Count loop
      if vListLim(i).Code = 3 /*and vListLim(i).PeriodType = ReferenceLimit.PeriodType_NONE*/ then
        vLimit := vListLim(i).Value;
      end if;
    end loop;
    return vLimit;
  end;
function GetCashLimit(pPan in varchar,pMbr in number) return number
  is
  vLimit  number;
  vListLim  APITypes.typeCardLimitList;
  begin
    vLimit := null;
    vListLim := ReferenceLimit.GetCardLimitsApiTypes(pPan, pMBR);
    for i in 1..vListLim.Count loop
      --if vListLim(i).Code = 102  then
      if vListLim(i).Code = 2  then
        vLimit := vListLim(i).Value;
      end if;
    end loop;
    return vLimit;
  end;
procedure SetWithDrawalForCard(pCardPan in varchar,pCardMBR in number,pWithdrawal in number)
  is
  begin
    for i in 1..sPanCounter loop
      if sPanArr(i).Pan = pCardPan and sPanArr(i).Mbr = pCardMBR then
        if sPanArr(i).CardAvailableCashLimit is not null then
          sPanArr(i).CardAvailableCashLimit := sPanArr(i).CardAvailableCashLimit - pWithdrawal;
          s.say('!!!!!!!!!!!!!!!!   SetWithDrawalForCard '||pCardPan||'   SUM '||pWithdrawal||' Result '|| sPanArr(i).CardAvailableCashLimit,2);
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
          s.say('!!!!!!!!!!!!!!!!   SetUsageForCard '||pCardPan||'   SUM '||pSum||' Result '|| sPanArr(i).CardAvailableLimit,2);
        end if;
        exit;
      end if;
    end loop;
  end;
procedure PrintTitle
  is
  vLimit number;
  vStr varchar2(300);
  Begin
    if vPageNumber > 0 then
      AddRec(chr(12)||ltrim(vFio));
    else
      AddRec(ltrim(vFio));
    end if;
    vStr := GetAddressLine(vAddressCont);
    AddRec(vAddressLine1);
    AddRec(vAddressLine2);
    AddRec(vAddressLine3);
    AddRec(ReferenceCity.GetName(vCountryCont,vRegionCont,vCityCont));
    AddRec(ReferenceRegion.GetName(vCountryCont,vRegionCont));
    AddRec(ReferenceCountry.GetName(vCountryCont));
    AddRec(LPad(ltrim(to_char(sBranch)),3,0));
    AddRec(' ');
    vPan := null;
    vMbr := null;
    AddRec(to_char(vToDate,'dd/mm/yy'));     --vToDate
    if not PrintOpenCard then
     if not PrintNotClosedCard then
       if not PrintAnyCard then
         AddRec(' ');
       end if;
     end if;
    end if;
    AddRec(ltrim(rtrim(vAbbreviation))||' '||ltrim(to_char(vNew_balance,'9999990.00')));
    vLimit := GetDailyLimit(vPan,vMBR);
    AddRec(ltrim(rtrim(vAbbreviation))||' '||nvl(ltrim(to_char(vLimit,'9999990.00')),' '));
    vPageNumber := vPageNumber + 1;
    AddRec(ltrim(rtrim(to_char(vPageNumber))));
    AddRec(' ');
    if vPrintPreBalance then
      if vPan is not null then
         AddRec('Transaction for Primary Card no. '||nvl(vPAN,' ')||'-'||ltrim(rtrim(to_char(nvl(vMBR,0)))));
      else
         AddRec('Transaction for Primary Card no. ');
      end if;
      SetPanPrinted(vPan);
      if vPrevious_balance > 0 then
        AddRec(Service.rightpad(Service.rightpad(' ',24)||'Previous Balance',87)||to_char(vPrevious_balance,'9999990.00')||' CR');
      else
        AddRec(Service.rightpad(Service.rightpad(' ',24)||'Previous Balance',87)||to_char(Abs(vPrevious_balance),'9999990.00')||' ');
      end if;
      vPrintPreBalance := false;
    else
      AddRec(' ');
      AddRec(' ');
    end if;
  End;
procedure SetPanPrinted(pPan in varchar)
 is
   vFound boolean := false;
 begin
   for i in 1..sPanCounter loop
     if sPanArr(i).Pan = pPan then
        sPanArr(i).Printed := 1;
        exit;
     end if;
   end loop;
 end;
procedure PrintUnPrintedCards(pPrimary in number)
 is
 begin
   for i in 1..sPanCounter loop
     if sPanArr(i).Printed = 0  then
       if pPrimary = 1 and sPanArr(i).IdClient = vIdClient then
        if vTransCounter >= 20 then
           PrintProgon;
           vTransCounter := 0;
           vRecCounter := 0;
           PrintTitle;
        end if;
         AddRec('Transaction for Primary Card no. '||nvl(sPanArr(i).PAN,' ')||'-'||ltrim(rtrim(to_char(nvl(sPanArr(i).MBR,0)))));
         vTransCounter := vTransCounter + 1;
         sPanArr(i).Printed := 1;
       end if;
       if pPrimary != 1 and sPanArr(i).IdClient != vIdClient then
        if vTransCounter >= 20 then
           PrintProgon;
           vTransCounter := 0;
           vRecCounter := 0;
           PrintTitle;
        end if;
         AddRec('Transaction for Supplementary Card no. '||nvl(sPanArr(i).PAN,' ')||'-'||ltrim(rtrim(to_char(nvl(sPanArr(i).MBR,0)))));
         vTransCounter := vTransCounter + 1;
         sPanArr(i).Printed := 1;
       end if;
     end if;
   end loop;
 end;
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
   s.say('sttotabledebit.GetNumber Error:'||sqlErrm);
   return false;
 end;--GetNumber
FUNCTION  GetAddressLine(pAddressCont in varchar) return varchar
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
  vWord varchar2(5) := null;
  begin
    vAddressLine1 := '';
    vAddressLine2 := '';
    vAddressLine3 := '';
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
    s.say(vAddressLine);
    if nvl(length(ltrim(rtrim(vAddressLine))),0) <= vLengthKol then
      --AddRec(Service.rightpad(vAddressLine,50,' '));
      vAddressLine1 := ltrim(rtrim(vAddressLine));
      vTranCount := vTranCount + 1;
      vAddressCont := vAddressLine;
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
                -- vAddressCont := Service.rightpad(substr(vAddressLine,vPos1,vKol1-vPos1),vLengthKol);
                 vAddressCont := substr(vAddressLine,vPos1,vKol1-vPos1);
                if GetNumber(substr(vAddressLine,vPos1)) then
                   -- AddRec( Service.rightpad(Service.rightpad(chr(55473)||chr(55682)||chr(55685) ||' '||substr(vAddressLine,vPos1,vKol1-vPos1),vLengthKol),vLengthKol));
                   vWord := chr(55473)||chr(55682)||chr(55685) ||' ';
                else
                   -- AddRec( Service.rightpad(Service.rightpad(substr(vAddressLine,vPos1,vKol1-vPos1),vLengthKol),vLengthKol));
                   vWord := '';
                end if;
                if vTranCount = 1 then
                  vAddressLine1 := ltrim(rtrim(vWord || vAddressCont));
                elsif vTranCount = 2 then
                  vAddressLine2 := ltrim(rtrim(vWord || vAddressCont));
                else
                  vAddressLine3 := ltrim(rtrim(vWord || vAddressCont));
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
            vTranCount := vTranCount + 1;
            if vTranCount < 4 then
                --vAddressCont := Service.rightpad(substr(vAddressLine,vPos1,vKol2-vPos1),vLengthKol);
                vAddressCont := substr(vAddressLine,vPos1,vKol2-vPos1);
                if GetNumber(substr(vAddressLine,vPos1)) then
                   --AddRec( Service.rightpad(Service.rightpad(chr(55473)||chr(55682)||chr(55685) ||' '||substr(vAddressLine,vPos1,vKol2-vPos1),vLengthKol),vLengthKol));
                   vWord := chr(55473)||chr(55682)||chr(55685) ||' ';
                else
                   --AddRec(Service.rightpad(Service.rightpad(substr(vAddressLine,vPos1,vKol2-vPos1),vLengthKol),vLengthKol));
                   vWord := '';
                end if;
                if vTranCount = 1 then
                  vAddressLine1 := ltrim(rtrim(vWord || vAddressCont));
                elsif vTranCount = 2 then
                  vAddressLine2 := ltrim(rtrim(vWord || vAddressCont));
                else
                  vAddressLine3 := ltrim(rtrim(vWord || vAddressCont));
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
        vTranCount := vTranCount + 1;
        if vTranCount < 4 then
          --vAddressCont := Service.rightpad(ltrim(rtrim(substr(vAddressLine,vPos1))),vLengthKol);
          vAddressCont := ltrim(rtrim(substr(vAddressLine,vPos1)));
          if GetNumber(ltrim(rtrim(substr(vAddressLine,vPos1)))) then
             --AddRec(Service.rightpad(chr(55473)||chr(55682)||chr(55685) ||' '||ltrim(rtrim(substr(vAddressLine,vPos1))),vLengthKol));
             vWord := chr(55473)||chr(55682)||chr(55685) ||' ';
          else
             --AddRec(Service.rightpad(ltrim(rtrim(substr(vAddressLine,vPos1))),vLengthKol));
             vWord := '';
          end if;
          if vTranCount = 1 then
            vAddressLine1 := ltrim(rtrim(vWord || vAddressCont));
          elsif vTranCount = 2 then
            vAddressLine2 := ltrim(rtrim(vWord || vAddressCont));
          else
            vAddressLine3 := ltrim(rtrim(vWord || vAddressCont));
          end if;
        end if;
    end if;
  /*
    for i in vTranCount..2 loop
      AddRec(' ');
    end loop;
  */
  return '';
  exception
   when OTHERS then
     s.say('sttotabledebit.GetAddressLine Error:'||sqlErrm);
     return ' ';
  end;
procedure PrintProgon
  is
  begin
    for i in vRecCounter+1..60 loop
      AddRec(' ');
    end loop;
  end;
procedure PrintTransactions
-- перенос строки без разрыва слов
  is
  vTextStr varchar(500);
  vRateStr varchar(40);
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
  vSing     varchar2(2) := null;
  vCardNo varchar2(20) := '';
  vCardMbr1 number := 0;
  vOrgTranValue number := null;
  vStateId number := 0;
  Begin
    vLengthKol := 63;
    if vtermlocation != ' ' then
      vTextStr := vDescription||' at '||vTermLocation;
    else
      vTextStr := vDescription;
    end if;
    vSing := '';
    if vMoney_In > 0 then
      vSing := 'CR';
    end if;
    if vMoney_out > 0 then
      vSing := 'DB';
    end if;
    if vCardPan != ' ' then
      vCardNo  := vCardPan;
      vCardMbr1:= vCardMbr;
    else
      vCardNo := vPrinaryCardNo;
      vCardMbr1:= vPrimaryMbr;
    end if;
    if vOrgValue != 0 then
      vOrgTranValue := vOrgValue;
    end if;
    vStateId  := sStatementId + GetStatId(nvl(vCardNo,' '));
    insert into tStatementDetailTable values(sBranch, vStateId ,sStatementRecNumber, null,vdeposit, vCardNo,vCardMbr1,
    vTranDate,vValuedate,vDescription,vTermLocation,vOrgTranValue,vOrgAbbr,vMoney_Out+vMoney_In,vSing,
    vDocno,vTranRef);

    if vPrintTransaction then
      vLengthKol := 63;
      if vMoney_In > 0 then
        vstr := to_char(vMoney_Out+vMoney_In,'9999990.00')||' CR';
      else
        vstr := to_char(vMoney_Out+vMoney_In,'9999990.00')||' DB';
      end if;
      if nvl(vOrgValue,0) > 0 then
        vRate := (vMoney_In + vMoney_Out)/vOrgValue;
      vRateStr := ' RATE '||to_char(vRate,'90.0000')||' '||ltrim(rtrim(vAbbreviation))||' '||ltrim(rtrim(vOrgAbbr));
      end if;
      if vtermlocation != ' ' and nvl(vOrgValue,0) != 0 then
        vTextStr := vDescription||' at '||vTermLocation||' Original currency Amount '||
               ltrim(rtrim(to_char(vOrgValue,'9999999990.00')))||' ('||
               ltrim(rtrim(vOrgAbbr))||')'||vRateStr;
        if vAeDocno != 0 then
          vTextStr := vTextStr||' '||to_char(vevaluedate,'dd/mm/yy');
        end if;
      elsif vtermlocation = ' ' and nvl(vOrgValue,0) != 0 then
        vTextStr := vDescription||' Original currency Amount '||
               ltrim(rtrim(to_char(vOrgValue,'9999999990.00')))||' ('||
               ltrim(rtrim(vOrgAbbr))||')'||vRateStr;
        if vAeDocno != 0 then
          vTextStr := vTextStr||' '||to_char(vevaluedate,'dd/mm/yy');
        end if;
      elsif vtermlocation != ' ' and nvl(vOrgValue,0) = 0 then
        vTextStr := vDescription||' at '||vTermLocation;
        if vAeDocno != 0 then
          vTextStr := vTextStr||' '||to_char(vevaluedate,'dd/mm/yy');
        end if;
      elsif vtermlocation = ' ' and nvl(vOrgValue,0) = 0 then
        vTextStr := vDescription;
        if vAeDocno != 0 then
          vTextStr := vTextStr||' '||to_char(vevaluedate,'dd/mm/yy');
        end if;
      else
        vTextStr := vDescription;
        if vAeDocno != 0 then
          vTextStr := vTextStr||' '||to_char(vevaluedate,'dd/mm/yy');
        end if;
      end if;
      if nvl(length(ltrim(rtrim(vTextStr))),0) <= vLengthKol then
        if vTransCounter >= 20 then
          PrintProgon;
          vTransCounter := 0;
          vRecCounter := 0;
          PrintTitle;
        end if;
        vTransCounter := vTransCounter + 1;
        vTranCount := vTranCount + 1;
        if vTranCount < 4 then
          AddRec(Service.rightpad(to_char(vValuedate,'dd/mm/yyyy'),16)||
            Service.rightpad(to_char(vDocno),8)||
            Service.rightpad(nvl(rtrim(ltrim(vTextStr)),' '),vLengthKol)||
            vstr);
        end if;
      else
        vKolCount :=length(ltrim(rtrim(vTextStr))) / vLengthKol;
        if vKolCount > 3 then
          vKolCount := 3;
        end if;
        if vTransCounter + vKolCount >=20 then
          PrintProgon;
          vTransCounter := 0;
          vRecCounter   := 0;
          PrintTitle;
        end if;
        vPos     := 1;
        vPos1    := 1;
        vPos2    := 0;
        vLengthKol2 := length(ltrim(rtrim(vTextStr)));
        vKol1    := Instr(substr(vTextStr,vPos1),' ');
        vKol2   := Instr(substr(vTextStr,vPos1),' ');
        vPrizn := 0;
        for i in 1..100000 loop
          if Instr(substr(vTextStr,vPos),' ') = 0 then
            if length(substr(vTextStr,vPos1)) > vLengthKol then
              if vPrizn = 0 then
                vTransCounter := vTransCounter + 1;
                vTranCount := vTranCount + 1;
                if vTranCount < 4 then
                  AddRec(Service.rightpad(to_char(vValuedate,'dd/mm/yyyy'),16)||
                    Service.rightpad(to_char(vDocno),8)||
                    Service.rightpad(Service.rightpad(substr(vTextStr,vPos1,vKol1-vPos1),vLengthKol),vLengthKol)||
                    vstr);
                end if;
                vPrizn := 5;
              else
                vTransCounter := vTransCounter + 1;
                vTranCount := vTranCount + 1;
                if vTranCount < 4 then
                   AddRec(Service.rightpad(' ',24)||
                     Service.rightpad(Service.rightpad(substr(vTextStr,vPos1,vKol1-vPos1),vLengthKol),vLengthKol));
                end if;
              end if;
              vPos1 := vPos;
            end if;
            exit;
          end if;
          if (vKol1-vPos1) >= vLengthKol then
            if (vKol1-vPos1) > vLengthKol then
               if vKol2 <> vKol1 then
                  vKol2 := vPos -1;
               end if;
            else
                  vKol2 := vKol1;
            end if;
            if vPrizn = 0 then
              vTransCounter := vTransCounter + 1;
              vTranCount := vTranCount + 1;
              if vTranCount < 4 then
                AddRec(Service.rightpad(to_char(vValuedate,'dd/mm/yyyy'),16)||
                  Service.rightpad(to_char(vDocno),8)||
                  Service.rightpad(Service.rightpad(substr(vTextStr,vPos1,vKol2-vPos1),vLengthKol),vLengthKol)||
                  vstr);
              end if;
              vPrizn := 5;
            else
              vTransCounter := vTransCounter + 1;
              vTranCount := vTranCount + 1;
              if vTranCount < 4 then
                AddRec(Service.rightpad(' ',24)||
                  Service.rightpad(Service.rightpad(substr(vTextStr,vPos1,vKol2-vPos1),vLengthKol),vLengthKol));
              end if;
            end if;
            vPos2 := vPos2 + length(substr(vTextStr,vPos1,vKol2-vPos1))+1;
            if vPos2 >= vLengthKol2 then
               exit;
            end if;
            vPos    := vKol2+1;
            vPos1   := vPos;
            vKol2  := vKol2+Instr(substr(vTextStr,vPos),' ');
            vKol1   := vKol2;
          else
            vPos  := vKol1+1;
            vKol1 := vKol1+Instr(substr(vTextStr,vPos),' ');
          end if;
        end loop;
        if vPrizn = 0 then
          vTransCounter := vTransCounter + 1;
          vTranCount := vTranCount + 1;
          if vTranCount < 4 then
            AddRec(Service.rightpad(to_char(vValuedate,'dd/mm/yyyy'),16)||
              Service.rightpad(to_char(vDocno),8)||
              Service.rightpad(ltrim(rtrim(substr(vTextStr,vPos1))),vLengthKol)||
              vstr);
          end if;
          vPrizn := 5;
        else
          vTransCounter := vTransCounter + 1;
          vTranCount := vTranCount + 1;
          if vTranCount < 4 then
            AddRec(Service.rightpad(' ',24)||
              Service.rightpad(ltrim(rtrim(substr(vTextStr,vPos1))),vLengthKol));
          end if;
        end if;
      end if;
    end if;
  End;

  procedure PrintEndingTr
  is
  vStrText varchar2(500) := null;
  Begin
   -- vTransCounter := vTransCounter + 1;
    PrintUnPrintedCards(1);
    PrintUnPrintedCards(2);

    if vPrevious_balance+vTotalMoney_In-vTotalMoney_Out > 0 then
      AddRec(Service.rightpad(Service.rightpad(' ',24)||'Close Balance',87)||
        to_char(vPrevious_balance+vTotalMoney_In-vTotalMoney_Out,'9999990.00')||' CR');
    else
      AddRec(Service.rightpad(Service.rightpad(' ',24)||'Close Balance',87)||
        to_char(Abs(vPrevious_balance+vTotalMoney_In-vTotalMoney_Out),'9999990.00')||' ');
    end if;
    while vTransCounter < 20 loop
       vTransCounter := vTransCounter + 1;
       AddRec(' ');
    end loop;
    AddRec(' ');
    if vPrevious_balance > 0 then
      AddRec(ltrim(to_char(vPrevious_balance,'9999990.00'))||' CR');
    else
      AddRec(ltrim(to_char(Abs(vPrevious_balance),'9999990.00'))||' ');
    end if;
    --AddRec(ltrim(to_char(vWithdrawal,'9999990.00')));
    AddRec(ltrim(to_char(vPurchases,'9999990.00')));
    --AddRec(ltrim(to_char(vCredit_Interest,'9999990.00')));
    AddRec(ltrim(to_char(vOther,'9999990.00')));
    AddRec(ltrim(to_char(vPayment_Credits,'9999990.00')));
    if vPrevious_balance+vTotalMoney_In-vTotalMoney_Out > 0 then
      AddRec(ltrim(to_char(vPrevious_balance+vTotalMoney_In-vTotalMoney_Out,'9999990.00'))||' CR');
    else
      AddRec(ltrim(to_char(Abs(vPrevious_balance+vTotalMoney_In-vTotalMoney_Out),'9999990.00'))||' ');
    end if;
    AddRec(' ');
    AddRec(' ');
    AddRec(' ');
    for i in 1..nvl(Length(vText1),0) loop
      if substr(vText1,i,1) = chr(10) then
        AddRec(vStrText);
        vStrText := '';
      else
        vStrText := vStrText||substr(vText1,i,1);
      end if;
    end loop;
    AddRec(nvl(vStrText,' '));
    PrintProgon;
    /*
    vPayment_Credits := 0;
    vWithdrawal := 0;
    vPurchases  := 0;                                          */
  End;

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
  vStateId number := 0;
  vRet number := 0;
  begin
    s.say('IIIIIINNNNNNSSSSSEEEEEERRRRRRTTTTT',2);
    vStateId  := sStatementId;

    for i in 1.. sPanCounter loop
      if sPanArr(i).IdClient != vIdClient then
        vPrimaryPan := vPrinaryCardNo;
      else
        vPrimaryPan := null;
      end if;
      if sPanArr(i).CardLimit is not null then
        if sPanArr(i).CardAvailableLimit > 0 and sPanArr(i).CardAvailableLimit <= sPanArr(i).CardLimit then
          vCardAvaiLim := sPanArr(i).CardAvailableLimit;
        elsif sPanArr(i).CardAvailableLimit > 0 and sPanArr(i).CardAvailableLimit > sPanArr(i).CardLimit then
          vCardAvaiLim := sPanArr(i).CardLimit;
        else
          vCardAvaiLim := 0;
        end if;
      else
        vCardAvaiLim := null;
      end if;
      if sPanArr(i).CardCashLimit is not null then
        if  sPanArr(i).CardAvailableCashLimit > 0 and sPanArr(i).CardAvailableCashLimit <= sPanArr(i).CardCashLimit then
          vCardAvaiCashLim := sPanArr(i).CardAvailableCashLimit;
        elsif  sPanArr(i).CardAvailableCashLimit > 0 and sPanArr(i).CardAvailableCashLimit > sPanArr(i).CardCashLimit then
          vCardAvaiCashLim := sPanArr(i).CardCashLimit;
        else
           vCardAvaiCashLim := 0;
        end if;
      else
        vCardAvaiCashLim := null;
      end if;
      vCardBarCode := GetBarCode(sPanArr(i).IdClient);
      GetAddressCard(sPanArr(i).IdClient,vCardCountry,vCardRegion,vCardCity,vCardzipCode,vCardAddress1,vCardAddress2,vCardAddress3);
      if vCardAvaiCashLim > vCardAvaiLim then
        vCardAvaiCashLim := vCardAvaiLim;
      end if;

      insert into tStatementMasterTable values(
      sBranch,vStateId, sStatementRecNumber,vFromDate,vToDate,
      vStatementType,vStatementSendType,vDueDate,vText1,null,null,
      --
      null,null,null,null,null,null,null,
      --
      vCustomerNo,vCustomerTitle,vFio,vCustomerCreateDate,vCustomerType,vCustomerStatus,vIdClient,
      vCompanyName,vOffice,vPosition,vEmployee,vCountryName,vRegionName,vCityName,vZIPCont,
      vAddressLine1,vAddressLine2,vAddressLine3,vAddressBarCode,
      --
      vDeposit,vExternal,vAccountCreateDate,vAccountType,vAccountStatus,vAbbreviation,
      vCountryName,vRegionName,vCityName,vZIPCont,
      vAddressLine1,vAddressLine2,vAddressLine3,vAddressBarCode,
      null,null,null,null,
      --
      sPanArr(i).Pan,sPanArr(i).Mbr,sPanArr(i).Cardtype,
      sPanArr(i).CardCreationDate,sPanArr(i).CardExpiryDate,sPanArr(i).CardLastModificationDate,
      sPanArr(i).CardProduct, sPanArr(i).CardState, sPanArr(i).CardStatus,sPanArr(i).CardStatusDate,
      sPanArr(i).CardCurrency,sPanArr(i).CardBirthDate,sPanArr(i).CardTitle,sPanArr(i).CardVip,sPanArr(i).CardPrimary,
      vPrimaryPan,sPanArr(i).CardBranchPart,sPanArr(i).CardBranchPartName,
      sPanArr(i).CardAccountNo,sPanArr(i).CardClientName,sPanArr(i).CardPaymentMethod,
      sPanArr(i).CardLimit,vCardAvaiLim,sPanArr(i).CardCashLimit,vCardAvaiCashLim,
      vCardCountry,vCardRegion,vCardCity,vCardzipCode,vCardAddress1,vCardAddress2,vCardAddress3,vCardBarCode,
 --
     0,0,0,0,vPurchases,vWithdrawal,0,
     0,vOther,vPayment_Credits ,vTotalMoney_Out,vTotalMoney_In,vTotalSuspends,0,
     0);
      vStateId := vStateId + 1;
    end loop;
    -----------
    if sPanCounter = 0 then
      insert into tStatementMasterTable values(
      sBranch,vStateId,sStatementRecNumber,vFromDate,vToDate,
      vStatementType,vStatementSendType,vDueDate,vText1,null,null,
      -- contract
      null,null,null,null,null,null,null,
      -- Client
      vCustomerNo,vCustomerTitle,vFio,vCustomerCreateDate,vCustomerType,vCustomerStatus,vIdClient,
      vCompanyName,vOffice,vPosition,vEmployee,vCountryName,vRegionName,vCityName,vZIPCont,
      vAddressLine1,vAddressLine2,vAddressLine3,vAddressBarCode,
      --  Account
      vDeposit,vExternal,vAccountCreateDate,vAccountType,vAccountStatus,vAbbreviation,
      vCountryName,vRegionName,vCityName,vZIPCont,
      vAddressLine1,vAddressLine2,vAddressLine3,vAddressBarCode,
      null,null,null,null,
      -- Card
      null,null,null,
      null,null,null,
      null,null,null,null,
      null,null,null,null,null,
      null,null,null,
      null,null,null,
      null,null,null,null,
      null,null,null,null,null,null,null, null,
      -- Total
      0,0,0,0,vPurchases,vWithdrawal,0,
      0,vOther,vPayment_Credits ,vTotalMoney_Out,vTotalMoney_In,vTotalSuspends,0,
      0);
      vStateId := vStateId + 1;
    end if;
    vRet := DataMember.SetNumber(Object.GetType('BRANCH'), sBranch, 'MSCCSTATEMENT_LASTID', vStateId, null);

    -----------
 commit;
 exception
   when others then
     s.say('StToTableContr.SaveDate2MasterTable Error:'||sqlErrm);
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
     s.say('StToTableContr.GetAddressCard Error:'||sqlErrm);
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
    fAddRec('EML=' || vEML || CHR(10));
    fAddRec('F=' || vFIO || CHR(10));
    fAddRec('StatementData=');
    fAddRec('<!-- sendto: '||vEML||' -->');
    FOR i IN 1..vaRecords.COUNT LOOP
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
    sCurrency   := Seance.GetCurrency;
    sFiid := Seance.GetFiid;
  End;
/
show errors package body sttotabledebit;


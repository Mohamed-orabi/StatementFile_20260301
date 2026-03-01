create or replace package body REPEG81 as
  /****************************************************************************\
   Detail Customer Statement
  ****************************************************************************
  Изменения:
  ----------------------------------------------------------------------------
  Дата        Автор            Описание
  ----------  ---------------  ---------------------------------------------
  06.03.2006  Аревкина Е.Н.    в поле Card Limit выводится Credit Limit(Overdraft)
  02.03.2006  Аревкина Е.Н.    добавлена отладочная информация по строке адреса
  01.03.2006  Аревкина Е.Н.    В шапке отчета отражается валюта счета
  28.02.2006  Аревкина Е.Н.    Разработка
  \****************************************************************************/
   type tRecPan is record(Pan varchar2(20),
                          Mbr number,
                          idClient number,
                          Status varchar2(1),
                          Printed number,
                          Total   number);
   type tArrPan is table of tRecPan index by binary_integer;

  sPanArr  tArrPan;
  sPanCounter number;
  svidText      number;
  sBranch       number;
  sCurrency     number;
  --
  sScrollerMode      number;
  sScrollerText      number:=1;
  sScrollerFindCard  number:=2;
  sStrNo        varchar2(3000);
  --
  vFromDate    date;
  vToDate      date;
  vContract    varchar2(20);
  vAccountno   varchar2(20);
  vAbbreviation varchar2(10);
  vIdClient    number;
  vFIO         varchar2(100);
  vPan         varchar2(20);
  vMBR         number;
  vAddressCont varchar2(100);
  vCityCont    varchar2(100);
  vRegionCont  varchar2(100);
  vCountryCont varchar2(100);
  vDepositDOM  varchar2(20);
  vDepositINT  varchar2(20);
  vPrevious_balance number;
  vMinimum_Payment number;
  vAvailable_Credit number;
  vOverdraft        number;
  --
  vTotalMoney_Out   number;
  vTotalMoney_In    number;
  vWithdrawal       number;
  vPurchases        number;
  vCredit_Interest  number;
  vOther            number;
  vPayment_Credits  number;
  vAll              number;
  --
  vOrgAbbr          varchar(10);
  vOrgCurrency      number;
  vOrgValue         number;
  vValuedate        date;
  vTranDate         date;
  vDescription      varchar(250);
  vTermLocation     varchar(250);
  vDocno            number;
  vDetail           varchar(250);
  vAeDocno          number;
  veValuedate       date;
  vMoney_Out        number;
  vMoney_In         number;


  vFileHandle       PLS_INTEGER;
  --
  vprizArh number;
  -- BatchProcessor params
  param_BeginDate    constant number := 1;
  param_EndDate      constant number := 2;
  param_CardNo       constant number := 3;
  param_CardMBR      constant number := 4;
  param_Path         constant number := 5;
  param_FileName     constant number := 6;
  param_UseArchive   constant number := 7;
  function  DialogParameters  return number;
  procedure Generate(pBegindate in date, pEnddate in date, pCardPan in varchar, pCardMBR in number, pPath in varchar, pFileName in varchar, pUseArchive in boolean);
  procedure GetContract(pCardPan in varchar, pCardMBR in number);
  procedure GetAccountno(pCardPan in varchar, pCardMBR in number);
  procedure GetFIO;
  PROCEDURE GetTransactions;
  procedure GenerateCreditStatement;
  procedure GenerateDebitStatement;
  procedure PrintTitle;
  PROCEDURE PrintTransactions;
  PROCEDURE PrintEndingTr;
  function  GetCreditInterest(pIdent in varchar) return number;
  function  GetOverdraft return number;
  PROCEDURE GetBeginBalance;
  FUNCTION  PrintOpenCard return boolean;
  FUNCTION  PrintNotClosedCard return boolean;
  FUNCTION  PrintAnyCard return boolean;
  FUNCTION  GetDailyLimit(pPan in varchar,pMbr in number) return number;
  PROCEDURE SetPanPrinted(pPan in varchar,pMbr in number);
  PROCEDURE PrintUnPrintedCards(pPrimary in number);
  PROCEDURE PrintHeader(pPrimary in number, pPan in varchar, pMBR in number);
  PROCEDURE SetPanTotal(pPan in varchar,pMBR in number,pCardTotal in number);
  PROCEDURE GetEntryGroup(pEntCode number,pMoney_Out number,pMoney_In number);
  PROCEDURE PrintCardTotal;
  PROCEDURE PrintRec(pText in varchar);
  function  GetNumber(pStr in varchar) return boolean;

procedure DialogExecProc
  (
    pWhat in char,
    pId   in number,
    pItem in varchar,
    pCmd  in number
  )
  is
    vId number;
  begin
    if pWhat = Dialog.wtDialogPre then
      svIdText := Text.Open('REPEG81');
    elsif pWhat = Dialog.wtDialogPost then
      Text.Close(svIdText);
    elsif pWhat = Dialog.wtItemPre then
      Dialog.SetItemDialog(pId, pItem, DialogParameters);
    elsif pWhat = Dialog.wtItemPost then
      vId := Dialog.GetItemDialog(pId, pItem);
      if Dialog.GetDialogCommand(vId) = Dialog.cmOk then
        Generate( Dialog.GetDate(vId,'BeginDate'),
                  Dialog.GetDate(vId,'EndDate'),
                  Dialog.Getchar(vId,'CardPan'),
                  Dialog.GetNumber(vId,'CardMBR'),
                  Dialog.GetChar(vId,'Path'),
                  Dialog.GetChar(vId,'FileName'),
                  Dialog.GetBool(vId,'UseArch'));
        Browser.ScrollerRefresh(pId,Dialog.FirstRecord);
      end if;
      Dialog.Destroy(vId);
      Dialog.SetItemDialog(pId, pItem, 0);
    end if;
  end;
procedure PrintRec(pText in varchar)
  is
  begin
    Text.Append(svIdText,pText);
    Term.WriteRecord(vFileHandle,pText);
  exception
   when OTHERS then
     s.say('REPEG81.PrintRec Error:'||sqlErrm);
     null;
  end;


function  DialogParameters
  return number is
    vId     number;
  begin
    vId := Dialog.New('Report parametrs', 0, 0, 54, 14);
    Dialog.InputDate(vId, 'BeginDate', 16,2, 'Enter begining date', 'Date from ');
    Dialog.PutDate(vId,'BeginDate', Seance.GetOperdate);
    Dialog.InputDate(vId, 'Enddate', 32,2, 'Enter ending date', ' to ');
    Dialog.PutDate(vId,'EndDate', Seance.GetOperdate);

    Dialog.InputChar(vId, 'CardPan', 14,4, 20,'Enter Card number (Press <Enter> - to find card)',null,'Card No:');
    Dialog.InputInteger(vId,'CardMBR',41,4 ,'Enter Card member number',1,'MBR:');
    Dialog.SetItemPost(vId, 'CardPan','REPEG81.DialogParametersProc');

    Dialog.StaticText(vId, 3,6, 'Unloading path: ');
    Dialog.InputChar(vId, 'Path', 14, 7, 30, 'Select directory',null,'Directory:');
    Dialog.SetItemAttributies(vId,'Path',Dialog.UpdateOFF);
    Dialog.Button(vId, 'Save', 46, 7, 4, '...', 0, 0, '');
    Dialog.InputChar(vId, 'FileName', 14, 8, 20, 'Enter file name',null, 'File name:');
    Dialog.SetItemPre(vId, 'Save', 'REPEG81.DialogParametersProc', Dialog.ProcType);
    Dialog.InputCheck(vId, 'UseArch',12, 10, 38, '','Use archive tables');

    Dialog.Button(vId, 'Ok', 13, 14, 8, '  Ok', Dialog.cmOk, 0, 'Generate report');
    Dialog.Button(vId, 'Cancel', 23, 14, 9, '  Cancel', Dialog.cmCancel, 0, 'Cancel report generation');
    Dialog.SetItemAttributies(vId, 'Ok', Dialog.DefaultOn);
    Dialog.SetDialogPre(vId,'REPEG81.DialogParametersProc');
    Dialog.SetDialogValid(vId, 'REPEG81.DialogParametersProc');
    return vId;
  end;
procedure DialogParametersProc
  (
    pWhat in char,
    pId   in number,
    pItem in varchar,
    pCmd  in number
  )
  is
    vContrTypStr varchar(300);
    vId          number;
    vBeginDate   date;
    vEndDate     date;
    vCardPan     varchar2(20);
    vCardMBR     number;
    vRet         number;
    vFN          varchar(250);
    vLegthPath   number;
    vPathString  varchar2(250);
    vFileName    varchar2(250);
  begin
    if pWhat = Dialog.wtDialogPre then
      Dialog.PutChar(pId, 'Path', DataMember.GetChar(Object.GetType('BRANCH'), sBranch, 'REPEG81_PATH'));
      Dialog.PutChar(pId, 'FileName', DataMember.GetChar(Object.GetType('BRANCH'), sBranch, 'REPEG81_FILENAME'));
    elsif pWhat = Dialog.wtItemPre then
      if pItem = 'SAVE' then
        vFN := Term.FileDialog('Enter path ',false,null, null, true);
        if vFN is not null then
          Dialog.PutChar(pId,'Path',vFN);
        end if;
      end if;
    elsif pWhat = Dialog.wtItemPost then
      if pItem = 'CARDPAN' then
          sScrollerMode := sScrollerFindCard;
          vId       := Finder.MainDialog(Finder.OT_CARD);
          Dialog.ExecDialog(vId);
          vCardPan := Finder.GetFinderResult(vId).Pan;
          vCardMBR := Finder.GetFinderResult(vId).MBR;
          Dialog.PutChar(pId, 'CARDPAN', vCardPan);
          Dialog.PutNumber(pId,'CARDMBR', vCardMBR);
          sScrollerMode := sScrollerText;
          Dialog.Destroy(vId);
          Dialog.GoItem(pId, 'OK');
      end if;
    elsif pWhat = Dialog.wtDialogValid then
      vBeginDate := Dialog.GetDate(pId,'BeginDate');
      vEndDate := Dialog.GetDate(pId,'EndDate');
      if vBeginDate is null then
        Dialog.SetHotHint(pId,'Error: Begining date not entered!');
        Dialog.GoItem(pId,'BeginDate');
        return;
      end if;
      if vEndDate is null then
        Dialog.SetHotHint(pId,'Error: Ending date not entered!');
        Dialog.GoItem(pId,'EndDate');
        return;
      end if;
      if Dialog.GetDate(pId, 'EndDate') < Dialog.GetDate(pId, 'BeginDate')  then
        Dialog.SetHotHint(pId, 'Error: Invalid date period');
        Dialog.GoItem(pId, 'BeginDate');
        return;
      end if;
      vCardPan := Dialog.GetChar(pId,'CARDPAN');
      vCardMBR := Dialog.GetNumber(pId,'CARDMBR');
      if vCardPan is null then
        Dialog.SetHotHint(pId,'Error: Card Number not entered!');
        Dialog.GoItem(pId,'CARDPAN');
        return;
      end if;
      if vCardMBR is null then
        Dialog.SetHotHint(pId,'Error: Card Member Number not entered!');
        Dialog.GoItem(pId,'CARDMBR');
        return;
      end if;
      begin
        select PAN, MBR
          into vCardPan,vCardMBR
          from tcard
         where branch = sBranch and
           pan = vCardPan and
           mbr = vCardMBR;
        exception
         when OTHERS then
           Dialog.SetHotHint(pId, 'Error: Invalid Card Number');
           Dialog.GoItem(pId, 'CARDPAN');
           return;
      end;
      vPathString := Dialog.GetChar(pId, 'Path');
      s.say('sPathString= '||vPathString);
      vLegthPath := Length(vPathString);
      if vPathString is null then
        Dialog.SetHotHint(pId,'Error: Unloading path not entered!~');
        Dialog.GoItem(pId,'Path');
        return;
      end if;
      vFileName :=  Dialog.GetChar(pId,'FILENAME');
      if vFileName is null then
        Dialog.SetHotHint(pId,'Error: File name not entered!~');
        Dialog.GoItem(pId,'Path');
        return;
      end if;
      vRet :=  DataMember.SetChar(Object.GetType('BRANCH'), sBranch, 'REPEG81_PATH', vPathString, null);
      vRet :=  DataMember.SetChar(Object.GetType('BRANCH'), sBranch, 'REPEG81_FILENAME', vFileName, null);
    end if;
  end;
procedure GetContract(pCardPan in varchar, pCardMBR in number)
 is
 begin
   vContract := null;
   vIdClient := -1;
   sPanCounter := 0;
   Begin
      select c.No, c.idClient
        into vContract, vidClient
        from tcontractitem ci, tcontract c
       where ci.branch = sBranch and
         ci.itemtype = 2 and
         ci.key = pCardPan||pCardMBR and
         c.branch  = ci.branch and
         c.no = ci.no and
         rownum = 1;
      exception
       when NO_DATA_FOUND then
         null;
   end;
   if vContract is not null then
      for i in (
          select c.Pan,c.idclient,c.CRD_STAT,c.mbr
            from tcontractitem ci, tcard c
           where ci.branch = sBranch and
             ci.no = vContract and
             ci.itemtype = 2 and
             c.branch = ci.branch and
             c.pan||c.mbr = ci.key
          order by 1 desc) loop
        sPanCounter := sPanCounter + 1;
        sPanArr(sPanCounter).Pan := i.Pan;
        sPanArr(sPanCounter).MBR := i.MBR;
        sPanArr(sPanCounter).Idclient := i.IdClient;
        sPanArr(sPanCounter).Status   := i.Crd_Stat;
        sPanArr(sPanCounter).Printed  := 0;
        sPanArr(sPanCounter).Total    := 0;
      end loop;
   end if;
 exception
   when others then
     s.say('REPEG81.GetContract Error:'||sqlErrm);
     null;
 end;
procedure GetAccountno(pCardPan in varchar, pCardMBR in number)
 is
 begin
   vAccountno := null;
   vIdClient := -1;
   sPanCounter := 0;
   Begin
      select a.AccountNo, a.idClient
        into vAccountno, vIdClient
        from tacc2card ac, taccount a
       where ac.branch = sBranch and
         ac.pan = pCardPan and
         ac.mbr = pCardMBR and
         a.branch = ac.branch and
         a.accountno = ac.accountno and
         rownum = 1;
      exception
       when NO_DATA_FOUND then
         null;
   end;
  if vAccountno is not null then
     for i in (
         select c.Pan,c.idclient,c.CRD_STAT,c.mbr
           from tacc2card ac, tcard c
          where ac.branch = sBranch and
            ac.accountno = vAccountno and
            c.branch = ac.branch and
            c.pan = ac.pan and
            c.mbr = ac.mbr
         order by 1 desc) loop
       sPanCounter := sPanCounter + 1;
       sPanArr(sPanCounter).Pan := i.Pan;
       sPanArr(sPanCounter).MBR := i.MBR;
       sPanArr(sPanCounter).Idclient := i.IdClient;
       sPanArr(sPanCounter).Status   := i.Crd_Stat;
       sPanArr(sPanCounter).Printed  := 0;
       sPanArr(sPanCounter).Total    := 0;
     end loop;
   end if;
 exception
   when others then
     s.say('REPEG81.GetAccountno Error:'||sqlErrm);
     null;
 end;
procedure Generate
  (pBeginDate in date,
   pEndDate in date,
   pCardPan in varchar,
   pCardMBR in number,
   pPath in varchar,
   pFileName in varchar,
   pUseArchive in boolean)
 is
 begin
  Text.Clear(svIdText);
  vFromDate    := nvl(pBeginDate, Seance.GetOperdate);
  vToDate      := nvl(pEndDate, Seance.GetOperDate);
  --vUseArchive  := nvl(pUseArchive,false);
  if nvl(pUseArchive,false) then
    vprizArh := 1;
  else
    vprizArh := 0;
  end if;
  s.say('vprizArh '||vprizArh,2);
  vFileHandle := Term.FileOpenWrite(pPath||pFileName||'.txt',pEnc=>'UTF8');
  GetContract(pCardPan, pCardMBR);
  vPan := ' ';
  vMBR := 0;
  if vContract is not null then
    GenerateCreditStatement;
  else
    GetAccountno(pCardPan, pCardMBR);
    GenerateDebitStatement;
  end if;
  Term.Close(vFileHandle);
  commit;
 exception
   when others then
     s.say('REPEG81.Generate Error:'||sqlErrm);
     Term.Close(vFileHandle);
     null;
 end;
procedure GetFIO
 is
 vCity    number := -1;
 vRegion  number := -1;
 vCountry number := -1;
 begin
   vFIO         := ' ';
   vAddressCont := ' ';
   vCityCont    := ' ';
   vRegionCont  := ' ';
   vCountryCont := ' ';
   vIdClient    := -1;
   vAbbreviation := ' ';
   begin
     select cl.FIO, cl.AddressCont, cl.CityCont,cl.RegionCont,cl.CountryCont, a.idclient, r.Abbreviation
       into vFio, vAddressCont, vCity, vRegion, vCountry, vIdClient, vAbbreviation
       from taccount a, tclientpersone cl, treferencecurrency r
      where a.branch = sBranch and
        a.accountno = vAccountno and
        cl.branch  = a.branch and
        cl.idClient = a.idClient and
        r.branch = a.branch and
        r.currency = a.currencyno;
     exception
      when NO_DATA_FOUND then
        null;
   end;
   vFio := nvl(vFio,' ');
   vCity := nvl(vCity,-1);
   vRegion := nvl(vRegion,-1);
   vCountry := nvl(vCountry,-1);
   vCityCont    := nvl(ReferenceCity.GetName(vCountry,vRegion,vCity),' ');
   vRegionCont  := nvl(ReferenceRegion.GetName(vCountry,vRegion),' ');
   vCountryCont := nvl(ReferenceCountry.GetName(vCountry),' ');
   vIdClient    := nvl(vIdClient,-1);
   vAbbreviation := nvl(vAbbreviation,' ');
 exception
   when others then
     s.say('REPEG81.GetFIO Error:'||sqlErrm);
     null;
 end;
  procedure GetTransactions is
    vFirstline  boolean := true;
    vPreviosPan varchar2(20) := null;
    vPreviosMBR number := 0;
    vPrimaryCard number := 0;
    vCardTotal   number := 0;
  Begin
    vTotalMoney_Out := 0;
    vTotalMoney_In  := 0;
    vWithdrawal := 0;
    vPurchases  := 0;
    vOther := 0;
    vPayment_Credits := 0;
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
      Trans.Description,
      Trans.Ident,
      Trans.DebitEntCode,
      Trans.Termlocation,
      Trans.Detail,
      Trans.OrgCurrency,
      Trans.AeDocno,
      Trans.eValueDate,
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
            e.Debitaccount = vAccountno and
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
            e.Creditaccount = vAccountno and
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
            e.Debitaccount = vAccountno and
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
            e.Creditaccount = vAccountno and
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
      Trans.Ident,
      Trans.Debitentcode,
      Trans.TermLocation,
      nvl(c.Pan,vpan),
      nvl(c.MBR,vMbr),
      Trans.Detail,
      Trans.OrgCurrency,
      Trans.AeDocno,
      decode(c.idclient,vIdClient,1,
              decode(c.idclient,null,1,-1)),
      Trans.eValueDate
      order by 1 desc,2 desc,3,4,5,6) loop
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
      if vMoney_Out != 0 or vMoney_In != 0 then
        if vContract is null then
          vPayment_Credits := vPayment_Credits + vMoney_In;
          vAll := vAll + vMoney_Out;
          if i.Ident in ('OPEXT_A10','OPEXT_B15','OPEXT_B23','OPEXT_V15','OPEXT_V23','ACC_ADD','ACC_SUB') and
            vMoney_In = 0 then
            vWithdrawal := vWithdrawal + vMoney_Out;
          elsif (instr(i.Ident,'OPCOM') != 0 or instr(i.Ident,'FEE') != 0) and
            vMoney_In = 0 then
            vOther := vOther+ vMoney_Out;
          end if;
        else
          if (instr(i.ident,'CHARGE_INTEREST_GROUP') = 0) then
            GetEntryGroup(i.DebitEntCode,vMoney_Out,vMoney_In);
          end if;
        end if;
        if vFirstline then
          vPreviosPan  := nvl(i.Pan,' ');
          vPreviosMBR  := nvl(i.MBR,0);
          vPrimaryCard := i.PrimaryCard;
          vFirstline := false;
          if nvl(vPan,' ') != nvl(i.Pan,' ') then
            PrintHeader(i.PrimaryCard,i.Pan,i.MBR);
          end if;
        end if;
        if vPreviosPan != nvl(i.Pan,' ') or vPreviosMBR  != nvl(i.MBR,0) then
          SetPanTotal(vPreviosPan,vPreviosMBR,vCardTotal);
          PrintHeader(i.PrimaryCard,i.Pan,i.MBR);
          vCardTotal := 0;
        end if;
        PrintTransactions;
        vPreviosPan  := nvl(i.Pan,' ');
        vPrimaryCard := i.PrimaryCard;
      end if;
      vCardTotal    := vCardTotal + vMoney_In - vMoney_Out;
    end loop;
    if vContract is null then
      vPurchases := vAll-vWithdrawal-vOther;
    end if;
    SetPanTotal(vPreviosPan,vPreviosMBR,vCardTotal);
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
    elsif i.CashAdvance = 0 then
      vPurchases := vPurchases + pMoney_out - pMoney_in;
    elsif i.CashAdvance = 2 then
      vPayment_credits := vPayment_credits - pMoney_out + pMoney_in;
    elsif i.CashAdvance = 3 then
      vOther := vOther + pMoney_out - pMoney_in;
    end if;
  found := true;
  end loop;
 if not found then
    vOther := vOther + pMoney_out - pMoney_in;
 end if;
 end;

PROCEDURE SetPanTotal(pPan in varchar,pMBR in number,pCardTotal in number)
  is
  begin
    for i in 1..sPanCounter loop
      if sPanArr(i).Pan = pPan and sPanArr(i).MBR = pMBR then
        sPanArr(i). Total := pCardTotal;
        exit;
      end if;
    end loop;
  end;


PROCEDURE PrintHeader(pPrimary in number, pPan in varchar, pMBR in number)
  is
  begin
    if pPan = ' ' or pPan is null then
      if pPrimary = 1 then
        PrintRec(Service.rightpad(' ',24)||'Transaction for Primary Card no. ');
      else
        PrintUnPrintedCards(1);
        PrintRec(Service.rightpad(' ',24)||'Transaction for Supplementary Card no. ');
      end if;
    else
      SetPanPrinted(pPan, pMBR);
      if pPrimary = 1 then
        PrintRec(Service.rightpad(' ',24)||'Transaction for Primary Card no. '||nvl(pPAN,' ')||'-'||ltrim(rtrim(to_char(nvl(pMBR,0)))));
      else
        PrintUnPrintedCards(1);
        PrintRec(Service.rightpad(' ',24)||'Transaction for Supplementary Card no. '||nvl(pPAN,' ')||'-'||ltrim(rtrim(to_char(nvl(pMBR,0)))));
      end if;

    end if;
  end;
procedure PrintUnPrintedCards(pPrimary in number)
 is
 begin
   for i in 1..sPanCounter loop
     if sPanArr(i).Printed = 0  then
       if pPrimary = 1 and sPanArr(i).IdClient = vIdClient then
         PrintRec(Service.rightpad(' ',24)||'Transaction for Primary Card no. '||nvl(sPanArr(i).PAN,' ')||'-'||ltrim(rtrim(to_char(nvl(sPanArr(i).MBR,0)))));
         sPanArr(i).Printed := 1;
       end if;
       if pPrimary != 1 and sPanArr(i).IdClient != vIdClient then
         PrintRec(Service.rightpad(' ',24)||'Transaction for Supplementary Card no. '||nvl(sPanArr(i).PAN,' ')||'-'||ltrim(rtrim(to_char(nvl(sPanArr(i).MBR,0)))));
         sPanArr(i).Printed := 1;
       end if;
     end if;
   end loop;
 end;
procedure GenerateCreditStatement
 is
 vLimDom number := 0;
 vLimInt number := 0;
 begin
  s.say(' CREDIT '||vContract);
    vDepositDOM := nvl(Contract.GetAccountNoByItemName(vContract,'ItemDepositDOM'),' ');
    vDepositINT := nvl(Contract.GetAccountNoByItemName(vContract,'ItemDepositINT'),' ');
    s.say(' CREDIT  '||vContract||' vDepositDOM '||vDepositDOM||' vDepositINT '||vDepositINT,2);
    SCH_Customer.GetAvailableLimit(vContract,vLimDom,vLimInt);
    vWithdrawal       := 0;
    vPurchases        := 0;
    vCredit_Interest  := 0;
    vOther            := 0;
    vPayment_Credits  := 0;
    vAll              := 0;
    vOverdraft        := 0;
   for nn in 1..2 loop
     if nn = 1 then
       vAccountno := vDepositDOM;
       vAvailable_Credit := vLimDom;
     else
       vAccountno := vDepositINT;
       vAvailable_Credit := vLimInt;
     end if;
     if vAccountno != ' ' then
       GetFIO;
       GetBeginBalance;
       vOverdraft := GetOverdraft;
       vCredit_Interest :=  GetCreditInterest('CHARGE_INTEREST_GROUP_');
       vMinimum_Payment := nvl(SCH_Customer.GetMinPayment(vContract,vAccountno,vToDate),0);
       PrintTitle;
       GetTransactions;
       PrintEndingTr;
     end if;
   end loop ;
 exception
   when others then
     s.say('REPEG81.GenerateCreditStatement Error:'||sqlErrm);
     null;
 end;
procedure GenerateDebitStatement
 is
 begin
   vWithdrawal       := 0;
   vPurchases        := 0;
   vCredit_Interest  := 0;
   vOther            := 0;
   vPayment_Credits  := 0;
   vAll              := 0;
   GetFIO;
   s.say(' DEBIT  '||vAccountno,2);
   vOverdraft := GetOverdraft;
   vMinimum_Payment := null;
   vPrevious_balance := 0;
   if vPrevious_balance < 0 then
     vAvailable_Credit := vOverdraft - Abs(vPrevious_balance);
   else
     vAvailable_Credit := vOverdraft;
   end if;
   GetBeginBalance;
   PrintTitle;
   GetTransactions;
   PrintEndingTr;
 exception
   when others then
     s.say('REPEG81.GenerateDebitStatement Error:'||sqlErrm);
     null;
 end;
function GetCreditInterest(pIdent in varchar) return number
  is
  vIdent    varchar2(40);
  vInterest number;
  Begin
    vIdent := pIdent;
    vInterest := 0;
    for i in (
      select
      sum(Money_Out) Money_Out
      from(
      select
      sum(e.value) Money_Out
      from tEntry e,tDocument d,tReferenceEntry re
      where e.Branch       = sbranch and
            e.Debitaccount = vAccountNo and
            e.Value       != 0 and
            re.Branch      = e.branch and
            re.code        = e.debitentcode and
            Instr(re.Ident,vIdent) != 0 and
            d.branch       = e.Branch and
            d.docno        = e.docno and
            d.NewDocNo     is null and
            d.OpDate       between vFromDate and vToDate
      union all
      select
      sum(e.value) Money_Out
      from tEntryarc e,tDocumentarc d,tReferenceEntry re
      where e.Branch       = sbranch and
            e.Debitaccount = vAccountNo and
            e.Value       != 0 and
            re.Branch      = e.branch and
            re.code        = e.debitentcode and
            Instr(re.Ident,vIdent) != 0 and
            d.branch       = e.Branch and
            d.docno        = e.docno and
            d.NewDocNo     is null and
            d.OpDate       between vFromDate and vToDate and
            vprizArh = 1 ) ) loop
      vInterest := i.Money_Out;
    end loop;
    return nvl(vInterest,0);
  End;
function GetOverdraft return number
  is
    vOverdraft number := 0;
  begin
    begin
      Select Overdraft
      into vOverdraft
      from tAccount
      where branch  = sBranch and
        accountno = vAccountno and
        rownum = 1;
    exception
      when NO_DATA_FOUND then
      null;
    end;
    return nvl(vOverdraft,0);
  end;
procedure GetBeginBalance
  is
  vRemain number := 0;
  Begin
    vPrevious_balance := 0;
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
        a.AccountNo = vAccountno
      union all
      select
        0 RemDep,
        -e.Value ValDep
      from tEntry e, tDocument d
      where e.Branch   = sBranch and
        e.DebitAccount = vAccountno and
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
        e.CreditAccount = vAccountno and
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
        e.DebitAccount = vAccountno and
        d.Branch       = e.Branch and
        d.DocNo        = e.DocNo and
        d.OpDate      >= vFromDate and
        d.NewDocNo is null and
        vPrizArh = 1
      union all
      select
        0 RemDep,
        e.Value ValDep
      from tEntryarc e, tDocumentarc d
      where e.Branch    = sBranch and
        e.CreditAccount = vAccountno and
        d.Branch        = e.Branch and
        d.DocNo         = e.DocNo and
        d.OpDate       >= vFromDate and
        d.NewDocNo is null and
        vPrizArh = 1
      ) T;
    vPrevious_balance := nvl(vRemain,0);
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
   s.say('Repeg81.GetNumber Error:'||sqlErrm);
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
    s.say(vAddressLine);
    if nvl(length(ltrim(rtrim(vAddressLine))),0) <= vLengthKol then
      PrintRec(Service.RightPad('Address', 18,' ')||':'|| Service.LeftPad(' ', 5,' ')||Service.rightpad(vAddressLine,vLengthKol));
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
                      PrintRec( Service.RightPad('Address', 18,' ')||':'|| Service.LeftPad(' ', 5,' ')||
                        Service.rightpad(Service.rightpad(chr(55473)||chr(55682)||chr(55685) ||' '||substr(vAddressLine,vPos1,vKol1-vPos1),vLengthKol),vLengthKol));
                      vFirstLine := false;
                   else
                      PrintRec( Service.LeftPad(' ', 24,' ')||Service.rightpad(Service.rightpad(chr(55473)||chr(55682)||chr(55685) ||' '||substr(vAddressLine,vPos1,vKol1-vPos1),vLengthKol),vLengthKol));
                   end if;
                else
                  if vFirstLine then
                    vFirstLine := false;
                    PrintRec( Service.RightPad('Address', 18,' ')||':'|| Service.LeftPad(' ', 5,' ')||
                      Service.rightpad(Service.rightpad(substr(vAddressLine,vPos1,vKol1-vPos1),vLengthKol),vLengthKol));
                  else
                    PrintRec( Service.LeftPad(' ', 24,' ')||Service.rightpad(Service.rightpad(substr(vAddressLine,vPos1,vKol1-vPos1),vLengthKol),vLengthKol));
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
                    PrintRec( Service.RightPad('Address', 18,' ')||':'|| Service.LeftPad(' ', 5,' ')||Service.rightpad(Service.rightpad(chr(55473)||chr(55682)||chr(55685) ||' '||substr(vAddressLine,vPos1,vKol2-vPos1),vLengthKol),vLengthKol));
                  else
                    PrintRec( Service.LeftPad(' ',24 ,' ')||Service.rightpad(Service.rightpad(chr(55473)||chr(55682)||chr(55685) ||' '||substr(vAddressLine,vPos1,vKol2-vPos1),vLengthKol),vLengthKol));
                  end if;
                else
                  if vFirstLine then
                    PrintRec( Service.RightPad('Address', 18,' ')||':'|| Service.LeftPad(' ', 5,' ')||Service.rightpad(Service.rightpad(substr(vAddressLine,vPos1,vKol2-vPos1),vLengthKol),vLengthKol));
                    vFirstLine := false;
                  else
                    PrintRec( Service.LeftPad(' ',24 ,' ')||Service.rightpad(Service.rightpad(substr(vAddressLine,vPos1,vKol2-vPos1),vLengthKol),vLengthKol));
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
              PrintRec(Service.RightPad('Address', 18,' ')||':'||
              Service.LeftPad(' ', 5,' ')||Service.rightpad(chr(55473)||chr(55682)||chr(55685) ||' '||ltrim(rtrim(substr(vAddressLine,vPos1))),vLengthKol));
              vFirstLine := false;
            else
              PrintRec(Service.LeftPad(' ', 24,' ')||Service.rightpad(chr(55473)||chr(55682)||chr(55685) ||' '||ltrim(rtrim(substr(vAddressLine,vPos1))),vLengthKol));
            end if;
          else
            if vFirstLine then
              vFirstLine := false;
              PrintRec( Service.RightPad('Address', 18,' ')||':'|| Service.LeftPad(' ', 5,' ')||
                Service.rightpad(ltrim(rtrim(substr(vAddressLine,vPos1))),vLengthKol));
            else
              PrintRec( Service.LeftPad(' ', 24,' ')||Service.rightpad(ltrim(rtrim(substr(vAddressLine,vPos1))),vLengthKol));
            end if;
          end if;
        end if;
    end if;
    for i in vTranCount..2 loop
      PrintRec(' ');
    end loop;
  exception
   when OTHERS then
     s.say('Repeg81.PrintAddress Error:'||sqlErrm);
  end;


FUNCTION PrintOpenCard return boolean
  is
  vFound boolean := false;
  vFistFound boolean := true;
  vCardList varchar2(3000) := null;
  begin
    vCardList := '';
    for i in 1 .. sPanCounter loop
      if sPanArr(i).Status = ReferenceCrd_Stat.CARD_OPEN and sPanArr(i).idClient = vIdClient then
          PrintRec(Service.RightPad('Card Number', 18,' ')||':'||Service.LeftPad(' ', 5,' ')||
            nvl(sPanArr(i).PAN,' ')||'-'||ltrim(rtrim(to_char(nvl(sPanArr(i).MBR,0)))));
          vFistFound := false;
          vPan := sPanArr(i).PAN;
          vMBR := sPanArr(i).MBR;
          vFound := true;
          exit;
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
      if sPanArr(i).Status != ReferenceCrd_Stat.CARD_CLOSED and sPanArr(i).idClient = vIdClient then
        PrintRec(Service.RightPad('Card Number', 18,' ')||':'||Service.LeftPad(' ', 5,' ')||
          nvl(sPanArr(i).PAN,' ')||'-'||ltrim(rtrim(to_char(nvl(sPanArr(i).MBR,0)))));
        vFistFound := false;
        vPan := sPanArr(i).PAN;
        vMBR := sPanArr(i).MBR;
        vFound := true;
        exit;
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
          PrintRec(Service.RightPad('Card Number', 18,' ')||':'||Service.LeftPad(' ', 5,' ')||
            nvl(sPanArr(i).PAN,' ')||'-'||ltrim(rtrim(to_char(nvl(sPanArr(i).MBR,0)))));
          vFistFound := false;
          vFound := true;
          vPan := sPanArr(i).PAN;
          vMBR := sPanArr(i).MBR;
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
      if vListLim(i).Code = 3 and vListLim(i).PeriodType = ReferenceLimit.PeriodType_NONE then
        vLimit := vListLim(i).Value;
      end if;
    end loop;
    return vLimit;
  end;
procedure SetPanPrinted(pPan in varchar, pMBR in number)
 is
   vFound boolean := false;
 begin
   for i in 1..sPanCounter loop
     if sPanArr(i).Pan = pPan and sPanArr(i).MBR = pMBR then
        sPanArr(i).Printed := 1;
        exit;
     end if;
   end loop;
 end;
procedure PrintTitle
 is
 vLimit number;
 begin
   PrintRec( Service.RightPad('Corporate Name', 18,' ')||':'|| Service.LeftPad(' ', 5,' ')||'NSGB');
   PrintRec( Service.RightPad('Full Name', 18,' ')||':'|| Service.LeftPad(' ', 5,' ')||vFIO);
   PrintAddress(vAddressCont);
   PrintRec( Service.RightPad('City', 18,' ')||':'|| Service.LeftPad(' ', 5,' ')||vCityCont);
   PrintRec( Service.RightPad('Region', 18,' ')||':'|| Service.LeftPad(' ', 5,' ')||vRegionCont);
   PrintRec( Service.RightPad('Country', 18,' ')||':'|| Service.LeftPad(' ', 5,' ')||vCountryCont);
   PrintRec(' ');
   PrintRec(' ');
   PrintRec( Service.RightPad('Statement Date', 18,' ')||':'|| Service.LeftPad(' ', 5,' ')||to_date(vToDate,'dd/mm/yy'));
   if not PrintOpenCard then
     if not PrintNotClosedCard then
       if not PrintAnyCard then
         PrintRec(Service.RightPad('Card Number', 18,' ')||':'||Service.LeftPad(' ', 5,' ')||' ');
       end if;
     end if;
   end if;
   --vLimit := GetDailyLimit(vPan,vMBR);
   PrintRec(Service.RightPad('Card Limit', 18,' ')||':'||Service.LeftPad(' ', 5,' ')||
     ltrim(rtrim(vAbbreviation))||' '||nvl(ltrim(to_char(vOverdraft,'9999990.00')),' '));
   if vAvailable_Credit < 0 then
     vAvailable_Credit := 0;
   end if;
   PrintRec(Service.RightPad('Available Limit', 18,' ')||':'|| Service.LeftPad(' ', 5,' ')||
     ltrim(rtrim(vAbbreviation))||' '||ltrim(to_char(vAvailable_Credit,'999999999990.00')));
   if vMinimum_Payment is not null then
     PrintRec(Service.RightPad('Minimum Payment', 18,' ')||':'|| Service.LeftPad(' ', 5,' ')||
       ltrim(rtrim(vAbbreviation))||' '||ltrim(to_char(vMinimum_Payment,'999999999990.00')));
   else
     PrintRec( Service.RightPad('Minimum Payment', 18,' ')||':');
   end if;
   PrintRec(' ');
   PrintRec(' ');
   PrintRec(  '----------------------------------------------------------------------------------------------------------');
   PrintRec(  'Date         Ref        Transaction Description                                              Amount in '||
     Service.LeftPad(nvl(vAbbreviation,' '), 3,' '));
   PrintRec(  '----------------------------------------------------------------------------------------------------------');
   if vPan != ' ' then
     PrintRec('                        Transaction for Primaty Card no. '||vPan||'-'||ltrim(rtrim(to_char(vMBR))));
   else
     PrintRec('                        Transaction for Primaty Card no. ');
   end if;
   SetPanPrinted(vPan,vMBR);
   if vPrevious_balance >= 0 then
     PrintRec(Service.rightpad(' ',24)||Service.rightpad('Previous Balance',63)|| Service.LeftPad(to_char(vPrevious_balance,'99999999990.00'), 16,' ')||' ');
   else
     PrintRec(Service.rightpad(' ',24)||Service.rightpad('Previous Balance',63)|| Service.LeftPad(to_char(Abs(vPrevious_balance),'99999999990.00'),16)||' DB');
   end if;
 end;
procedure PrintTransactions
-- перенос строки без разрыва слов
  is
  vTextStr varchar(500) := null;
  vRateStr varchar(40) := null;
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
  vstr       varchar2(300) := null;
  Begin
    vTotalMoney_Out := vTotalMoney_Out + vMoney_Out;
    vTotalMoney_In  := vTotalMoney_In  + vMoney_In;
   -- vLengthKol := 43;
    vLengthKol := 63;
    if vMoney_In > 0 then
      vstr := Service.LeftPad(to_char(vMoney_Out+vMoney_In,'99999999990.00'), 16,' ')||'  ';
    else
      vstr := Service.LeftPad(to_char(vMoney_Out+vMoney_In,'99999999990.00'), 16,' ')||' DB';
    end if;
    if vOrgValue > 0 then
      vRate := (vMoney_In + vMoney_Out)/vOrgValue;
    vRateStr := ' RATE '||to_char(vRate,'90.0000')||' '||ltrim(rtrim(vAbbreviation))||' '||ltrim(rtrim(vOrgAbbr));
    end if;
    if vtermlocation != ' ' and vOrgValue != 0 then
      vTextStr := vDescription||' at '||vTermLocation||' Original currency Amount '||
             ltrim(rtrim(to_char(vOrgValue,'9999999990.00')))||' ('||
             ltrim(rtrim(vOrgAbbr))||')'||vRateStr;
      if vAeDocno != 0 then
        vTextStr := vTextStr||' '||to_char(vevaluedate,'dd/mm/yy');
      end if;
    elsif vtermlocation = ' ' and vOrgValue != 0 then
      vTextStr := vDescription||' Original currency Amount '||
             ltrim(rtrim(to_char(vOrgValue,'9999999990.00')))||' ('||
             ltrim(rtrim(vOrgAbbr))||')'||vRateStr;
      if vAeDocno != 0 then
        vTextStr := vTextStr||' '||to_char(vevaluedate,'dd/mm/yy');
      end if;
    elsif vtermlocation != ' ' and vOrgValue = 0 then
      vTextStr := vDescription||' at '||vTermLocation;
      if vAeDocno != 0 then
        vTextStr := vTextStr||' '||to_char(vevaluedate,'dd/mm/yy');
      end if;
    elsif vtermlocation = ' ' and vOrgValue = 0 then
      vTextStr := vDescription;
      if vAeDocno != 0 then
        vTextStr := vTextStr||' '||to_char(vevaluedate,'dd/mm/yy');
      end if;
    else
      vTextStr := ltrim(rtrim(vDescription));
      if vAeDocno != 0 then
        vTextStr := vTextStr||' '||to_char(vevaluedate,'dd/mm/yy');
      end if;
    end if;
    if nvl(length(ltrim(rtrim(vTextStr))),0) <= vLengthKol then
      vTranCount := vTranCount + 1;
      if vTranCount < 4 then
        PrintRec(Service.rightpad(to_char(vValuedate,'dd/mm/yyyy'),13)||
          Service.rightpad(to_char(vDocno),11)||
          Service.rightpad(nvl(rtrim(ltrim(vTextStr)),' '),vLengthKol)||
          vstr);
      end if;
    else
      vKolCount :=length(ltrim(rtrim(vTextStr))) / vLengthKol;
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
              vTranCount := vTranCount + 1;
              if vTranCount < 4 then
                PrintRec(Service.rightpad(to_char(vValuedate,'dd/mm/yyyy'),13)||
                  Service.rightpad(to_char(vDocno),11)||
                  Service.rightpad(Service.rightpad(substr(vTextStr,vPos1,vKol1-vPos1),vLengthKol),vLengthKol)||
                  vstr);
              end if;
              vPrizn := 5;
            else
              vTranCount := vTranCount + 1;
              if vTranCount < 4 then
                 PrintRec(Service.rightpad(' ',24)||
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
            vTranCount := vTranCount + 1;
            if vTranCount < 4 then
              PrintRec(Service.rightpad(to_char(vValuedate,'dd/mm/yyyy'),13)||
                Service.rightpad(to_char(vDocno),11)||
                Service.rightpad(Service.rightpad(substr(vTextStr,vPos1,vKol2-vPos1),vLengthKol),vLengthKol)||
                vstr);
            end if;
            vPrizn := 5;
          else
            vTranCount := vTranCount + 1;
            if vTranCount < 4 then
              PrintRec(Service.rightpad(' ',24)||
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
        vTranCount := vTranCount + 1;
        if vTranCount < 4 then
          PrintRec(Service.rightpad(to_char(vValuedate,'dd/mm/yyyy'),13)||
            Service.rightpad(to_char(vDocno),11)||
            Service.rightpad(ltrim(rtrim(substr(vTextStr,vPos1))),vLengthKol)||
            vstr);
        end if;
        vPrizn := 5;
      else
        vTranCount := vTranCount + 1;
        if vTranCount < 4 then
          PrintRec(Service.rightpad(' ',24)||
            Service.rightpad(ltrim(rtrim(substr(vTextStr,vPos1))),vLengthKol));
        end if;
      end if;
    end if;
  End;
procedure PrintCardTotal
 is
 begin
   for i in 1..sPanCounter loop
     if sPanArr(i).idClient = vIdclient then
       if sPanArr(i).Total >= 0 then
         PrintRec( Service.LeftPad(' ', 24,' ')||
         Service.RightPad('Total for Primary Card '||sPanArr(i).Pan||'-'||sPanArr(i).MBR, 63,' ')||
         Service.LeftPad(to_char(sPanArr(i).Total,'99999999990.00' ),16 ,' '));
       else
         PrintRec(Service.LeftPad(' ', 24,' ')||
         Service.RightPad('Total for Primary Card '||sPanArr(i).Pan||'-'||sPanArr(i).MBR, 63,' ')||
         Service.LeftPad(to_char(abs(sPanArr(i).Total),'99999999990.00' ),16 ,' ')||' DB');
       end if;
       sPanArr(i).Total := 0;
     end if;
   end loop;
   for i in 1..sPanCounter loop
     if sPanArr(i).idClient != vIdclient then
       if sPanArr(i).Total >= 0 then
         PrintRec( Service.LeftPad(' ', 24,' ')||
         Service.RightPad('Total for Supplementary Card '||sPanArr(i).Pan||'-'||sPanArr(i).MBR, 63,' ')||
         Service.LeftPad(to_char(sPanArr(i).Total,'99999999990.00' ),16 ,' '));
       else
         PrintRec(Service.LeftPad(' ', 24,' ')||
         Service.RightPad('Total for Supplementary Card '||sPanArr(i).Pan||'-'||sPanArr(i).MBR, 63,' ')||
         Service.LeftPad(to_char(abs(sPanArr(i).Total),'99999999990.00' ),16 ,' ')||' DB');
       end if;
       sPanArr(i).Total := 0;
     end if;
   end loop;
 exception
   when others then
     s.say('REPEG81.PrintCardTotal Error:'||sqlErrm);
     null;
 end;
  procedure PrintEndingTr
  is
  vStrText varchar2(500) := null;
  Begin
   -- vTransCounter := vTransCounter + 1;
    PrintUnPrintedCards(1);
    PrintUnPrintedCards(2);
    PrintRec(' ');
    PrintCardTotal;
    if vTotalMoney_In-vTotalMoney_Out >= 0 then
      PrintRec(Service.rightpad(' ',24)||Service.rightpad('Total for all Cards',63)||
        Service.LeftPad(to_char(vTotalMoney_In-vTotalMoney_Out,'9999990.00'),16 ,' ')||'  ');
    else
      PrintRec(Service.rightpad(' ',24)||Service.rightpad('Total for all Cards',63)||
        Service.LeftPad(to_char(Abs(vTotalMoney_In-vTotalMoney_Out),'9999990.00'), 16,' ')||' DB');
    end if;
    PrintRec(' ');
    PrintRec('----------------------------------------------------------------------------------------------------------');
    PrintRec(' ');
    if vPrevious_balance >= 0 then
      PrintRec('Previous Balance  :'||ltrim(to_char(vPrevious_balance,'9999990.00'))||'  ');
    else
      PrintRec('Previous Balance  :'||ltrim(to_char(Abs(vPrevious_balance),'9999990.00'))||' DB');
    end if;
    PrintRec(  'Cash Withdraws    :'||ltrim(to_char(vWithdrawal,'9999990.00')));
    PrintRec(  'Purchases         :'||ltrim(to_char(vPurchases,'9999990.00')));
    PrintRec(  'Total Interest    :'||ltrim(to_char(vCredit_Interest,'9999990.00')));
    PrintRec(  'Charges           :'||ltrim(to_char(vOther,'9999990.00')));
    PrintRec(  'Payments/Credits  :'||ltrim(to_char(vPayment_Credits,'9999990.00')));
    if vPrevious_balance+vTotalMoney_In-vTotalMoney_Out >= 0 then
      PrintRec('Closing Balance   :'||ltrim(to_char(vPrevious_balance+vTotalMoney_In-vTotalMoney_Out,'9999990.00'))||'  ');
    else
      PrintRec('Closing Balance   :'||ltrim(to_char(Abs(vPrevious_balance+vTotalMoney_In-vTotalMoney_Out),'9999990.00'))||' DB');
    end if;
    PrintRec(' ');
    PrintRec('----------------------------------------------------------------------------------------------------------');
    PrintRec(' ');
    PrintRec(' ');
    PrintRec(' ');
    vWithdrawal       := 0;
    vPurchases        := 0;
    vCredit_Interest  := 0;
    vOther            := 0;
    vPayment_Credits  := 0;
    vAll              := 0;
  End;
function  GetCount
  return number is
  begin
    return Text.GetCount(svIdtext);
  end;
function  GetText
  (
    pNo in number
  )
  return varchar is
  begin
    return Text.GetText(svIdtext,pNo);
  end;

function  GetReportName
  return varchar is
  begin
    return 'Detail Customer Statement';
  end;

 -- ************************************
 -- BatchProc support
 -- ************************************

procedure ExecReportOper (pOper number)
  is
    vCurParam number;
    vBeginDate    date := null;
    vEndDate      date := null;
    vCardNo       varchar2(40) := null;
    vCardMBR      number := null;
    vFileName     varchar2(3000) := null;
    vPath         varchar2(3000) := null;
    vUseAchive    boolean := false;
  begin
    if pOper = ReportBP.OP_GetParams then
      ReportBP.RegisterParam(param_BeginDate,  ReportBP.TYPE_DATE,   'Start date', null, false, to_char(Seance.GetOperDate,ReportBP.DateFormat));
      ReportBP.RegisterParam(param_EndDate  ,  ReportBP.TYPE_DATE,   'End date', null, false, to_char(Seance.GetOperDate,ReportBP.DateFormat));
      ReportBP.RegisterParam(param_CardNo,     ReportBP.TYPE_STRING, 'Card Pan', 20, true);
      ReportBP.RegisterParam(param_CardMBR,    ReportBP.TYPE_NUMBER, 'Card MBR', 1,true);
      ReportBP.RegisterParam(param_Path,       ReportBP.TYPE_STRING, 'Unloading path ', 60, true,DataMember.GetChar(Object.GetType('BRANCH'), sBranch, 'REPEG81_PATH'));
      ReportBP.RegisterParam(param_FileName,   ReportBP.TYPE_STRING, 'File name ', 60, true,DataMember.GetChar(Object.GetType('BRANCH'), sBranch, 'REPEG81_FILENAME'));
      ReportBP.RegisterParam(param_UseArchive, ReportBP.TYPE_BOOL,   'Use archive table', null,false,'0');
    elsif pOper = ReportBP.OP_Generate then
      ReportBP.Message('Start generate');
      --запросить значение всех параметров
      vBeginDate  := ReportBP.GetDateParam(param_BeginDate);
      vEndDate    := ReportBP.GetDateParam(param_EndDate);
      vCardNo     := ReportBP.GetCharParam(param_CardNo);
      vCardMBR    := ReportBP.GetNumberParam(param_CardMBR);
      vPath       := ReportBP.GetCharParam(param_Path);
      vFileName   := ReportBP.GetCharParam(param_FileName);
      vUseAchive  := ReportBP.GetBoolParam(param_UseArchive);

      svIdText := Text.Open('REPEG81');
      Generate(vBeginDate,
               vEndDate,
               vCardNo,
               vCardMBR,
               vPath,
               vFileName,
               vUseAchive);
      Text.Close(svIdText);
      ReportBP.Message('Finish generate');
    end if;
  end;
begin
  sBranch     := Seance.GetBranch;
  sCurrency   := Seance.GetCurrency;
end;
/
show errors package body REPEG81;

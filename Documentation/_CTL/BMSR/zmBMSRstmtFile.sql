drop procedure zmbmsrstmtfile;

Create DBA Procedure "cap".zmbmsrstmtfile(pIsVip Smallint default 0)
returning char(798);

--Returning
--	Char(25), 			--01 Card number
--	Decimal(16,3), 			--02 Credit Limit
--	VarChar(50),			--03 Address Line1
--	VarChar(50), 			--04 Address Line2
--	VarChar(50), 			--05 Address Line3
--	VarChar(50), 			--06 Address Line4
--	VarChar(50), 			--07 Address Line5
--	date, 				--08 billing date
--	date, 				--09 print due date
--	decimal(16,3),			--10 fOpeningBalance
--	decimal(16,3), 			--11 fDueAmount
--	decimal(16,3), 			--12 fMinDueAmount
--	Date, 				--13 fPostdate
--	Date,				--14 fi013_trxn_date
--	Char(54), 			--15 trxn details
--	Char(3), 			--16 fTrxnCur
--	Smallint,               	--17 fDecDigits
--	Decimal(16,3), 			--18 fTrxnAmount
--	Decimal(16,3), 			--19 fAmount
--	Char(30), 			--20 fName
--	Integer,			--21 fStmtSerno
--	Char(25), 			--22 fMainCardNumber
--	Integer, 			--23 fPrevSerno
--	Integer, 			--24 fTrxnSerno
--	Char(3), 			--25 fCurrAlpha
--	decimal(16,3), 			--26 fTotalDebits
--	decimal(16,3),			--27 fTotalCredits
--	varchar(50),			--28 global Message 1
--	varchar(50),			--29 global Message 2
--	varchar(50),			--30 global Message 3
--	char(1),			--31 Statement Type (F- financial, M-Memo)
--	varchar(50),			--32 global Message 4
--	varchar(50),			--33 global Message 5
--	decimal(16,3), 			--34 fTotalFees_Int
--	decimal(16,3), 			--35 fTotalCash_Purch
--	decimal(16,3),			--36 fTotalPymts
--	varchar(30),			--37 fBankAccount
--	varchar(30)			--38 fi031_ARN
--	;

------------------------------------------------------------------------------------
Define 	xnumber	        Char(25);
Define 	xCreditLimit	Decimal(16,3);
Define 	xAddLine1	VarChar(50);
Define 	xAddLine2	VarChar(50);
Define 	xAddLine3	VarChar(50);
Define 	xAddLine4	VarChar(50);
Define 	xAddLine5	VarChar(50);
Define 	xbilldate	date;
Define 	xprduedate	date;
Define 	xOpenBal	decimal(16,3);
Define 	xDueAmount	decimal(16,3);
Define 	xMinDueAmt	decimal(16,3);
Define 	xPostdate	Date;
Define 	xtrxndate	Date;
Define 	xtrxndetails	Char(54);
Define 	xTrxnCur	Char(3);
Define 	xDecDigits	Smallint;
Define 	xTrxnAmt	Decimal(16,3);
Define 	xAmount	        Decimal(16,3);
Define 	xName	        Char(30);
Define 	xStmtSerno	Integer;
Define 	xMainCardNum	Char(25);
Define 	xPrevSerno	Integer;
Define 	xTrxnSerno	Integer;
Define 	xCurrAlpha	Char(3);
Define 	xTotalDB	decimal(16,3);
Define 	xTotalCR	decimal(16,3);
Define 	xgMessage1	varchar(50);
Define 	xgMessage2	varchar(50);
Define 	xgMessage3	varchar(50);
Define 	xStmtType	char(1);
Define 	xgMessage4	varchar(50);
Define 	xgMessage5	varchar(50);
Define 	xTotalFees	decimal(16,3);
Define 	xTotCashPur	decimal(16,3);
Define 	xTotalPymts	decimal(16,3);
Define 	xBankAcc	varchar(30);
Define 	xtrxnref	varchar(30);

Define  foldcard   char(16);
Define  fDelimiter char(1);
Define  fline      Char(798);
Define  fcount	   Integer;
Define  faddcount  Integer;

Define fexpirydate date;
Define	FPeplName  char(60);

Define 	fZipCode	varchar(5);
Define 	fBarCode	varchar(16);

Let fDelimiter = '|';
Let foldcard= 00;
Let fline ='';
Return  fline  with resume;

Foreach
 Select tnumber, tCreditLimit, tAddLine1, tAddLine2, tAddLine3, tAddLine4, 
  tAddLine5, tbilldate, tprduedate, tOpenBal, tDueAmount, tMinDueAmt,
  tPostdate, ttrxndate, ttrxndetails, tTrxnCur, tDecDigits, tTrxnAmt,
  tAmount, tName, tStmtSerno, tMainCardNum, tPrevSerno, tTrxnSerno,
  tCurrAlpha, tTotalDB, tTotalCR, tgMessage1, tgMessage2, tgMessage3,
  tStmtType, tgMessage4, tgMessage5, tTotalFees, tTotCashPur, tTotalPymts,
  tBankAcc, ttrxnref
 into xnumber, xCreditLimit, xAddLine1, xAddLine2, xAddLine3, xAddLine4, 
  xAddLine5, xbilldate, xprduedate, xOpenBal, xDueAmount, xMinDueAmt,
  xPostdate, xtrxndate, xtrxndetails, xTrxnCur, xDecDigits, xTrxnAmt,
  xAmount, xName, xStmtSerno, xMainCardNum, xPrevSerno, xTrxnSerno,
  xCurrAlpha, xTotalDB, xTotalCR, xgMessage1, xgMessage2, xgMessage3,
  xStmtType, xgMessage4, xgMessage5, xTotalFees, xTotCashPur, xTotalPymts,
  xBankAcc, xtrxnref
 from stmtfile
  where tnumber IN (select t.tnumber from stmtfile t
    where t.tnumber[1, 6] = '420598'
    GROUP BY t.tnumber
    HAVING COUNT (*) > 4) 
  or tnumber[1, 6] in('410724','498813')
  order by tnumber,tPostdate,tTrxnSerno
 --where tbilldate is not Null
-- order by tnumber,tPostdate,tTrxnSerno

if pIsVip = 1 then   ---select VIP Cards
  If ( select p.vip
   from cardx c ,people p 
   where c.peopleserno = p.serno AND c.number = xnumber )= 0 then
    CONTINUE FOREACH;
  End If;
elIf pIsVip= 2 then  ---dont select VIP Cards
  If ( select p.vip
   from cardx c ,people p 
   where c.peopleserno = p.serno AND c.number = xnumber )= 1 then
    CONTINUE FOREACH;
  End If;
End If

   let FPeplName = '';
   if xnumber != xMainCardNum then -- if transaction fro sub card
     Select trim(p.firstname) || ' ' || trim(p.midname) 
        || ' ' || trim(p.lastname) as PeplName  
      into FPeplName
     from cardx c ,people p 
     where c.peopleserno = p.serno 
     and c.number = xMainCardNum;
   end if;      
 
 
 Select expirydate 
  Into fexpirydate
  From Cardx Where  Number = xnumber;
      
 
 if xnumber[1,16] <> foldcard  then
  --if xtrxndate is not Null then    
  let fline[1,1]     = "H";
  let fline[2,17]	 = xnumber[1,16];
  let fline[18,33]   = lpad(round(xCreditLimit*100), 16,'0');
  let fline[34,83]   = xAddLine1;
  let fline[84,133]  = xAddLine2;
  let fline[134,183] = xAddLine3;
  let fline[184,233] = xAddLine4;
  let fline[234,283] = xAddLine5;
  --05/02/04 if xTotalDB = xTotalPymts
  -- then
  --   let fline[284,291] = "";
  --   let fline[292,299] = "";
  --else
  let fline[284,291] = To_Char(xbilldate, '%Y%m%d');
  let fline[292,299] = To_Char(xprduedate, '%Y%m%d');
  --end if; 
  --05/02/04 if xTotalDB = xTotalPymts
  --  then
  --   let fline[300,315] = "0000000000000000";
  --else
  let fline[300,315] = lpad(round(abs(xMinDueAmt)*100), 16,'0');
  --end if;
  let fline[316,345] = xName;
  let fline[346,361] = lpad(round(abs(xOpenBal)*100), 16,'0');
  if  xopenbal > 0  then
   let fline[362,363] = "CR";
  else 
   let fline[362,363] = "DB";
  end if;
  let fline[364,379] = lpad(round(xTotalDB*100), 16,'0');
  let fline[380,395] = lpad(round(xTotalPymts*100), 16,'0');
  let fline[396,411] = lpad(round(abs(xDueAmount)*100), 16,'0');
  if  xDueAmount > 0  then
   let fline[412,413] = "CR";
  else 
   let fline[412,413] = "DB";
  end if;
  let fline[414,463] = xgMessage1;
  let fline[464,513] = xgMessage2;
  let fline[514,563] = xgMessage3;
  let fline[564,613] = xgMessage4;
  let fline[614,663] = xgMessage5;
  let fline[664,693] = xBankAcc;
  let fline[694,709] = xMainCardNum[1,16];
  let fline[710,710] = xStmtType;
  let fline[711,726] = lpad(round(xTotalCR*100), 16,'0');
  let fline[727,742] = lpad(round(xTotalFees*100), 16,'0');
  let fline[743,758] = lpad(round(xTotCashPur*100), 16,'0');
  let fline[759,759] = xDecDigits;       
  let fline[760,762] = xCurrAlpha;

 select d.zip ,d.position
  into fZipCode, fBarCode 
  from cardx c, cardaddressview v, caddresses d 
  where c.serno = v.cardserno 
--   AND v.mailaddserno = d.serno 
   AND v.stmtaddserno = d.serno 
   and c.number = xnumber;

  select count(*) Into fcount
   from stmtfile
   where tnumber = xnumber
   and ttrxndate is not null;

  select count(*) Into faddcount
   from stmtfile
   where tnumber = xnumber
   and ttrxndate is not null
   and ( ttrxndetails[40,54] <> ""
   or tTrxnAmt <> 0 ) ;

  let fcount= fcount + faddcount;

  let fline[763,772] = To_Char(fexpirydate, '%Y%m%d');
  let fline[773,775] = lpad(fcount, 3,'0');
  let fline[776,782] = fZipCode;
  let fline[783,798] = fBarCode;

  Return  fline  with resume;
   let fline[1,1]    = "D";
   let fline[2,17]	 = xnumber[1,16];
   let fline[18,25]  = To_Char(xPostdate, '%Y%m%d');
   let fline[26,33]  = To_Char(xtrxndate, '%Y%m%d');
   let fline[34,87]  = xtrxndetails;
   let fline[88,90]  = xTrxnCur;
   --let fline[91,91]   = xDecDigits;
   let fline[91,106]  = lpad(round(abs(xTrxnAmt)*100), 16,'0');
   let fline[107,122] = lpad(round(abs(xAmount)*100), 16,'0');
   if  xAmount > 0  then
    let fline[123,124] = "CR";
   else 
    let fline[123,124] = "DB";
   end if;
   let fline[125,154] = xtrxnref;
   let fline[155,160] = substr(xTrxnSerno,1,6);
   let fline[161,176] = xMainCardNum[1,16];
   let fline[177,236] = FPeplName;
   let fline[237,798] = '';
   --let fline[162,164] = xCurrAlpha;

   if fline[2,2] <> 1 then
    Return  fline  with resume;
    let foldcard = xnumber[1,16];         
   end if;      
   --end if;
  else
   -- Details 
   let fline[1,1]     = "D";
   let fline[2,17]	 = xnumber[1,16];
   let fline[18,25]   = To_Char(xPostdate, '%Y%m%d');
   let fline[26,33]   = To_Char(xtrxndate, '%Y%m%d');
   let fline[34,87]   = xtrxndetails;
   let fline[88,90]   = xTrxnCur;
   --let fline[91,91]   = xDecDigits;
   let fline[91,106]  = lpad(round(abs(xTrxnAmt)*100), 16,'0');
   let fline[107,122] = lpad(round(abs(xAmount)*100), 16,'0');
   if  xAmount > 0  then
    let fline[123,124] = "CR";
   else 
    let fline[123,124] = "DB";
   end if;
   let fline[125,154] = xtrxnref;
   let fline[155,160] = substr(xTrxnSerno,1,6);
   --let fline[162,164] = xCurrAlpha;
   let fline[161,176] = xMainCardNum[1,16];
   let fline[177,236] = FPeplName;
   let fline[237,798] = '';

   Return  fline  with resume;
  end if;
End Foreach; 
End Procedure;

grant execute on procedure zmbmsrstmtfile to public as cap;


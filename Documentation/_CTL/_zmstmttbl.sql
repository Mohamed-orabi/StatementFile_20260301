drop procedure _zmstmttbl;
Create DBA Procedure "cap"._zmstmttbl(pProduct Char(20), pBatchSerno Integer );

--Returning
--	Char(25), 			--01 Card number
--	Decimal(16,3), 			--02 Credit Limit
--	VarChar(100),			--03 Address Line1
--	VarChar(100), 			--04 Address Line2
--	VarChar(100), 			--05 Address Line3
--	VarChar(100), 			--06 Address Line4
--	VarChar(100), 			--07 Address Line5
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
--	Char(30), 				--20 fName
--	Integer,				--21 fStmtSerno
--	Char(25), 				--22 fMainCardNumber
--	Integer, 				--23 fPrevSerno
--	Integer, 				--24 fTrxnSerno
--	Char(3), 				--25 fCurrAlpha
--	decimal(16,3), 			--26 fTotalDebits
--	decimal(16,3),			--27 fTotalCredits
--	varchar(100),			--28 global Message 1
--	varchar(100),			--29 global Message 2
--	varchar(100),			--30 global Message 3
--	char(1),				--31 Statement Type (F- financial, M-Memo)
--	varchar(100),			--32 global Message 4
--	varchar(100),			--33 global Message 5
--	decimal(16,3), 			--34 fTotalFees_Int
--	decimal(16,3), 			--35 fTotalCash_Purch
--	decimal(16,3),			--36 fTotalPymts
--	varchar(30),			--37 fBankAccount
--	varchar(30)				--38 fi031_ARN
-->apach
--	,Date 				    --39 fexpirydate 
--	,Char(4)				--40 fi000_msg_type
--	,Char(6)				--41 fi003_proc_code
--	,Date 				    --40 fprevbillingdate 
--<	
--	;

------------------------------------------------------------------------------------
Define famount, fTrxnAmount,fTotalDebits, fTotalCredits Decimal(16,3);
Define fi013_trxn_date,  fpostdate date;
Define fTrxnCur Char(3);
Define fDecdigits Smallint;
Define fLine1 Char(54);
Define fStmtSerno, fCardSerno,fPrimaryCardSerno, fMainCardSerno, fPeopleSerno Integer;
Define fnumber, fMainCardNumber , faccNumber Char(25);
Define fCreditLimit decimal(16,3);
Define faddress1, faddress2, faddress3, faddress4, faddress5 VarChar(100);
Define fStmtAddress1, fStmtAddress2, fStmtAddress3, fStmtAddress4, fStmtAddress5 VarChar(100);
Define fcity char(20);
Define fzip char(10);
Define fName char(100);
Define fBillingDate, fPrintDueDate date;
Define fOpeningBalance decimal(16,3);
Define fDueAmount,fMinDueAmount decimal(16,3);
Define fTitle Char(20);
Define fFirstName, fLastname,   fmidname Char(30);
Define fStmtCount Integer;
Define fPrevSerno, fTrxnSerno Integer;
Define fCloseAmount Decimal(16,3);
Define fCurrAlpha Char(3);
Define fCurrency Char(3);
Define fMerCountryCode Char(3);
define fTotalFees_Int, fTotalCash_Purch, fTotalPymts decimal(16,3);
define fBankAccount varchar(30);
define fCount integer;
define fMultiPageTest char(1);
define fi031_ARN varchar(30);

define rPrimaryCardSerno, rCAccSerno, dCAccSerno, dCardSerno, dTypeSerno_Divert integer;
define fDiversion char(1);
-->apach
Define fexpirydate date;
Define fprevstmtserno Integer;
--Define fprevbillingdate date;
--<

--CH 13/12/2002 - remove curr stmt
--CH 10/09/2002
-- Define fCurrFlag SmallInt ;
-- Define fAccSerno Integer ;
-- Define Global gCurrFlag SmallInt default 0 ;
-- Define Global gAccSerno Integer default 0 ;
--End CH

Define fSortOrder Integer;
Define StmtMessage1, StmtMessage2, StmtMessage3, StmtMessage4, StmtMessage5 varChar(100);

-- gProduct is used to indicate the product to be used for statement selection.
-- when gProduct is "statement sorting" then the selection has already been performed in VB.
Define Global gProduct Char(20)  Default Null;

-- When gStmtSerno is Set, it denotes viewing of a single statement through the Cards program
Define Global gStmtSerno Integer Default Null;

-- gCardSerno is used in combination with gStmtSerno, when set, it denotes viewing of a memo statement from the Cards program for gStmtSerno and gCardSerno
Define Global gCardSerno Integer Default NULL;

-- gRePrint used only to diffentiate between original statements and batch reprints, when 0 then original print, o.w. a re-print
Define Global gRePrint SmallInt Default 0;

-- gStmtCount denotes the maximum number of statements to print for the current run.
Define Global gStmtCount Integer  Default 0;

--gBatchSerno is the batch to link to when printing original statements, and is the batch to select from during re-print.
Define Global gBatchSerno Integer Default Null;

--gStmtPrintedCount is the number of statements actually printed. It is used to communicate back to VB the number of printed statements.
Define Global gStmtPrintedCount Integer Default 0;

--gMemoStatements is a flag indicating what statements to print. It takes the following values:
-- 1 - Print Financial Statements only.
-- 2 - Print Both Financial and Memo Statements.
-- 3 - Print Memo Statements only.
-- 4 - Print a single memo statement (denoted by gStmtSerno and gCardSerno)
Define Global gMemoStatements SmallInt Default 0;

-- When gUseCardInfo is set, then the financial statement uses the main card details (number, person, address), when not set, the account details are used.
Define Global gUseCardInfo Integer Default 1;

Define fMode SmallInt;
Define fcAccSerno, fSerno Integer;
Define fSummaryIndicator Char(1);
Define fPrevCardSerno,fRiskDomainSerno Integer;
Define fi003_proc_code Char(6);
Define fi000_msg_type Char(4);
Define fOrig_msg_type Char(5);
Define fi002_number Char(25);
Define fi004_amt_trxn, fi006_amt_bill Decimal(16,3);
Define fi048_text_data VarChar(250,0);
Define fi049_cur_trxn, fi051_cur_bill Char(3);
Define fi043a_merch_name char(25);
Define fi043b_merch_city char(20);
Define fi043c_merch_cnt Char(3);
Define fAlphaCode Char(3);
Define fMakeAuthGroups SmallInt;
Define fAuthCurrency Char(3);
Define fAuthLimit,fCardCreditLimit Decimal(16,3);
Define fCardProduct Char(20);
-- Trace for Procedure _zmstmttbl
Define Sql_e Smallint;
Define Isam_e smallint;
Define Txt_e Char(80);

On Exception                                 -- In case of error catch it here and log it.
	Set Sql_e, Isam_e, Txt_e
	Begin
	Define Sql_trace Smallint;
	Define Isam_trace smallint;
	Define Txt_trace Char(80);
	On Exception                                   -- In case of an error during trace
		Set  Sql_trace, Isam_trace, Txt_trace
		Trace  '	' || sql_trace || '	' || isam_trace || '	' || txt_trace;
	End Exception With Resume;
	Set Debug File To cap_getparam('center', 'error log file',  '\temp\caperr.log') With Append;
	Trace  Current || '	' || user || '	' || sql_e || '	' || isam_e || '	' || txt_e;
	Trace '	Procedure: _zmstmttbl';
	End
	Raise Exception Sql_e, Isam_e, Txt_e;
End Exception;
-- End Trace for Procedure _zmstmttbl

On Exception				 -- In case of error catch it here and log it.
	Set Sql_e, Isam_e, Txt_e
	Begin
		On Exception
		End Exception with resume;
		Drop Table tStmtSerno;
	End
	Raise Exception Sql_e, Isam_e, Txt_e;
End Exception;

--Set Debug File to 'g:\maher\zmstmttbl.txt';
--Trace On;

let fMultiPageTest = 'N';
Let fPrevSerno = 0 ;
Let fStmtCount = 0 ;
--CH 13/12/2002 - remove curr stmt
--CH 10/09/2002
-- Let fCurrFlag = gCurrFlag ;
-- Let fAccSerno = gAccSerno ;
-- 
-- If fCurrFlag = 1 Then
-- 	Begin Work ;
-- 	execute procedure acct_close_temp(fAccSerno, 0, null, null, 0) ;
-- End If ;
--End CH
--let gstmtserno = 4090 ; -- testing
-- gRePrint used only to diffentiate between original statements and batch reprints
-->apach
 let fi000_msg_type = '';
 let fi003_proc_code = '';
--<	
----------------( new changes )----------------------------------------
--Create  Table stmtfile(
-- tnumber	Char(25),                       
-- tCreditLimit	Decimal(16,3), 
-- tAddLine1	VarChar(50), 
-- tAddLine2	VarChar(50), 
-- tAddLine3	VarChar(50), 
-- tAddLine4	VarChar(50), 
-- tAddLine5	VarChar(50), 
-- tbilldate	date, 
-- tprduedate	date, 
-- tOpenBal	decimal(16,3), 
-- tDueAmount	decimal(16,3), 
-- tMinDueAmt	decimal(16,3), 
-- tPostdate	Date, 
-- ttrxndate	Date, 
-- ttrxndetails	Char(54), 
-- tTrxnCur	Char(3), 
-- tDecDigits	Smallint,               
-- tTrxnAmt	Decimal(16,3), 
-- tAmount	Decimal(16,3), 
-- tName	        Char(30), 
-- tStmtSerno	Integer, 
-- tMainCardNum	Char(25), 
-- tPrevSerno	Integer, 
-- tTrxnSerno	Integer, 
-- tCurrAlpha	Char(3), 
-- tTotalDB	decimal(16,3), 
-- tTotalCR	decimal(16,3), 
-- tgMessage1	varchar(50), 
-- tgMessage2	varchar(50), 
-- tgMessage3	varchar(50), 
-- tStmtType 	char(1), 
-- tgMessage4	varchar(50), 
-- tgMessage5	varchar(50), 
-- tTotalFees     decimal(16,3), 
-- tTotCashPur	decimal(16,3), 
-- tTotalPymts	decimal(16,3), 
-- tBankAcc 	varchar(30), 
-- ttrxnref	varchar(30) 
--) ;

-----------------------------------------------------------------------
--Create temp Table stmtfile(rec_line char(936)) with No log;
delete from stmtfile;
insert into stmtfile ( tnumber,tpostdate,tbilldate,ttrxndate ) values (1,1,1,1);

---------------------------Printing Original Statements----------------- 
If pBatchSerno = 0 Then  --gRePrint = 0 AND gStmtSerno is Null
 -- Printing Original Statements
 Let fMode = 0;
 -- Use the sorted temp tables if already sorted
-- If lower(pProduct) = "statement sorting" Then
--  Select SortOrder, Serno
--  From tSortedStmt
--  Into Temp tStmtSerno With No Log;
-- Else
  -- Normal select into temp table using a dummy sortorder
  Select S.Serno as SortOrder, S.Serno
   From cStatements s, cAccounts a
   Where a.serno = s.cAccSerno 
    AND a.Product = pProduct 
    AND s.BatchSerno is Null 
    and s.reason <> 'O'
   Into Temp tStmtSerno With No Log;
-- End If;
 Create Index tStmtSernoi01 on tStmtSerno(SortOrder,Serno);
 Update Statistics High For Table tStmtSerno(SortOrder,Serno);
---------------------------Viewing a statement----------------------------
--Elif gStmtSerno is Not  Null  Then
 -- Viewing a statement
-- Let fMode = 1;
-- Let fStmtSerno = gStmtSerno;
-- If gCardSerno is NULL Then
--  Let gMemoStatements = 2;
-- Else
--  Let gMemoStatements = 4;
-- End if;
-- Create Temp Table tStmtSerno(SortOrder Integer, Serno Integer) 
--  with No log;
-- Insert into tStmtSerno (SortOrder, Serno) Values (fStmtSerno, fStmtSerno);
---------------------------Re-Print----------------------------
Else
 -- Re-Print
 Let fMode = 2;
 -- Use the sorted temp tables if already sorted
-- If lower(pProduct) = "statement sorting" Then
--  Select SortOrder, Serno
--   From tSortedStmt
--   Into Temp tStmtSerno With No Log;
-- Else
  -- Normal select into temp table using a dummy sortorder
  Select S.Serno as SortOrder, S.Serno
   From cStatements s, cAccounts a
   Where a.serno = s.cAccSerno 
    AND a.Product = pProduct 
    AND s.BatchSerno = pBatchSerno
  Into Temp tStmtSerno With No Log;
-- End If;
 Create Index tStmtSernoi01 on tStmtSerno(SortOrder,Serno);
 Update Statistics High For Table tStmtSerno(SortOrder,Serno);
 Let gMemoStatements = 2;
End If;
------------------------------------------------------------------
Foreach
Select SortOrder, Serno
 Into fSortOrder, fStmtSerno
 From tStmtSerno
 Order by 1

let fStmtSerno = fStmtSerno;
-- 10/09/2003 VMS N124464 Initialization of Summary Fields
let fTotalFees_Int = 0;
let fTotalCash_Purch = 0;
let fTotalPymts = 0;
-- Statement Information
Select cAccSerno,BillingDate, ClosingBalance, Openingbalance, MinDueAmount
 , PrintDueDate, TotalDebits, TotalCredits
 Into fcAccSerno, fBillingDate, fDueAmount, fOpeningBalance, fMinDueAmount
 , fPrintDueDate, fTotalDebits, fTotalCredits
 From cStatements
 Where Serno = fStmtSerno;

-->apach
-- Select billingdate
--  Into fprevbillingdate
--  From cStatements
--  Where Serno = fprevstmtserno;
--<

-- Account Information
Select Number,CreditLimit, PrimaryCardSerno, Currency
 Into fNumber, fCreditLimit, fPrimaryCardSerno, fCurrency
 From cAccounts
 Where Serno = fcAccSerno;

let faccNumber = fNumber;  -- CRAG
-- 12/03/2004 VMS <Diversion Support> for MSCC
--rb get the first diverted card
let dCardSerno=null;
foreach
 select cardxserno
  into dCardSerno
  from caccountrouting
  where caccserno = fcAccSerno
 exit foreach;
end foreach;

if dCardSerno is not null then -- Diversion is Defined
 let rCAccSerno = (select caccserno from cardx where serno = dCardSerno);
 let rPrimaryCardSerno = (select primarycardserno from cAccounts 
  where Serno = rCAccSerno);
		
 -- Replace PrimaryCardSerno
 if rPrimaryCardSerno is not null then
  let fPrimaryCardSerno = rPrimaryCardSerno;
 end if;
 let fDiversion = 'Y';
else
 let fDiversion = 'N';
end if;

 -- Account Statement Details
 select bankaccount
  into fBankAccount
  from caccountstmt
  where serno = fcAccSerno;
	
 -- Currency Alpha
 Select AlphaCode Into fCurrAlpha From Currencies Where numcode = fCurrency;

 -- Get Product Level Statement Messages
 Let fCardProduct = (select product from cardx where Serno = fPrimaryCardSerno);
-- Let StmtMessage1 = cap_getparam('statement messages', 'cr message line 1', fCardProduct);
-- Let StmtMessage2 = cap_getparam('statement messages', 'cr message line 2', fCardProduct);
-- Let StmtMessage3 = cap_getparam('statement messages', 'cr message line 3', fCardProduct);
-- Let StmtMessage4 = cap_getparam('statement messages', 'cr message line 4', fCardProduct);
-- Let StmtMessage5 = cap_getparam('statement messages', 'cr message line 5', fCardProduct);

Let StmtMessage1 = cap_getparam('statement messages', 'cr message line 1', '(Default)', fCardProduct);
Let StmtMessage2 = cap_getparam('statement messages', 'cr message line 2', '(Default)', fCardProduct);
Let StmtMessage3 = cap_getparam('statement messages', 'cr message line 3', '(Default)', fCardProduct);
Let StmtMessage4 = cap_getparam('statement messages', 'cr message line 4', '(Default)', fCardProduct);
Let StmtMessage5 = cap_getparam('statement messages', 'cr message line 5', '(Default)', fCardProduct);

	
 -- Print Financial Statement if gMemoStatements in (1,2)
 If gMemoStatements in (1,2) Then
 -- People/Number information from account OR card based on gUseCardInfo
  If gUseCardInfo = 0 OR fPrimaryCardSerno is NULL Then
   Select PeopleSerno Into fPeopleSerno From cAccountStmt 
    Where Serno = fcAccSerno;
  Else
-->apach   --commented
 Select Number, PeopleSerno,expirydate 
  Into fNumber, fPeopleSerno , fexpirydate
  From Cardx Where Serno = fPrimaryCardSerno;
--<
  End If;
		
  -- 12/03/2004 VMS <Diversion Support> for MSCC
  if fDiversion = 'Y' then
   let fNumber = trim(fNumber) || ' *';
   -- SM N122691 07/05/2004 for diversion cards creditlimit is zero
   let fCreditLimit=0;
  end if;

  -- People Info
  Select Title, Firstname, Lastname, Midname 
   Into fTitle, fFirstName, fLastName, fMidName 
   From People Where Serno = fPeopleSerno;
 
   --  Special For CRAG to ADD fMidName to Customer Name in Statement
   If fCardProduct in (select name from products 
    where appliesto = 'C' And institution = 'CRAG' ) then
    Let fName = nvl(Trim(fTitle),'')|| ' ' || Trim(fFirstName) || ' ' || Trim(fMidName) || ' ' || Trim(fLastName);
    If Length(fName) > 30 Then
     Let fName = Trim(fTitle) || ' ' || Trim(fFirstName) || ' ' || Trim(fMidName) || ' ' || Upper(fLastName[1,1]) || '.';
     If Length(fName) > 30 Then
      Let fName = Trim(fFirstName) || ' ' || Trim(fMidName) || ' ' || Upper(fLastName[1,1]) || '.';
     End If;
    End If;
   ELSE
    Let fName = nvl(Trim(fTitle),'')|| ' ' || Trim(fFirstName) || ' ' || Trim(fLastName);
    If Length(fName) > 30 Then
     Let fName = Trim(fTitle) || ' ' || Upper(fFirstName[1,1]) || '. ' || Trim(fLastName);
     If Length(fName) > 30 Then
      Let fName = Upper(fFirstName[1,1]) || '. ' || Trim(fLastName);
     End If;
    End If;
   End If;   
   		
   --- Address
   Select A.address1, A.address2, A.address3, A.address4, A.address5
    , A.city, A.zip
   Into faddress1, faddress2, faddress3, faddress4, faddress5
    , fcity, fzip
   From PeopleAddressView PV , cAddresses A
   Where PV.PeopleSerno = fPeopleSerno 
   AND PV.StmtAddSerno = A.Serno;
   --	PV.CardAddSerno = A.Serno;

If fCardProduct in (select name from products 
   where appliesto = 'C' And institution in ('BDCA')) then
   let fStmtAddress1 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, "", Trim(fZip) || ' ' || fCity, Null, Null, 1);
   Let fStmtAddress2 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, "", Trim(fZip) || ' ' || fCity, Null, Null, 2);
   Let fStmtAddress3 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, "", Trim(fZip) || ' ' || fCity, Null, Null, 3);
   Let fStmtAddress4 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, "", Trim(fZip) || ' ' || fCity, Null, Null, 4);
   Let fStmtAddress5 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, "", Trim(fZip) || ' ' || fCity, Null, Null, 5);
else
   let fStmtAddress1 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, fAddress5, Trim(fZip) || ' ' || fCity, Null, Null, 1);
   Let fStmtAddress2 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, fAddress5, Trim(fZip) || ' ' || fCity, Null, Null, 2);
   Let fStmtAddress3 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, fAddress5, Trim(fZip) || ' ' || fCity, Null, Null, 3);
   Let fStmtAddress4 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, fAddress5, Trim(fZip) || ' ' || fCity, Null, Null, 4);
   Let fStmtAddress5 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, fAddress5, Trim(fZip) || ' ' || fCity, Null, Null, 5);
end if;

  If (select count(*) From StatementLinks 
   Where StatementSerno = fStmtSerno AND LinkType = 'B') <> 0 Then
  Foreach
   Select TransactionSerno, SummaryIndicator
    Into fTrxnSerno, fSummaryIndicator
    From StatementLinks
    Where StatementSerno = fStmtSerno 
     and LinkType = 'B'
			
   Select i003_proc_code, i000_msg_type, Orig_Msg_type,i002_number
    , i004_amt_trxn, i006_amt_bill, i013_trxn_date, i048_text_data
    , i049_cur_trxn, i051_cur_bill, amount, postdate, caccserno
    , cardserno, typeserno_divert
   Into fi003_proc_code, fi000_msg_type, fOrig_msg_type ,fi002_number
    , fi004_amt_trxn, fi006_amt_bill, fi013_trxn_date, fi048_text_data
    , fi049_cur_trxn, fi051_cur_bill, famount, fpostdate, dCAccSerno
    , dCardSerno, dTypeSerno_Divert
   From cTransactions
   Where Serno = fTrxnSerno;
			
   -- 12/03/2004 VMS <Diversion Support> for MSCC
   if dTypeSerno_Divert is not null then
    select caccserno
     into rCAccSerno
     from caccountrouting
     where cardxserno = dCardSerno 
      and rtrxntype = dTypeSerno_Divert;
				
	 {
     if rCAccSerno is null then
      let rCAccSerno = -1;
      let fDiversion = 'N';
     else
      let rCAccSerno = rCAccSerno;
      let fDiversion = 'Y';
     end if;
     }
				
     if rCAccSerno <> fcAccSerno then
      --	continue foreach;
     end if;
    end if;
			
    Select i043a_merch_name, i043b_merch_city, i043c_merch_cnt
     Into fi043a_merch_name, fi043b_merch_city, fi043c_merch_cnt
     From cIsotrxns
     Where Serno = fTrxnSerno;
		        
    If NVL(fi043c_merch_cnt, '') = '' Then
     Let fMerCountryCode = '';
    else
     If Length(Trim(fi043c_merch_cnt)) = 2 Then
      Select alphacode3char
       Into fMerCountryCode
       From countries
       Where alphacode=fi043c_merch_cnt;
      Elif Length(Trim(fi043c_merch_cnt)) = 3 Then
       Select alphacode3char
        Into fMerCountryCode
        From countries
        Where alphacode3char=fi043c_merch_cnt;
      End If;
      If NVL(fMerCountryCode, '') = '' Then
       Let fMerCountryCode = fi043c_merch_cnt;
      end if;
     end if;

     Select alphacode Into falphacode From Currencies 
      Where Numcode = fi049_cur_trxn;

     Select ToTrxnSerno Into fPrevSerno From cTrxnLinks 
      Where FromtrxnSerno = fTrxnSerno and LinkType= 'F';
			
     Let fDecDigits = 0;
     If fi043a_merch_name is Null OR fOrig_msg_type in ( 'REDEM', 'TEXT') Then
      Let fLine1 = fi048_text_data;
      Let fTrxnCur = Null;
      Let fTrxnAmount = 0;
				
      If fi049_cur_trxn <> fi051_cur_bill Then
       Let fTrxnCur = falphacode;
       Let fTrxnAmount = fi004_amt_trxn;
      End If;
     Else
      Let fLine1 = trim(fi043a_merch_name);
     If NVL(fi043b_merch_city, '') <> '' Then
      Let fLine1 = Trim(fLine1) || ", " || Trim(fi043b_merch_city);
     End If;
				
     If NVL(fMerCountryCode, '') <> '' Then
      Let fLine1 = Trim(fLine1) || ", " || Trim(fMerCountryCode);
     End If;
				
     If fi049_cur_trxn = fi051_cur_bill Then
      Let fTrxnCur = Null;
      Let fTrxnAmount = 0;
     Else
      Select AlphaCode,DecDigits Into fAlphaCode, fDecDigits 
       From Currencies Where Numcode = fi049_cur_trxn;
     
      Let fTrxnCur = falphacode;
      Let fTrxnAmount = fi004_amt_trxn;
     End if;
    End If;
			
    -- Total Cash + Purchases
    if fi000_msg_type = '0220' And fi003_proc_code in ('000000', '010000') then
     let fTotalCash_Purch = fTotalCash_Purch + fAmount;
    end if;
			
    -- Total Interest + Fees etc
    let fTotalFees_Int = abs(fTotalDebits) - abs(fTotalCash_Purch);
			
    -- Total Payments
    let fTotalPymts = fTotalCredits;
			
   -- Obtain ARN Number
   let fi031_ARN = (select I031_arn from cisotrxns where serno = fTrxnSerno);
			
   -- CRAG  --fCardProduct
    If fi002_number <> faccNumber 
     and pProduct in (select name from products 
     where appliesto = 'C' And institution in ('CRAG','BMSR')) then

    insert into stmtfile values	( fnumber, fCreditLimit, fStmtAddress1
     , fStmtAddress2, fStmtAddress3, fStmtAddress4, fStmtAddress5
     , fBillingDate, fPrintDueDate, fOpeningBalance, fDueAmount
     , fMinDueAmount, fPostdate, fi013_trxn_date, fLine1, fTrxnCur
     , fDecDigits, fTrxnAmount, fAmount, fName[1,30], fStmtSerno
     , fi002_number, fPrevSerno, fTrxnSerno, fCurrAlpha, fTotalDebits
     , fTotalCredits, StmtMessage1, StmtMessage2, StmtMessage3, 'F'
     , StmtMessage4, StmtMessage5, abs(fTotalFees_Int)
     , abs(fTotalCash_Purch) , abs(fTotalPymts), fBankAccount, fi031_ARN );
    else 
     insert into stmtfile values ( fnumber, fCreditLimit, fStmtAddress1
      , fStmtAddress2, fStmtAddress3, fStmtAddress4, fStmtAddress5
      , fBillingDate, fPrintDueDate, fOpeningBalance, fDueAmount
      , fMinDueAmount, fPostdate, fi013_trxn_date, fLine1, fTrxnCur
      , fDecDigits, fTrxnAmount, fAmount, fName[1,30], fStmtSerno
      , fNumber, fPrevSerno, fTrxnSerno, fCurrAlpha, fTotalDebits
      , fTotalCredits, StmtMessage1, StmtMessage2, StmtMessage3, 'F'
      , StmtMessage4, StmtMessage5, abs(fTotalFees_Int)
      , abs(fTotalCash_Purch), abs(fTotalPymts), fBankAccount, fi031_ARN );
     end if;
                        
     if fMultiPageTest = 'Y' then
      let fMultiPageTest = 'N';
      for fCount = 1 to 100
       insert into stmtfile values ( fnumber, fCreditLimit, fStmtAddress1
        , fStmtAddress2, fStmtAddress3, fStmtAddress4, fStmtAddress5
        , fBillingDate, fPrintDueDate, fOpeningBalance, fDueAmount
        , fMinDueAmount, fPostdate, fi013_trxn_date, fLine1, fTrxnCur
        , fDecDigits, fTrxnAmount, fAmount, fName[1,30], fStmtSerno
        , fNumber, fPrevSerno, fTrxnSerno, fCurrAlpha, fTotalDebits
        , fTotalCredits, StmtMessage1, StmtMessage2, StmtMessage3, 'F'
        , StmtMessage4, StmtMessage5, abs(fTotalFees_Int)
        ,abs(fTotalCash_Purch) , abs(fTotalPymts), fBankAccount
        , fi031_ARN );
       end for;
      end if;
     End Foreach;
    else
     insert into stmtfile values ( fnumber, fCreditLimit, fStmtAddress1
      , fStmtAddress2, fStmtAddress3, fStmtAddress4, fStmtAddress5
      , fBillingDate, fPrintDueDate, fOpeningBalance, fDueAmount
      , fMinDueAmount, null, null, '', '', '', '', '', fName[1,30]
      , fStmtSerno, fNumber, fPrevSerno, '', fCurrAlpha, fTotalDebits
      , fTotalCredits, StmtMessage1, StmtMessage2, StmtMessage3, 'F'
      , StmtMessage4, StmtMessage5, abs(fTotalFees_Int)
      , abs(fTotalCash_Purch), abs(fTotalPymts), fBankAccount, null );
     end if;
		
     -- Number of statements printed
     Let fStmtCount = fStmtCount + 1;
    End If;
	
    -- Print Memo Statements if gMemoStatements in (2,3,4)

    If fCardProduct <> "CB-V-CR-BUS" then
       If fCardProduct in (select name from products 
          where appliesto = 'C' And institution in ('CRAG','BMSR','CBB')) then
          let gMemoStatements = 1;
          end if;
    end if;  


    if gMemoStatements in (2,3,4) then
     --rb If gMemoStatements in (2,3,4) and fDiversion = 'N' Then 
     -- 12/03/2004 VMS <Diversion Support> for MSCC
     Let fPrevCardSerno = -1;
     Foreach
     Select StatementSerno, TransactionSerno, CardSerno, SummaryIndicator
      Into fSerno, fTrxnSerno, fCardSerno, fSummaryIndicator
      From StatementLinks
      Where StatementSerno = fStmtSerno 
       AND CardSerno is not NULL 
       AND LinkType = 'B'
      Order by StatementSerno, CardSerno
			
     -- If gMemoStatements = 4 then generate a memo statement for gCardSerno only.
     If gMemoStatements = 4 AND gCardSerno <> fCardSerno Then
      Continue ForEach;
      -- Do not generate memo statement for PrimaryCard if gUseCardInfo is set
     ElIf gUseCardInfo <> 0 AND NVL(fPrimaryCardSerno, 0) = fCardSerno Then
      Continue ForEach;
     End If;
			
     If fPrevCardSerno <> fCardSerno Then
      -- SM N122691 07/05/2004 add * for supplementary card number
      Select Number, PeopleSerno, RiskDomainSerno, Product Into fNumber
       , fPeopleSerno, fRiskDomainSerno, fCardProduct 
      From Cardx Where Serno = fCardSerno;
      -- Find Credit Limit
	  -- SM N122691 07/05/2004 add * for supplementary card number
      -- 12/03/2004 VMS <Diversion Support> for MSCC
      --rb get the first diverted card
      let fcAccSerno= (select caccserno from cstatements 
        where serno = fStmtSerno);
      let dCardSerno=null;
      foreach
       select cardxserno
        into dCardSerno
        from caccountrouting
       where caccserno = fcAccSerno
       exit foreach;
      end foreach;
		
      if dCardSerno is not null then -- Diversion is Defined
      --let rCAccSerno = (select caccserno from cardx where serno = dCardSerno);
      --let rPrimaryCardSerno = (select primarycardserno from cAccounts where Serno = rCAccSerno);
			
      -- Replace PrimaryCardSerno
      --if rPrimaryCardSerno is not null then
      -- let fPrimaryCardSerno = rPrimaryCardSerno;
      --end if;
      let fDiversion = 'Y';
     else
      let fDiversion = 'N';
     end if;

     If fRiskDomainSerno is not NULL Then
      Select makeauthgroups, authcurrency, authlimit 
       Into fMakeAuthGroups, fAuthCurrency, fAuthLimit 
       From RiskDomains 
       Where Serno = fRiskDomainSerno;
     
      If fMakeAuthGroups <> 0 AND fAuthLimit >= 0 Then
       Let fCardCreditLimit = fAuthLimit * 
          (Case When fAuthCurrency <> fCurrency 
           Then cap_exrate(fAuthCurrency, fCurrency, fCardProduct) 
           Else 1.0 End);
       Else
        Let fCardCreditLimit = fCreditLimit;
       End If;
      Else
      Let fCardCreditLimit = fCreditLimit;
     End If;

     if fDiversion = 'Y' then
      let fNumber = trim(fNumber) || ' *';
      -- SM N122691 07/05/2004 for diversion cards creditlimit is zero
      let fCardCreditLimit=0;
     end if;

     -- People Info
     Select Title, Firstname, Lastname, Midname Into fTitle, fFirstName
       , fLastName, fMidName 
      From People 
      Where Serno = fPeopleSerno;
      --  Special For CRAG to ADD fMidName to Customer Name in Statement
     
     If fCardProduct in (select name from products 
       where appliesto = 'C' And institution = 'CRAG' ) then
      Let fName = nvl(Trim(fTitle),'')|| ' ' || Trim(fFirstName) || ' ' || Trim(fMidName) || ' ' || Trim(fLastName);
      If Length(fName) > 30 Then
       Let fName = Trim(fTitle) || ' ' || Trim(fFirstName) || ' ' || Trim(fMidName) || ' ' || Upper(fLastName[1,1]) || '.';
       If Length(fName) > 30 Then
        Let fName = Trim(fFirstName) || ' ' || Trim(fMidName) || ' ' || Upper(fLastName[1,1]) || '.';
       End If;
      End If;
     ELSE
      Let fName = nvl(Trim(fTitle),'')|| ' ' || Trim(fFirstName) || ' ' || Trim(fLastName);
      If Length(fName) > 30 Then
       Let fName = Trim(fTitle) || ' ' || Upper(fFirstName[1,1]) || '. ' || Trim(fLastName);
       If Length(fName) > 30 Then
        Let fName = Upper(fFirstName[1,1]) || '. ' || Trim(fLastName);
       End If;
      End If;
     End If;   
				
     --- Address
     Select A.address1, A.address2, A.address3, A.address4, A.address5
      , A.city, A.zip
     Into faddress1, faddress2, faddress3, faddress4, faddress5
      , fcity, fzip
     From PeopleAddressView PV , cAddresses A
     Where PV.PeopleSerno = fPeopleSerno 
      AND PV.StmtAddSerno = A.Serno;
      -- PV.CardAddSerno = A.Serno;
     
--     Let fStmtAddress1 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, fAddress5, Trim(fZip) || ' ' || fCity, Null, Null, 1);
--     Let fStmtAddress2 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, fAddress5, Trim(fZip) || ' ' || fCity, Null, Null, 2);
--     Let fStmtAddress3 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, fAddress5, Trim(fZip) || ' ' || fCity, Null, Null, 3);
--     Let fStmtAddress4 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, fAddress5, Trim(fZip) || ' ' || fCity, Null, Null, 4);
--     Let fStmtAddress5 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, fAddress5, Trim(fZip) || ' ' || fCity, Null, Null, 5);


If fCardProduct in (select name from products 
   where appliesto = 'C' And institution in ('BDCA')) then
   let fStmtAddress1 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, "", Trim(fZip) || ' ' || fCity, Null, Null, 1);
   Let fStmtAddress2 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, "", Trim(fZip) || ' ' || fCity, Null, Null, 2);
   Let fStmtAddress3 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, "", Trim(fZip) || ' ' || fCity, Null, Null, 3);
   Let fStmtAddress4 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, "", Trim(fZip) || ' ' || fCity, Null, Null, 4);
   Let fStmtAddress5 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, "", Trim(fZip) || ' ' || fCity, Null, Null, 5);
else
   let fStmtAddress1 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, fAddress5, Trim(fZip) || ' ' || fCity, Null, Null, 1);
   Let fStmtAddress2 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, fAddress5, Trim(fZip) || ' ' || fCity, Null, Null, 2);
   Let fStmtAddress3 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, fAddress5, Trim(fZip) || ' ' || fCity, Null, Null, 3);
   Let fStmtAddress4 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, fAddress5, Trim(fZip) || ' ' || fCity, Null, Null, 4);
   Let fStmtAddress5 = cap_kth_string(fAddress1, fAddress2, fAddress3, fAddress4, fAddress5, Trim(fZip) || ' ' || fCity, Null, Null, 5);
end if;




     Let fPrevCardSerno = fCardSerno;
    End if;
			
    Select i003_proc_code, Orig_Msg_type,i002_number, i004_amt_trxn
     , i006_amt_bill, i013_trxn_date,i048_text_data, i049_cur_trxn
     , i051_cur_bill, amount, postdate
    Into fi003_proc_code, fOrig_msg_type ,fi002_number, fi004_amt_trxn
     , fi006_amt_bill, fi013_trxn_date, fi048_text_data, fi049_cur_trxn
     , fi051_cur_bill, famount, fpostdate
    From cTransactions
    Where Serno = fTrxnSerno;
			
    Select i043a_merch_name, i043b_merch_city, i043c_merch_cnt
     Into fi043a_merch_name, fi043b_merch_city, fi043c_merch_cnt
     From cIsotrxns
     Where Serno = fTrxnSerno;
			        
    If NVL(fi043c_merch_cnt, '') = '' Then
     Let fMerCountryCode = "";
    else
     If Length(Trim(fi043c_merch_cnt)) = 2 Then
       Select alphacode3char
        Into fMerCountryCode
        From countries
        Where alphacode=fi043c_merch_cnt;
	
     Elif Length(Trim(fi043c_merch_cnt)) = 3 Then
      Select alphacode3char
       Into fMerCountryCode
       From countries
       Where alphacode3char=fi043c_merch_cnt;
     End If;
     
     If NVL(fMerCountryCode, '') = '' Then
      Let fMerCountryCode = fi043c_merch_cnt;
     end if;
    end if;
			
    Select alphacode Into falphacode From Currencies Where Numcode = fi049_cur_trxn;
    
    Select ToTrxnSerno Into fPrevSerno From cTrxnLinks Where FromtrxnSerno = fTrxnSerno and LinkType= 'F';
			
    Let fDecDigits = 0;
    If fi043a_merch_name is Null 
      OR fOrig_msg_type in ( 'REDEM', 'TEXT') Then
     Let fLine1 = fi048_text_data;
     Let fTrxnCur = Null;
     Let fTrxnAmount = 0;
     If fi049_cur_trxn <> fi051_cur_bill Then
      Let fTrxnCur = falphacode;
      Let fTrxnAmount = fi004_amt_trxn;
     End If;
    Else
     Let fLine1 = trim(fi043a_merch_name);
     If NVL(fi043b_merch_city, '') <> '' Then
      Let fLine1 = Trim(fLine1) || ", " || Trim(fi043b_merch_city);
     End If;
				
     If NVL(fMerCountryCode, '') <> '' Then
      Let fLine1 = Trim(fLine1) || ", " || Trim(fMerCountryCode);
     End If;
     If fi049_cur_trxn = fi051_cur_bill Then
      Let fTrxnCur = Null;
      Let fTrxnAmount = 0;
     Else
      Select AlphaCode,DecDigits 
       Into fAlphaCode, fDecDigits 
       From Currencies 
       Where Numcode = fi049_cur_trxn;
      
      Let fTrxnCur = falphacode;
      Let fTrxnAmount = fi004_amt_trxn;
     End if;
    End If;
    -- Obtain ARN Number
    let fi031_ARN = (select I031_arn from cisotrxns 
      where serno = fTrxnSerno);
			
     insert into stmtfile values ( fnumber, fCardCreditLimit, fStmtAddress1
      , fStmtAddress2, fStmtAddress3, fStmtAddress4, fStmtAddress5
      , fBillingDate, fPrintDueDate, 0.0, 0.0, 0.0, fPostdate
      , fi013_trxn_date, fLine1, fTrxnCur, fDecDigits, fTrxnAmount
      , fAmount, fName[1,30], fStmtSerno, fNumber, fPrevSerno
      , fTrxnSerno, fCurrAlpha, 0.0, 0.0, StmtMessage1, StmtMessage2
      , StmtMessage3,'M', StmtMessage4, StmtMessage5, abs(fTotalFees_Int)
      , abs(fTotalCash_Purch), abs(fTotalPymts), fBankAccount, fi031_ARN );
   End Foreach;
		
   -- Number of statements printed
   Let fStmtCount = fStmtCount + 1;
  End If;

If fMode = 0 Then
 -- Update PrintCount, LastPrintDate and BatchSerno
 -- AO 10/01/2002 Do not update if printing memo statements only
 If gMemoStatements <> 3 Then
    Update cStatements Set	PrintCount = PrintCount + 1
     ,LastPrintDate = Current,BatchSerno = pBatchSerno 
     Where Serno = fStmtSerno;
 End if
 -- If fStmtCount >= gStmtCount Then 
 -- Check max number of statements to print
 -- Exit Foreach;
 -- End If;
 ElIf fMode = 2 Then
 -- Update PrintCount and LastPrintDate
 -- AO 10/01/2002 Do not update if printing memo statements only
 If gMemoStatements <> 3 Then
    Update cStatements Set	PrintCount = PrintCount + 1
     , LastPrintDate = Current Where Serno = fStmtSerno;
 End if
--   If fStmtCount >= gStmtCount Then 
    -- Check max number of statements to print
--    Exit Foreach;
--   End If;
End If;



End Foreach;

Drop Table tStmtSerno;

End Procedure; 

grant execute on procedure _zmstmttbl to public as cap;
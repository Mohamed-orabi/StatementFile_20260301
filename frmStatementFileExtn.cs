using System;
using System.Diagnostics;
using System.Windows.Forms;
using CrystalDecisions.Shared;

partial class frmStatementFile
{

    //private void b4FrmLoad()
    //    {
    //    AssemblyInfo ainfo = new AssemblyInfo();
    //    lblVer.Text = "Ver " + ainfo.Version;
    //    //104

    //    chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire Classic >> Text 5/m", 49));//49
    //    chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire Gold >> Text 5/m", 58));//58
    //    chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Standard >> Text 5/m", 138));//138
    //    chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Gold >> Text 5/m", 139));//139
    //    chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Installment >> Text 5/m", 173));//173

    //    chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default Credit Text 7/m", 212));//212

    //    chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default Credit Text 12/m", 213));//213

    //    chkLstProducts.Items.Add(new ListItem("4) BMSR Bank Misr Moga >> Text 16/m", 15));//15
    //    chkLstProducts.Items.Add(new ListItem("14) BIC Banco BIC, S.A. >> PDF 16/m", 9));//9
    //    chkLstProducts.Items.Add(new ListItem("21) ZEN Zenith Bank PLC >> MS Access MDB 16/m", 17));//17
    //    chkLstProducts.Items.Add(new ListItem("21) ZEN Zenith Bank PLC >> Default Text for Email 16/m", 66));//66
    //    chkLstProducts.Items.Add(new ListItem("[21] ZEN Zenith Bank PLC >> Email 16/m", 41));//41
    //    chkLstProducts.Items.Add(new ListItem("22) BPC BANCO DE POUPANCA E CREDITO, SARL Credit >> PDF 16/m", 14));//14
    //    chkLstProducts.Items.Add(new ListItem("[22] BPC BANCO DE POUPANCA E CREDITO, SARL Credit >> Email 16/m", 144));//144
    //    chkLstProducts.Items.Add(new ListItem("31) NBS NBS Bank Limited >> PDF 16/m", 40));//40
    //    chkLstProducts.Items.Add(new ListItem("[31] NBS NBS Bank Limited >> Email 16/m", 145));//145
    //    //chkLstProducts.Items.Add(new ListItem("33) ZENG Zenith Bank (Ghana) Limited >> Default Text 16/m", 42));//42
    //    chkLstProducts.Items.Add(new ListItem("33) ZENG Zenith Bank (Ghana) Limited >> Default PDF 16/m", 42));//42
    //    chkLstProducts.Items.Add(new ListItem("36) DBN Diamond Bank Nigeria >> Reward Default Text 16/m", 46));//46
    //    chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> Email 16/m", 48));//48
    //    chkLstProducts.Items.Add(new ListItem("36) DBN Diamond Bank Nigeria - VIP >> Reward Default Text 16/m", 97));//97
    //    chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria - VIP >> Email 16/m", 98));//98
    //    chkLstProducts.Items.Add(new ListItem("36) DBN Diamond Bank Nigeria >> VISA Platinum - ParkNShop Co-Brand Text 16/m", 111));//111
    //    chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> VISA Platinum - ParkNShop Co-Brand Email 16/m", 112));//112
    //    chkLstProducts.Items.Add(new ListItem("36) DBN Diamond Bank Nigeria >> VISA Platinum - EXCO_VIP-Brand Text 16/m", 149));//149
    //    chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> VISA Platinum - EXCO_VIP-Brand Email 16/m", 150));//150 
    //    chkLstProducts.Items.Add(new ListItem("36) DBN Diamond Bank Nigeria >> MasterCard Credit Text 16/m", 220));//220
    //    chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> MasterCard Credit Email 16/m", 221));//221
    //    //chkLstProducts.Items.Add(new ListItem("38) GTBN Guaranty trust bank plc nigeria  >> Credit Default Text 1/m", 161));//161
    //    chkLstProducts.Items.Add(new ListItem("38) GTBN Guaranty trust bank plc nigeria  >> Credit Default PDF 16/m", 161));//161
    //    chkLstProducts.Items.Add(new ListItem("[38] GTBN GUARANTY TRUST BANK PLC NIGERIA  >> Credit Emails 16/m", 162));//162
    //    chkLstProducts.Items.Add(new ListItem("55) FBP Fidelity Bank PLC >> Default Text 16/m", 95));//95
    //    chkLstProducts.Items.Add(new ListItem("[55] FBP Fidelity Bank PLC >> Emails 16/m", 96));//96
    //    chkLstProducts.Items.Add(new ListItem("55) FBP Fidelity Bank PLC >> Default Text Debit 16/m", 151));//151
    //    chkLstProducts.Items.Add(new ListItem("58) SBP SKYE BANK PLC >> Default Text 16/m", 79));//79
    //    chkLstProducts.Items.Add(new ListItem("[58] SBP SKYE BANK PLC >> Emails 16/m", 80));//80
    //    chkLstProducts.Items.Add(new ListItem("[58] SBP SKYE BANK PLC >> MasterCard Platinum Credit Emails 16/m", 253));//253
    //    //chkLstProducts.Items.Add(new ListItem("58) SBP SKYE BANK PLC >> Default Debit Text 16/m", 90));//90
    //    chkLstProducts.Items.Add(new ListItem("69) FBN First Bank of Nigeria  >> Credit Default Text 16/m", 123));//123
    //    chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Credit Emails 16/m", 124));//124
    //    chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Supplementary Credit Emails 16/m", 152));//152
    //    chkLstProducts.Items.Add(new ListItem("69) FBN First Bank of Nigeria  >> Prepaid Default Text 16/m", 115));//115
    //    chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Prepaid Emails 16/m", 117));//117
    //    chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> MasterCard Prepaid Emails 16/m", 203));//203
    //    chkLstProducts.Items.Add(new ListItem("77) I&M Bank Rwanda Limited >> Credit PDF 15/m", 201));//201
    //    //chkLstProducts.Items.Add(new ListItem("77) I&M Bank Rwanda Limited >> Debit PDF 15/m", 202));//202
    //    chkLstProducts.Items.Add(new ListItem("85) UTBG UT Bank Limited Ghana >> Default Credit PDF 1/m", 167));//167
    //    chkLstProducts.Items.Add(new ListItem("87) WEMA BANK PLC NIGERIA >> Credit Default 15/m", 179));//179
    //    chkLstProducts.Items.Add(new ListItem("[87] WEMA BANK PLC NIGERIA >> Credit Emails 15/m", 180));//180
    //    //chkLstProducts.Items.Add(new ListItem("87) WEMA BANK PLC NIGERIA >> Debit/Prepaid Default 15/m", 182));//182
    //    //chkLstProducts.Items.Add(new ListItem("[87] WEMA BANK PLC NIGERIA >> Debit/Prepaid Emails 15/m", 183));//183
    //    chkLstProducts.Items.Add(new ListItem("98) GTBK Guaranty trust bank Kenya  >> Credit PDF 1/m", 228));//228
    //    chkLstProducts.Items.Add(new ListItem("[98] GTBK Guaranty trust bank Kenya  >> Credit Emails 1/m", 230));//230
    //    chkLstProducts.Items.Add(new ListItem("106) DSBJ EAST AFRICA BANK Djibouti  >> Debit Text 15/m", 233));//233
    //    chkLstProducts.Items.Add(new ListItem("106) DSBJ EAST AFRICA BANK Djibouti  >> Credit Islamic Text 15/m", 234));//234

    //    chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default Credit Text 17/m", 214));//214

    //    chkLstProducts.Items.Add(new ListItem("51) TMB Trust Merchant Bank >> Raw VISA Text 20/m", 75));//75
    //    chkLstProducts.Items.Add(new ListItem("51) TMB Trust Merchant Bank >> Raw MasterCard Text 20/m", 204));//204
    //    chkLstProducts.Items.Add(new ListItem("51) TMB Trust Merchant Bank >> PDF 20/m", 108));//108
    //    chkLstProducts.Items.Add(new ListItem("51) TMB Trust Merchant Bank >> Corporate Text 20/m", 105));//105
    //    chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default Credit Text 20/m", 133));//133
    //    chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default Debit Text 20/m", 134));//134

    //    chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default Credit Text 23/m", 215));//215

    //    chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default Credit Text 27/m", 216));//216

    //    chkLstProducts.Items.Add(new ListItem("1) QNB ALAHLI Credit >> Text Splitted by product 1/m", 1));//1
    //    //chkLstProducts.Items.Add(new ListItem("[1] NSGB Credit >> Email 1/m", 54));//54
    //    chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> Classic Email 1/m", 169));//169
    //    chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> Gold Email 1/m", 170));//170
    //    chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> Platinum Email 1/m", 171));//171
    //    chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> Infinite Email 1/m", 172));//172
    //    //chkLstProducts.Items.Add(new ListItem("1) NSGB Individual >> Text 1/m", 19));//19
    //    chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> Business Individual Email 1/m", 56));//56
    //    //chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> MasterCard Standard Email 1/m", 174));//174
    //    chkLstProducts.Items.Add(new ListItem("1) QNB ALAHLI Business >> Text 1/m", 2));//2
    //    chkLstProducts.Items.Add(new ListItem("1) QNB ALAHLI Business - SME >> Text 1/m", 78));//78
    //    chkLstProducts.Items.Add(new ListItem("1) QNB ALAHLI Mastercard Business - SME >> Text 1/m", 148));//148
    //    chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Mastercard Business - SME >> Email 1/m", 236));//236
    //    //chkLstProducts.Items.Add(new ListItem("[1] NSGB Business >> Email 1/m", 55));//55
    //    chkLstProducts.Items.Add(new ListItem("1) QNB ALAHLI MasterCard Salary Prepaid >> PDF 1/m", 114));//114
    //    chkLstProducts.Items.Add(new ListItem("3) NBK VISA Classic >> Text 1/m", 3));//3
    //    chkLstProducts.Items.Add(new ListItem("3) NBK VISA Gold >> Text 1/m", 60));//60
    //    chkLstProducts.Items.Add(new ListItem("4) BMSR Bank Misr Credit Classic >> Text 1/m", 16));//16
    //    chkLstProducts.Items.Add(new ListItem("4) BMSR Bank Misr Credit Gold >> Text 1/m", 88));//88
    //    chkLstProducts.Items.Add(new ListItem("4) BMSR Bank Misr Credit Mobinet >> Text 1/m", 89));//89
    //    //chkLstProducts.Items.Add(new ListItem("4) BMSR Bank Misr Save >> Text 1/m", 12));//12
    //    chkLstProducts.Items.Add(new ListItem("5) AIB Raw Data VISA USD>> Credit Text 1/m", 4));//4
    //    chkLstProducts.Items.Add(new ListItem("5) AIB Raw Data VISA EUR>> Credit Text 1/m", 64));//64
    //    chkLstProducts.Items.Add(new ListItem("5) AIB Raw Data VISA EGP>> Credit Text 1/m", 187));//187
    //    chkLstProducts.Items.Add(new ListItem("5) AIB Raw Data MasterCard USD >> Credit Text 1/m", 237));//237
    //    chkLstProducts.Items.Add(new ListItem("5) AIB PDF >> Credit PDF 1/m", 107));//107
    //    chkLstProducts.Items.Add(new ListItem("5) AIB Electron EGP >> Debit Text 1/m", 188));//188
    //    //chkLstProducts.Items.Add(new ListItem("6) AAIB Arab African International Bank >> Text 1/m", 11));//11
    //    //chkLstProducts.Items.Add(new ListItem("6) AAIB Arab African International Bank >> Prepaid Text 1/m", 189));//189
    //    //chkLstProducts.Items.Add(new ListItem("[6] AAIB Arab African International Bank >> Prepaid Email 1/m", 190));//190
    //    chkLstProducts.Items.Add(new ListItem("6) AAIB Arab African International Bank >> Raw Text 1/m", 205));//205
    //    chkLstProducts.Items.Add(new ListItem("6) AAIB Arab African International Bank >> Text Spool 1/m", 217));//217
    //    chkLstProducts.Items.Add(new ListItem("7) BAI Prducts >> PDF 1/m", 7));//7
    //    chkLstProducts.Items.Add(new ListItem("[7] BAI Prducts >> Email 1/m", 85));//85
    //    chkLstProducts.Items.Add(new ListItem("7) BAI Prepaid >> PDF 1/m", 127));//127
    //    chkLstProducts.Items.Add(new ListItem("[7] BAI Prepaid >> Email 1/m", 128));//128
    //    //chkLstProducts.Items.Add(new ListItem("7) BAI Corporate >> Text 1/m", 140));//140
    //    chkLstProducts.Items.Add(new ListItem("7) BAI Corporate >> PDF 1/m", 140));//140
    //    chkLstProducts.Items.Add(new ListItem("[7] BAI Corporate >> Email 1/m", 181));//181
    //    //chkLstProducts.Items.Add(new ListItem("XXX 10) BNP Credit >> Text 1/m", 6));//6
    //    //chkLstProducts.Items.Add(new ListItem("10) BNP Credit Reward >> Text 1/m", 37));//37
    //    //chkLstProducts.Items.Add(new ListItem("10) BNP Corporate >> Text 1/m", 36));//36
    //    //chkLstProducts.Items.Add(new ListItem("10)[Development Cost not Paid] BNP Debit >> Text 1/m", 86));//86
    //    chkLstProducts.Items.Add(new ListItem("12) BSIC Alwaha Bank Libya Inc (Oasis Bank) >> PDF 1/m", 50));//50
    //    //chkLstProducts.Items.Add(new ListItem("13) BDK Bank De Kigali >> Default Debit Text 1/m", 100));//100
    //    //chkLstProducts.Items.Add(new ListItem("[13] BDK Bank De Kigali >> Emails 1/m", 101));//101
    //    chkLstProducts.Items.Add(new ListItem("13) BK Bank of Kigali >> Default Credit Text 1/m", 154));//154
    //    chkLstProducts.Items.Add(new ListItem("[13] BK Bank of Kigali >> Credit Emails 1/m", 155));//155
    //    //chkLstProducts.Items.Add(new ListItem("15) SSB SG-SSB Ltd Corporate >> PDF 1/m", 18));//18
    //    chkLstProducts.Items.Add(new ListItem("15) SSB SG-SSB Ltd Debit >> PDF 1/m", 74));//74
    //    //chkLstProducts.Items.Add(new ListItem("15) SSB SG-SSB Ltd Debit >> Text 1/m", 74));//74
    //    chkLstProducts.Items.Add(new ListItem("16) ABP Access Bank Plc with Label Credit >> Text 1/m", 10));//10
    //    //chkLstProducts.Items.Add(new ListItem("XXX 16) ABP Access Bank Plc with Label Credit >> Reward Default Text 1/m", 81));//81
    //    chkLstProducts.Items.Add(new ListItem("16) ABP Access Bank Plc with Label Dual Currency >> Text 1/m", 35));//35
    //    chkLstProducts.Items.Add(new ListItem("16) ABP Access Bank Plc with Label Infinite Dual Currency >> Text 1/m", 61));//61
    //    chkLstProducts.Items.Add(new ListItem("16) ABP Access Bank Plc with Label Corporate >> Text 1/m", 43));//43
    //    chkLstProducts.Items.Add(new ListItem("[16] ABP Access Bank Plc  >> Credit Emails 1/m", 211));//211
    //    chkLstProducts.Items.Add(new ListItem("[16] ABP Access Bank Plc  >> Corporate Cardholder Emails 1/m", 225));//225
    //    //chkLstProducts.Items.Add(new ListItem("16) ABP Access Bank Plc with Label Corporate >> PDF 1/m", 160));//160
    //    //chkLstProducts.Items.Add(new ListItem("18) BOAB Bank of Africa Benin >> Debit >> PDF 1/m", 65));//65
    //    chkLstProducts.Items.Add(new ListItem("18) BOAB Bank of Africa Benin >> Default Credit Text 1/m", 185));//185
    //    //chkLstProducts.Items.Add(new ListItem("19) BCNS Banco Cabov. de Negocios >> Credit BUSINESS >> PDF 1/m", 21));//21
    //    //chkLstProducts.Items.Add(new ListItem("19) BCNS Banco Cabov. de Negocios >> Debit Pre Travel >> PDF 1/m", 22));//22
    //    chkLstProducts.Items.Add(new ListItem("20) BICV Banco Intertlantico >> Credit Classic >> PDF 1/m", 23));//23
    //    chkLstProducts.Items.Add(new ListItem("20) BICV Banco Intertlantico >> Credit Classic >> Text 1/m", 165));//23
    //    chkLstProducts.Items.Add(new ListItem("20) BICV Banco Intertlantico >> Debit >> PDF 1/m", 25));//25
    //    chkLstProducts.Items.Add(new ListItem("20) BICV Banco Intertlantico >> Debit >> Text 1/m", 166));//25
    //    chkLstProducts.Items.Add(new ListItem("22) BPC BANCO DE POUPANCA E CREDITO, SARL Corporate >> Credit PDF 1/m", 52));//52
    //    chkLstProducts.Items.Add(new ListItem("[22] BPC BANCO DE POUPANCA E CREDITO, SARL Credit >> Cardholder Email 16/m", 186));//186
    //    chkLstProducts.Items.Add(new ListItem("22) BPC BANCO DE POUPANCA E CREDITO, SARL Debit >> Debit PDF 1/m", 191));//191
    //    chkLstProducts.Items.Add(new ListItem("[22] BPC BANCO DE POUPANCA E CREDITO, SARL Debit >> Email 1/m", 192));//192
    //    chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) Credit Gold Standard>> PDF 1/m", 26));//26
    //    chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) Credit Classic Standard>> PDF 1/m", 27));//27
    //    chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) MC Credit Gold Standard>> PDF 1/m", 196));//196
    //    chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) MC Credit Classic Standard>> PDF 1/m", 197));//197
    //    chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) Credit Gold Staff>> PDF 1/m", 242));//242
    //    chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) Credit Classic Staff>> PDF 1/m", 243));//243
    //    chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) MC Credit Gold Staff>> PDF 1/m", 244));//244
    //    chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) MC Credit Classic Staff>> PDF 1/m", 245));//245
    //    //chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) >> Text 1/m", 249));//249
    //    chkLstProducts.Items.Add(new ListItem("24) BOAL Bank of Africa Mali >> Credit >> French PDF 1/m", 168));//168
    //    chkLstProducts.Items.Add(new ListItem("24) BOAL Bank of Africa Mali >> Debit >> Default PDF 1/m", 76));//76
    //    //chkLstProducts.Items.Add(new ListItem("XXX 25) AUB Ahli United Bank >> PDF 1/m", 13));// 26/m  13
    //    chkLstProducts.Items.Add(new ListItem("25) AUB Ahli United Bank >> Text 1/m", 51));// 51
    //    chkLstProducts.Items.Add(new ListItem("26) BOAS Bank of Africa Senegal >> Credit >> French PDF 1/m", 199));//199
    //    chkLstProducts.Items.Add(new ListItem("27) BCA BANCO COMERCIAL DO ATLANTICO >> Credit >> PDF 1/m", 20));//20
    //    chkLstProducts.Items.Add(new ListItem("27) BCA BANCO COMERCIAL DO ATLANTICO >> Corporate >> PDF 1/m", 70));//70
    //    chkLstProducts.Items.Add(new ListItem("27) BCA BANCO COMERCIAL DO ATLANTICO >> Debit >> Default Text 1/m", 69));//69
    //    chkLstProducts.Items.Add(new ListItem("28) CECV CAIXA ECONOMICA DE CABO VERDE >> Credit Classic >> PDF 1/m", 24));//24
    //    chkLstProducts.Items.Add(new ListItem("28) CECV CAIXA ECONOMICA DE CABO VERDE >> Credit Classic >> Text 1/m", 163));//24
    //    chkLstProducts.Items.Add(new ListItem("28) CECV CAIXA ECONOMICA DE CABO VERDE >> Debit - Prepaid >> PDF 1/m", 91));//91
    //    chkLstProducts.Items.Add(new ListItem("28) CECV CAIXA ECONOMICA DE CABO VERDE >> Debit - Prepaid >> Text 1/m", 164));//91
    //    //chkLstProducts.Items.Add(new ListItem("29) UBA United Bank for Africa Plc Nigeria >> Default Text 1/m", 8));//8
    //    chkLstProducts.Items.Add(new ListItem("29) UBA United Bank for Africa Plc Nigeria >> PDF 1/m", 8));//8
    //    chkLstProducts.Items.Add(new ListItem("29) UBA United Bank for Africa Plc Nigeria >> Emails 1/m", 45));//45
    //    //chkLstProducts.Items.Add(new ListItem("30) BMG Banque Misr Gulf >> Text 1/m", 73));//73
    //    //chkLstProducts.Items.Add(new ListItem("32) UBAG UNITED BANK FOR AFRICA GHANA LIMITED >> PDF 1/m", 31));//31
    //    chkLstProducts.Items.Add(new ListItem("32) UBAG UNITED BANK FOR AFRICA GHANA LIMITED >> Debit PDF 1/m", 193));//193
    //    chkLstProducts.Items.Add(new ListItem("[32] UBAG UNITED BANK FOR AFRICA GHANA LIMITED >> Debit Email 1/m", 194));//194
    //    chkLstProducts.Items.Add(new ListItem("35) BOAC Bank of Africa Cote D'Ivoire >> Credit >> French PDF 1/m", 198));//198
    //    chkLstProducts.Items.Add(new ListItem("36) DBN Diamond Bank Nigeria - VIP 1,2,5 >> Reward Default Text 1/m", 83));//83
    //    chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria - VIP 1,2,5 >> Email 1/m", 84));//84
    //    chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria - VIP 1,2,5 Supplementary>> Email 1/m", 102));//102
    //    //chkLstProducts.Items.Add(new ListItem("37) PHB PLATINUM HABIB BANK PLC >> Default Text 1/m", 57));//57
    //    //chkLstProducts.Items.Add(new ListItem("[37] PHB PLATINUM HABIB BANK PLC >> Emails 1/m", 62));//62
    //    //chkLstProducts.Items.Add(new ListItem("37) KBL Keystone Bank >> Default Text 1/m", 57));//57
    //    //chkLstProducts.Items.Add(new ListItem("[37] KBL Keystone Bank >> Emails 1/m", 62));//62
    //    //chkLstProducts.Items.Add(new ListItem("39) AFN Afribank Nigeria PLC >> Default Credit Text 1/m", 53));//53
    //    //chkLstProducts.Items.Add(new ListItem("39) AFN Afribank Nigeria PLC >> Default Debit Text 1/m", 68));//68
    //    //chkLstProducts.Items.Add(new ListItem("39) MBL Mainstreet Bank Limited >> Default Credit Text 1/m", 53));//53
    //    //chkLstProducts.Items.Add(new ListItem("39) MBL Mainstreet Bank Limited >> Default Debit Text 1/m", 68));//68
    //    //chkLstProducts.Items.Add(new ListItem("40) ICB INTERCONTINENTAL BANK PLC >> Credit Default Text 1/m", 44));//44
    //    //chkLstProducts.Items.Add(new ListItem("40) ICB INTERCONTINENTAL BANK PLC >> Corporate Default Text 1/m", 71));//71
    //    chkLstProducts.Items.Add(new ListItem("42) OBI Oceanic Bank International >> Default Credit Text 1/m", 206));//206
    //    chkLstProducts.Items.Add(new ListItem("42) OBI Oceanic Bank International >> Default Debit Text 1/m", 119));//119
    //    chkLstProducts.Items.Add(new ListItem("44) BOAK Bank of Africa Kenya >> Default Debit Text 1/m", 248));//248
    //    chkLstProducts.Items.Add(new ListItem("45) BOCD Bank Of Commerce and Development Debit >> PDF 1/m", 63));//63
    //    //chkLstProducts.Items.Add(new ListItem("47) SBN Sterling Bank Nigeria >> Default Debit Text 1/m", 120));//120
    //    chkLstProducts.Items.Add(new ListItem("47) SBN Sterling Bank Nigeria >> Default Credit Text 1/m", 222));//222
    //    chkLstProducts.Items.Add(new ListItem("[47] SBN Sterling Bank Nigeria >> Default Credit Email 1/m", 223));//223
    //    chkLstProducts.Items.Add(new ListItem("50) ICBG Intercontinental Bank Ghana Limited Credit >> PDF 1/m", 99));//99
    //    chkLstProducts.Items.Add(new ListItem("50) ICBG Intercontinental Bank Ghana Limited Corporate >> PDF 1/m", 106));//106
    //    chkLstProducts.Items.Add(new ListItem("50) ICBG Intercontinental Bank Ghana Limited Debit >> Text 1/m", 87));//87
    //    chkLstProducts.Items.Add(new ListItem("50) ICBG Intercontinental Bank Ghana Limited Prepaid >> PDF 1/m", 113));//113
    //    //chkLstProducts.Items.Add(new ListItem("53) RCB ROKEL COMMERCIAL BANK S/L LTD >> Debit PDF 1/m", 82));//82 //[Development Cost not Paid]
    //    chkLstProducts.Items.Add(new ListItem("55) FBP Fidelity Bank PLC >> Default Text 1/m", 218));//218
    //    chkLstProducts.Items.Add(new ListItem("[55] FBP Fidelity Bank PLC >> Emails 1/m", 219));//219
    //    chkLstProducts.Items.Add(new ListItem("56) NCB National Commercial Bank >> Credit >> PDF 1/m", 77));//77
    //    chkLstProducts.Items.Add(new ListItem("[56] NCB National Commercial Bank >> Credit >> Email 1/m", 224));//224
    //    chkLstProducts.Items.Add(new ListItem("56) NCB National Commercial Bank >> Prepaid >> PDF 1/m", 226));//226
    //    chkLstProducts.Items.Add(new ListItem("[56] NCB National Commercial Bank >> Prepaid >> Email 1/m", 227));//227
    //    chkLstProducts.Items.Add(new ListItem("58) SBP SKYE BANK PLC >> Default Credit Text 1/m", 92));//92
    //    chkLstProducts.Items.Add(new ListItem("[58] SBP SKYE BANK PLC >> Emails 1/m", 93));//93
    //    //chkLstProducts.Items.Add(new ListItem("58) SBP SKYE BANK PLC >> Default Debit Text 1/m", 94));//94
    //    chkLstProducts.Items.Add(new ListItem("58) SBP SKYE BANK PLC >> Default Prepaid Text 1/m", 175));//175
    //    chkLstProducts.Items.Add(new ListItem("[58] SBP SKYE BANK PLC >> Prepaid VISA Emails 1/m", 176));//176
    //    chkLstProducts.Items.Add(new ListItem("[58] SBP SKYE BANK PLC >> Prepaid VISA-NTDC Emails 1/m", 184));//184
    //    chkLstProducts.Items.Add(new ListItem("[58] SBP SKYE BANK PLC >> Prepaid MasterCard Emails 1/m", 177));//177
    //    chkLstProducts.Items.Add(new ListItem("58) SBP SKYE BANK PLC >> Default Corporate Text 1/m", 125));//125
    //    chkLstProducts.Items.Add(new ListItem("[58] SBP SKYE BANK PLC >> Corporate Emails 1/m", 126));//126
    //    //chkLstProducts.Items.Add(new ListItem("60) GUM GUMHOURIA Bank  >> Prepaid PDF 1/m", 207));//207
    //    chkLstProducts.Items.Add(new ListItem("62) WHDA WHDA Bank  >> Prepaid PDF 1/m", 200));//200
    //    chkLstProducts.Items.Add(new ListItem("65) SBPG SKYE BANK PLC Gambia >> Default Debit Text 1/m", 121));//121
    //    chkLstProducts.Items.Add(new ListItem("66) ABPG Access Bank Plc Gambia >> Default Debit Text 1/m", 129));//129
    //    //chkLstProducts.Items.Add(new ListItem("67) ABPC Access Bank Plc Congo >> PDF 1/m", 130));//130
    //    //chkLstProducts.Items.Add(new ListItem("69) FBN First Bank of Nigeria  >> Credit Default Text 1/m", 109));//109
    //    //chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Credit Emails 1/m", 110));//110
    //    //chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Supplementary Credit Emails 1/m", 153));//153
    //    //chkLstProducts.Items.Add(new ListItem("69) FBN First Bank of Nigeria  >> Prepaid Default Text 1/m", 115));//115
    //    //chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Prepaid Emails 1/m", 117));//117
    //    chkLstProducts.Items.Add(new ListItem("69) FBN First Bank of Nigeria  >> Corporate Default Text 1/m", 116));//116
    //    chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Corporate Cardholder Emails 1/m", 122));//122
    //    //chkLstProducts.Items.Add(new ListItem("69) FBN First Bank of Nigeria  >> MasterCard Credit Standard Default Text 1/m", 142));//142
    //    //chkLstProducts.Items.Add(new ListItem("69) FBN First Bank of Nigeria  >> MasterCard Credit Standard Default Emails 1/m", 147));//142
    //    chkLstProducts.Items.Add(new ListItem("71) BOAD BANK OF AFRICA DEMOCRATIC REPUBLIC OF CONGO >> Default Debit Text 1/m", 136));//136
    //    chkLstProducts.Items.Add(new ListItem("72) ABPR ACCESS BANK RWANDA SA >> Default Credit Text 1/m", 131));//131
    //    chkLstProducts.Items.Add(new ListItem("72) ABPR ACCESS BANK RWANDA SA >> Default Debit Text 1/m", 132));//132
    //    chkLstProducts.Items.Add(new ListItem("74) BLME Blom Bank Egypt VISA>> Credit Text 1/m", 143));//143
    //    chkLstProducts.Items.Add(new ListItem("74) BLME Blom Bank Egypt VISA>> Credit PDF 1/m", 135));//135
    //    chkLstProducts.Items.Add(new ListItem("74) BLME Blom Bank Egypt MasterCard>> Credit Text 1/m", 141));//141
    //    chkLstProducts.Items.Add(new ListItem("74) BLME Blom Bank Egypt MasterCard>> Credit PDF 1/m", 178));//178
    //    //chkLstProducts.Items.Add(new ListItem("75) UMB UNIVERSAL MERCHANT BANK GHANA >> Default Debit Text 1/m", 137));//137
    //    chkLstProducts.Items.Add(new ListItem("76) EDBE EXPORT DEVELOPMENT BANK OF EGYPT >> Default Credit Text 1/m", 146));//146
    //    //chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Corporate Company Emails 1/m", 118));//118
    //    chkLstProducts.Items.Add(new ListItem("81) SIBN STANBIC IBTC BANK NIGERIA  >> Credit Default Text 1/m", 156));//156
    //    //chkLstProducts.Items.Add(new ListItem("[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Emails 1/m", 157));//157
    //    chkLstProducts.Items.Add(new ListItem("[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Platinum Emails 1/m", 208));//208
    //    chkLstProducts.Items.Add(new ListItem("[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Gold Emails 1/m", 209));//209
    //    chkLstProducts.Items.Add(new ListItem("[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Infinite Emails 1/m", 210));//210
    //    chkLstProducts.Items.Add(new ListItem("81) SIBN STANBIC IBTC BANK NIGERIA  >> Corporate Default Text 1/m", 158));//158
    //    chkLstProducts.Items.Add(new ListItem("[81] STANBIC IBTC BANK NIGERIA >> Corporate Cardholder Emails 1/m", 159));//159
    //    chkLstProducts.Items.Add(new ListItem("90) UBG UniBank Ghana Limited  >> Debit PDF 1/m", 195));//195
    //    chkLstProducts.Items.Add(new ListItem("95) GTBL Guaranty trust bank Liberia  >> Debit Text 1/m", 229));//229
    //    //chkLstProducts.Items.Add(new ListItem("98) GTBK Guaranty trust bank Kenya  >> Credit PDF 15/m", 228));//228
    //    //chkLstProducts.Items.Add(new ListItem("[98] GTBK Guaranty trust bank Kenya  >> Credit Emails 15/m", 230));//230
    //    chkLstProducts.Items.Add(new ListItem("102) VCBK Victoria Commercial Bank Ltd. Kenya  >> Credit Text 1/m", 231));//231
    //    chkLstProducts.Items.Add(new ListItem("102) VCBK Victoria Commercial Bank Ltd. Kenya  >> Raw Text 1/m", 235));//235
    //    //chkLstProducts.Items.Add(new ListItem("104) DSBS Dahabshil Bank International Somalia  >> Debit Text 1/m", 232));//232
    //    chkLstProducts.Items.Add(new ListItem("109) CGBK COMPAGNIE GENERALE DE BANQUE LTD >> MasterCard Credit Text 1/m", 238));//238
    //    chkLstProducts.Items.Add(new ListItem("109) CGBK COMPAGNIE GENERALE DE BANQUE LTD >> MasterCard Prepaid Text 1/m", 3));//239
    //    //chkLstProducts.Items.Add(new ListItem("110) HBLN Heritage Banking Company Nigeria >> MasterCard Credit Text 1/m", 240));//240
    //    //chkLstProducts.Items.Add(new ListItem("110) HBLN Heritage Banking Company Nigeria >> Prepaid Text 1/m", 246));//246
    //    //chkLstProducts.Items.Add(new ListItem("[110] HBLN Heritage Banking Company Nigeria >> MasterCard Credit Emails 1/m", 241));//241
    //    //chkLstProducts.Items.Add(new ListItem("[110] HBLN Heritage Banking Company Nigeria >> MasterCard Prepaid Emails 1/m", 247));//247
    //    //chkLstProducts.Items.Add(new ListItem("112) GTBR Guaranty trust bank Rwanda >> Credit Text 1/m", 250));//250
    //    //chkLstProducts.Items.Add(new ListItem("112) GTBR Guaranty trust bank Rwanda  >> Debit Text 1/m", 251));//251
    //    //chkLstProducts.Items.Add(new ListItem("112) GTBR Guaranty trust bank Rwanda  >> Corporate Text 1/m", 252));//252
    //    //chkLstProducts.Items.Add(new ListItem("113) GTBU Guaranty trust bank Uganda  >> Debit Text 1/m", 254));//254


    //    //chkLstProducts.Items.Add(new ListItem("Common English >> PDF", 0));//0
    //    //chkLstProducts.Items.Add(new ListItem("Common English - Nice 1 >> PDF", 5));//5
    //    //chkLstProducts.Items.Add(new ListItem("Common English - Nice 2 >> PDF", 30));//30
    //    //chkLstProducts.Items.Add(new ListItem("Common English Arabic >> PDF", 28));//28
    //    //chkLstProducts.Items.Add(new ListItem("Common English Debit >> PDF", 32));//32
    //    //chkLstProducts.Items.Add(new ListItem("Common English Reward >> PDF", 34));//34
    //    //chkLstProducts.Items.Add(new ListItem("Common Portuguese >> PDF", 29));//29
    //    //chkLstProducts.Items.Add(new ListItem("Common Corporate >> PDF", 33));//33
    //    //chkLstProducts.Items.Add(new ListItem("Common Credit English Text - No Charge", 38));//38
    //    //chkLstProducts.Items.Add(new ListItem("Common Reward English Text - No Charge", 39));//39
    //    //chkLstProducts.Items.Add(new ListItem("Common Corporate Detail English Text - No Charge", 72));//72
    //    //chkLstProducts.Items.Add(new ListItem("Common Debit English Text - No Charge", 67));//67
    //    //chkLstProducts.Items.Add(new ListItem("Test Statement", 59));//59
    //    }

    private void runStatement(int pCmbProducts)
    {

        try
        {

            if (!clsDbOracleLayer.doActionGrantProc("STMT.ZM_STMT_APP.GrantTable", txtTblMaster.Text.Substring(16, 11))) ;
            setCurrentBank();
            BeginInvoke(setStatusDelegate, new object[] { bankCode, "Start Create Statement for " + strStatementType });//strStatementType


            switch (pCmbProducts)//cmbProducts.SelectedIndex
            {
                // Create PDF statement
                case 0:    // Common
                case 5:    // Statement_Common_English_Nice1 
                case 28:   // Common Statement English Arabic
                case 29:   // Common Statement Portuguese
                case 30:   // Statement Common English Nice2 
                case 32:   // Common English Debit
                case 114: // 1] NSGB MasterCard Salary Prepaid >> PDF 1/m
                case 107:  //  5) AIB Credit >> PDF 1/m
                case 50:   // 12) BSIC Alwaha Bank Libya Inc (Oasis Bank) >> PDF 1/m
                case 9:    // 14) BIC
                //case 74:   // 15) SSB SG-SSB Ltd Debit >> PDF 1/m
                case 65:   // 18) BOAB Bank of Africa Benin >> Debit >> PDF 1/m
                case 21:   // 19) BCNS
                case 22:   // 19) BCNS Debit
                case 25:   // 20) BICV Debit
                case 23:   // 20) BICV
                case 14:   // 22) BANCO DE POUPANCA E CREDITO, SARL >> BPC  >> 22
                case 191: // 22) BPC BANCO DE POUPANCA E CREDITO, SARL Debit >> PDF 1/m      
                case 76:   // 24) BOAL Bank of Africa Mali >> Debit >> Default PDF 1/m
                case 168:  // 24) BOAL Bank of Africa Mali >> Credit >> French PDF 1/m
                case 199:  // 35) BOAS Bank of Africa Senegal >> Credit >> French PDF 1/m
                case 20:   // 27) BCA
                case 24:   // 28) CECV CAIXA ECONOMICA DE CABO VERDE >> Credit Classic >> PDF 1/m
                case 91:   // 28) CECV CAIXA ECONOMICA DE CABO VERDE >> Debit - Prepaid >> PDF 1/m
                case 8:    // 29) UBA United Bank for Africa Plc Nigeria PDF
                case 436:    // 29) UBA United Bank for Africa Plc Nigeria PDF
                case 40:   // 31) NBS NBS Bank Limited >> PDF 16/m
                case 31:   // 32) UBAG UNITED BANK FOR AFRICA GHANA LIMITED
                case 193: //32) UBAG UNITED BANK FOR AFRICA GHANA LIMITED >> Debit PDF 1/m
                //case 198:  // 35) BOAC Bank of Africa Cote D'Ivoire >> Credit >> French PDF 1/m
                //case 42:   // 33) ZENG Zenith Bank (Ghana) Limited >> Default Text 16/m
                case 161:  //38) GTBN GUARANTY TRUST BANK PLC NIGERIA  >> Credit Default PDF 16/m
                case 63:   // 45) BOCD Bank Of Commerce and Development >> PDF 1/m
                case 99:   // 50) ICBG Intercontinental Bank Ghana Limited Credit 1/m
                case 113:  // 50) ICBG Intercontinental Bank Ghana Limited Prepaid >> PDF 1/m
                case 108:  // 51] TMB Trust Merchant Bank >> PDF 20/m
                case 82:   // 53) RCB ROKEL COMMERCIAL BANK S/L LTD >> PDF 1/m
                case 77:   // 56) NCB National Commercial Bank >> Credit >> PDF 1/m
                case 226:   // 56) NCB National Commercial Bank >> Prepaid >> PDF 1/m
                case 207: //   60) GUM GUMHOURIA Bank  >> Prepaid PDF 1/m
                //case 200: //  62) WHDA WHDA Bank  >> Prepaid PDF 1/m", 200));//200
                case 130:  // 67) ABPC Access Bank Plc Congo >> PDF 1/m
                case 7100:
                case 135:  // 74) BLME Blom Bank Egypt VISA>> Credit PDF 1/m
                case 178:  // 74) BLME Blom Bank Egypt MasterCard>> Credit PDF 1/m
                case 201: //  77) I&M Bank Rwanda Limited  >> Credit PDF 15/m", 201));//201
                case 202: //  77) I&M Bank Rwanda Limited  >> Prepaid PDF 15/m", 202));//202
                case 167:  //85) UTBG UT Bank Limited Ghana >> Default Credit PDF 1/m
                case 195:  //90) UBG UniBank Ghana Limited  >> Debit PDF 1/m
                case 467:   //94) FABG First Atlantic Bank Ghana >> Credit PDF 1/m
                case 228: //  98) GTBK Guaranty trust bank Kenya  >> Credit PDF 15/m
                case 256: //  98) GTBK Guaranty trust bank Kenya  >> Prepaid PDF 15/m
                case 295:   // 128) EGB Egyptian Gulf Bank of Egypt  >> Credit OMR PDF 30/m
                case 292:  // 127) AIBK Arab Investment Bank>> Credit PDF 1/m
                case 293:  // 127) AIBK Arab Investment Bank>> Credit Customer PDF 1/m
                case 306:  // 127) AIBK Arab Investment Bank>> Credit Staff PDF 1/m
                case 307:  // 127) AIBK Arab Investment Bank>> Installment Customer PDF 1/m
                case 308:  // 127) AIBK Arab Investment Bank>> Installment Staff PDF 1/m
                case 309:  // 127) AIBK Arab Investment Bank>> Installment PDF 1/m
                case 310:  // 127) AIBK Arab Investment Bank>> Prepaid Customer PDF 1/m
                case 311:  // 127) AIBK Arab Investment Bank>> Prepaid Staff PDF 1/m
                case 470:  // 153) BRKA AL Baraka Bank of Egypt >> Credit PDF 30/m

                    //clsMaintainData maintainDatapdf = new clsMaintainData();
                    //maintainDatapdf.fixAddress(bankCode);
                    //maintainDatapdf = null;
                    //if (pCmbProducts == 40)
                    //    {
                    //    clsMaintainData mntnData = new clsMaintainData();
                    //    mntnData.fixDecimalData(bankCode);
                    //    mntnData = null;
                    //    }
                    clsStatement_ExportRpt exportRpt = new clsStatement_ExportRpt();
                    exportRpt.setFrm = this;
                    exportRpt.mantainBank(bankCode);
                    if (pCmbProducts == 31 || pCmbProducts == 76 || pCmbProducts == 0 || pCmbProducts == 5 || pCmbProducts == 28 || pCmbProducts == 29 || pCmbProducts == 30 || pCmbProducts == 32 || pCmbProducts == 107 || pCmbProducts == 135 || pCmbProducts == 161 || pCmbProducts == 201)
                        exportRpt.StatLable = bankName + " Statement";
                    if (pCmbProducts == 135)
                    {
                        exportRpt.StatLable = "كشف حساب بطاقة فيزا الائتمانية";
                        //exportRpt.productCond = "('Visa Classic','Visa Gold','Visa Platinum','Visa Classic with 1.5% Interest','Visa Gold with 1.5% Interest','Visa Platinum with 1.5 interest')";
                        exportRpt.productCond = "'Visa%'";
                    }
                    if (pCmbProducts == 178)
                    {
                        exportRpt.StatLable = "كشف حساب بطاقة ماستر كارد الائتمانية";
                        //exportRpt.productCond = "('MasterCard Classic','MasterCard Platinum','MasterCard Titanium','MasterCard Gold','MasterCard Classic with 1.5 interest','MasterCard Platinum with 1.5 interest','MasterCard Titanium with 1.5 interest','MasterCard Gold with 1.5 interest')";
                        exportRpt.productCond = "'MasterCard%'";
                    }
                    if (pCmbProducts == 114)
                    {
                        exportRpt.closeBalance = "ALL";
                        exportRpt.productCond = "('MasterCard Prepaid Salary Card')";
                    }

                    // Debit products
                    if (pCmbProducts == 22 || pCmbProducts == 91 || pCmbProducts == 25 || pCmbProducts == 76 || pCmbProducts == 32 || pCmbProducts == 113 || pCmbProducts == 130 || pCmbProducts == 193 || pCmbProducts == 195 || pCmbProducts == 200 || pCmbProducts == 207 || pCmbProducts == 191 || pCmbProducts == 226 || pCmbProducts == 256)
                        exportRpt.closeBalance = "ALL";
                    if (pCmbProducts == 113)
                        exportRpt.PrepaidCondition = "('Visa Classic Prepaid','Visa Classic Prepaid Gift')";
                    if (pCmbProducts == 226)
                        exportRpt.PrepaidCondition = "('Electron Travel Prepaid')";
                    if (pCmbProducts == 256)
                        exportRpt.PrepaidCondition = "('MasterCard Prepaid General Spend Card')";
                    if (pCmbProducts == 202)
                        exportRpt.PrepaidCondition = "('Visa Classic Prepaid')";
                    if (pCmbProducts == 7100)
                        exportRpt.PrepaidCondition = "('MasterCard Corporate')";
                    exportRpt.export(txtFileName.Text, strStatementType, bankCode, strFileName, reportFleName, ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                    exportRpt.CreateZip();
                    exportRpt = null;
                    break;

                //Default Text Credit Statement
                case 38:   // Common Credit English Text - No Charge
                case 39:   // Common Reward English Text - No Charge
                case 297:  // 5) AIB Arab International Bank >> Corporate Text Spool 1/m
                case 217:  // 6) AAIB Arab African International Bank >> Text Spool 1/m
                //case 140:  //7) BAI Corporate >> Text 1/m
                case 185: // 18) BOAB Bank of Africa Benin >> Default Credit Text 1/m
                case 165:  //20) BICV Banco Intertlantico >> Credit Classic >> Text 1/m
                case 66:   // 21) ZEN Zenith Bank PLC >> Default Text for Email 16/m
                //case 52:   // 22) BPC BANCO DE POUPANCA E CREDITO, SARL Corporate >> Default Text 1/m
                //case 249: // 23) Suez Canal Bank(SCB) >> Text 1 / m
                case 300: //27) BCA BANCO COMERCIAL DO ATLANTICO >> Credit >> Default Text 1/m ==>iatta
                case 341: //27) BCA BANCO COMERCIAL DO ATLANTICO >> Corporate Text 30/m", 341));//341
                case 163:   // 28) CECV CAIXA ECONOMICA DE CABO VERDE >> Credit Classic >> Text 1/m
                //case 8:    // 29) UBA United Bank for Africa Plc Nigeria Default Text
                //case 42:   // 33) ZENG Zenith Bank (Ghana) Limited >> Default Text 16/m
                case 57:   // 37) PHB PLATINUM HABIB BANK PLC >> Default Text 1/m
                //case 161:  //38) GTBN GUARANTY TRUST BANK PLC NIGERIA  >> Credit Default Text 1/m
                case 326: //38) GTBN GUARANTY TRUST BANK PLC NIGERIA  >> Corprate Default Text 16/m
                case 53:   // 39) MBL Mainstreet Bank Limited >> Default Text 1/m
                case 44:   // 40) ICB INTERCONTINENTAL BANK PLC >> Default Text
                case 71:   // 40) ICB INTERCONTINENTAL BANK PLC >> Corporate Default Text 1/m
                case 206: // 42) ENG EchoBank Nigeria >> Default Credit Text 1/m
                case 222: // 47) SBN Sterling Bank Nigeria >> Default Credit Text 1/m
                case 315: // 47) SBN Sterling Bank Nigeria >> Default Credit Text 15/m
                case 105:  // 51] TMB Trust Merchant Bank >> Corporate Text 20/m
                case 95:   // 55) FBP Fidelity Bank PLC >> Default Text 16/m
                case 218:   // 55) FBP Fidelity Bank PLC >> Default Text 1/m
                case 92:   // 58) SBP SKYE BANK PLC >> Default Text 1/m
                case 125:  // 58) SBP SKYE BANK PLC >> Default Corporate Text 1/m
                case 79:   // 58) SBP SKYE BANK PLC >> Default Text 16/m
                case 109:  //69) First Bank of Nigeria  >> Credit Default Text 1/m
                case 116:  //69) First Bank of Nigeria  >> Corporate Default Text 1/m
                case 123:  //69) First Bank of Nigeria  >> Credit Default Text 16/m
                case 142:  //69) FBN First Bank of Nigeria  >> MasterCard Credit Standard Default Text 1/m
                case 131:  //72) ABPR ACCESS BANK RWANDA SA >> Default Credit Text 1/m
                //case 133:  //73) FCMB First City Monument Bank Nigeria >> Default Credit Text 20/m
                //case 146: // 76) EDBE EXPORT DEVELOPMENT BANK OF EGYPT >> Default Credit Text 1/m
                case 154:  // 13) BDK Bank De Kigali >> Default Credit Text 1/m
                case 328:  // 13) BDK Bank De Kigali >> Default Credit VIP Text 5/m
                case 329:  // 13) BDK Bank De Kigali >> Default Corporate Text 1/m
                case 330:  // 13) BDK Bank De Kigali >> Default Corporate VIP Text 5/m
                case 156: //81) SIBN STANBIC IBTC BANK NIGERIA >> Credit Default Text 1/m
                case 158: //81) SIBN STANBIC IBTC BANK NIGERIA  >> Corporate Default Text 1/m
                case 179: //87) WEMA BANK PLC NIGERIA >> Credit Default 15/m
                case 231: //102) VCBK Victoria Commercial Bank Ltd. Kenya  >> Credit Text 1/m
                case 234: //106) DSBJ EAST AFRICA BANK Djibouti  >> Credit Islamic Text 1/m
                case 238: //109) CGBK COMPAGNIE GENERALE DE BANQUE LTD >> MasterCard Credit Text 1/m
                case 240: //110) HBLN Heritage Banking Company Nigeria >> Credit Text 1/m
                case 250: //112) GTBR Guaranty trust bank Rwanda >> Credit Text 1/m
                case 252: //112) GTBR Guaranty trust bank Rwanda >> Corporate Text 1/m
                case 258: //114) IDBE Industrial Development & Workers Bank of Egypt >> Credit Text 1/m
                //case 293: //127) AIBK Arab Investment Bank of Egypt  >> Credit Text 30/m
                case 299: //122) ALXB ALEXBANK >> Credit Text 30/m
                case 343: //122) ALXB ALEXBANK >> Credit MF Text 30/m
                case 305: //122) ALXB ALEXBANK >> Corporate Text 30/m
                case 319: //136) BPG Bank Prestigo >> Credit Test 1/m
                case 400: //11) GTB Guaranty Trust Bank Ghana >> Credit Text 1/m
                case 437: //130) UNB Union National Bank >> Credit Text 10th/m
                case 440: //23) Suez Canal Bank(SCB) >> Corporate Text 1/m
                case 441: //154) UBP Unity Bank PlC >> Credit Text 15/m
                case 459: //158) RBGH Republic Bank(Ghana) Limited >> Default Credit Text 23rd / m
                case 502: //[146] Fidelity Bank Ghana Limited  >> PrePaid text 1 / m
                case 1999: //[146] Fidelity Bank Ghana Limited  >> PrePaid text 1 / m
                    //clsMaintainData maintainDatacr = new clsMaintainData();
                    //maintainDatacr.fixAddress(bankCode);
                    //maintainDatacr = null;
                    if (pCmbProducts == 53)
                    {
                        clsMaintainData maintainDataCrTxt = new clsMaintainData();
                        //maintainDataCrTxt.makeMainCardNum(bankCode);
                        maintainDataCrTxt.makeBranchAsMainCard(bankCode);
                        maintainDataCrTxt = null;
                    }
                    clsStatTxtLbl statTxtLbl = new clsStatTxtLbl();// + "ABP_Statement_File.txt"
                    statTxtLbl.setFrm = this;
                    //Participate Email Service
                    if (pCmbProducts == 57 || pCmbProducts == 66 || pCmbProducts == 79 || pCmbProducts == 95 || pCmbProducts == 109 || pCmbProducts == 123 || pCmbProducts == 154 || pCmbProducts == 156 || pCmbProducts == 161 || pCmbProducts == 179 || pCmbProducts == 218 || pCmbProducts == 222 || pCmbProducts == 240 || pCmbProducts == 315 || pCmbProducts == 328)
                    {
                        statTxtLbl.emailService = true;
                    }
                    //alx Bank >> condition here is contract type
                    if (pCmbProducts == 299) //122) ALXB ALEXBANK  >> Credit Text 30/m
                    {
                        statTxtLbl.emailService = true;
                        statTxtLbl.isFiltered = true;
                        statTxtLbl.productCond = "('MC World Credit - UnSecured','MC Platinum Credit - UnSecured','MC Titanium Credit - UnSecured','MC Gold Credit - UnSecured','VISA Classic Credit - Secured', 'VISA Classic Credit - Unsecured', 'VISA Classic Credit - No interest', 'VISA Classic Credit - Staff', 'VISA Gold Credit - Secured', 'VISA Gold Credit - Unsecured', 'VISA Gold Credit - No interest', 'VISA Gold Credit - Staff', 'MC Gold Credit - Secured', 'MC Gold Credit – Un Secured', 'MC Gold Credit - No Interest', 'MC Gold Credit - Staff', 'MC Platinum Credit - Secured', 'MC Platinum Credit – Unsecured', 'MC Platinum Credit - No interest', 'MC Platinum Credit - Staff', 'MC Titanium Credit - Secured', 'MC Titanium Credit – Unsecured', 'MC Titanium Credit - No interest', 'MC Titanium Credit - Staff', 'MC World Credit - Secured', 'MC World Credit – Unsecured', 'MC World Credit - No interest', 'MC World Credit - Staff')";

                    }
                    if (pCmbProducts == 343) //"122) ALXB ALEXBANK >> Credit MF Text 30/m
                    {
                        statTxtLbl.emailService = true;
                        statTxtLbl.isFiltered = true;
                        statTxtLbl.productCond = "('MC MF Individual Credit','MC MF Corporate Credit Cardholder')";
                    }
                    //CreateCorporate
                    if (pCmbProducts == 105 || pCmbProducts == 252 || pCmbProducts == 297 || pCmbProducts == 305 || pCmbProducts == 326 || pCmbProducts == 440)
                    {
                        statTxtLbl.CreateCorporate = true;
                    }
                    if (pCmbProducts == 116 || pCmbProducts == 125 || pCmbProducts == 158 || pCmbProducts == 329 || pCmbProducts == 330)
                    {
                        statTxtLbl.emailService = true;
                        statTxtLbl.CreateCorporate = true;
                    }
                    if (pCmbProducts == 53)
                        statTxtLbl.mainSortOrder = " USERCUSTFIELD2,m.accountno,m.cardprimary desc,m.cardno "; //m.cardproduct,m.CARDBRANCHPART,
                    if (pCmbProducts == 39 || pCmbProducts == 217)
                        statTxtLbl.isRewardVal = true;

                    if (pCmbProducts == 92)
                    {
                        statTxtLbl.isFiltered = true;
                        statTxtLbl.emailService = true;
                        statTxtLbl.productCond = "('Visa Gold','Visa Platinum')";
                    }

                    if (pCmbProducts == 222)
                    {
                        statTxtLbl.isFiltered = true;
                        //SBN EOM
                        //statTxtLbl.productCond = "('Visa Classic Credit Naira Contract','Visa Platinum Credit Naira Contract','VISA Gold Credit Naira Contract')";
                        statTxtLbl.productCond = "('Visa Classic Credit Dollar Contract')";
                    }

                    if (pCmbProducts == 315)
                    {
                        statTxtLbl.isFiltered = true;
                        //SBN 15th
                        //statTxtLbl.productCond = "('Visa Infinite Credit (USD)','Visa Signature Credit (NGN)')";
                        statTxtLbl.productCond = "('Visa Classic Credit Naira Contract', 'Visa Platinum Credit Naira Contract', 'Visa Platinum Credit Dollar Contract', 'VISA Gold Credit Naira Contract', 'Visa Infinite Credit (USD)', 'Visa Signature Credit (NGN)', 'Visa ULTRA Classic Credit Naira Contract', 'Visa ULTRA Platinum Credit Naira Contract', 'VISA ULTRA Gold Credit Naira Contract', 'Visa ULTRA Signature Credit (NGN)')";
                    }

                    if (pCmbProducts == 66)
                    {
                        statTxtLbl.isExcluded = true;
                        statTxtLbl.ExcludeCond = "('VISA PRIORITY PASS')";
                    }
                    if (pCmbProducts == 123)
                    {
                        statTxtLbl.isExcluded = true;
                        statTxtLbl.ExcludeCond = "('MasterCard Credit Standard')";
                    }
                    if (pCmbProducts == 234)
                    {
                        statTxtLbl.isFiltered = true;
                        statTxtLbl.productCond = "('MasterCard Credit Standard')";
                    }
                    if (pCmbProducts == 502)
                    {
                        statTxtLbl.productCond = "('Visa Classic Prepaid Card','Visa GoG Staff and Travel Prepaid Card','MC Business Prepaid','MC Personal Prepaid')";
                    }
                    //if (pCmbProducts == 1999)
                    //{
                    //    statTxtLbl.productCond = "('Visa Business Enhanced Credit')";
                    //    statTxtLbl.isFiltered = true;
                    //}
                    checkErrRslt = statTxtLbl.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLbl = null;
                    break;

                case 439: //94) FABG First Atlantic Bank Ghana >> Credit Text 1/m
                    clsStatTxtLbl_FABG statTxtLbl_FABG = new clsStatTxtLbl_FABG();// + "ABP_Statement_File.txt"
                    statTxtLbl_FABG.setFrm = this;
                    checkErrRslt = statTxtLbl_FABG.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);
                    statTxtLbl_FABG = null;
                    break;

                //Create Debit English Text
                case 67:   // Common Debit English Text - No Charge
                case 189: // 6) AAIB Arab African International Bank >> Prepaid Text 1/m
                //case 127: // 7) BAI Prepaid >> PDF 1/m
                case 339: //12) BSIC Alwaha Bank Libya Inc (Oasis Bank) >> Default Text 1/m
                case 100:  // 13) BDK Bank De Kigali >> Default Prepaid Text 1/m
                //case 74:   // 15) SSB SG-SSB Ltd Debit >> Text 1/m
                case 166:  //20) BICV Banco Intertlantico >> Debit >> Text 1/m
                //case 191: //22) BPC BANCO DE POUPANCA E CREDITO, SARL Debit >> Text 1/m
                case 164:   // 28) CECV CAIXA ECONOMICA DE CABO VERDE >> Debit - Prepaid >> Text 1/m
                case 68:   // 39) MBL Mainstreet Bank Limited >> Default Debit Text 1/m
                case 119: //42) OBI Oceanic Bank International >> Default Debit Text 1/m
                //case 120: //47) SBN Sterling Bank Nigeria >> Default Debit Text 1/m
                case 151:   //55) FBP Fidelity Bank PLC >> Default Text Debit 16/m
                case 90:   // 58) SBP SKYE BANK PLC >> Default Debit Text 16/m
                case 94:   // 58) SBP SKYE BANK PLC >> Default Debit Text 1/m
                //case 125: // 58) SBP SKYE BANK PLC >> Default Corporate Text 1/m
                case 121: //65) SBPG SKYE BANK PLC Gambia >> Default Debit Text 1/m
                case 129: // 66) ABPG Access Bank Plc Gambia >> Default Debit Text 1/m
                case 136: // 71) BOAD BANK OF AFRICA DEMOCRATIC REPUBLIC OF CONGO >> Default Debit Text 1/m
                //case 132: // 72) ABPR ACCESS BANK RWANDA SA >> Default Debit Text 1/m
                case 134: //73) FCMB First City Monument Bank Nigeria >> Default Debit Text 20/m
                //case 137: //75) UMB UNIVERSAL MERCHANT BANK GHANA >> Default Prepaid Text 1/m
                //case 182: //87) WEMA BANK PLC NIGERIA >> Debit/Prepaid Default 15/m
                case 229: //95) GTBL Guaranty trust bank Liberia  >> Debit Text 1/m
                case 232: //104) DSBS Dahabshil Bank International Somalia  >> Debit Text 1/m
                case 233: //106) DSBJ EAST AFRICA BANK Djibouti  >> Debit Text 1/m
                case 246: //110) HBLN Heritage Banking Company Nigeria >> Prepaid Text 1/m
                case 251: //112) GTBR Guaranty trust bank Rwanda >> Debit Text 1/m
                    //clsMaintainData maintainDatadb = new clsMaintainData();
                    //maintainDatadb.fixAddress(bankCode);
                    //maintainDatadb = null;
                    clsStatTxtLblDb statTxtLblDb = new clsStatTxtLblDb();
                    statTxtLblDb.setFrm = this;
                    //if (pCmbProducts == 115 || pCmbProducts == 125 || pCmbProducts == 127)
                    //if (pCmbProducts == 115 || pCmbProducts == 125)
                    //if (pCmbProducts == 182 || pCmbProducts == 189 || pCmbProducts == 246)
                    if (pCmbProducts == 100 || pCmbProducts == 189 || pCmbProducts == 246)
                        statTxtLblDb.emailService = true;
                    if (pCmbProducts == 100)
                        statTxtLblDb.PrepaidCondition = "('Visa Prepaid')";
                    if (pCmbProducts == 339)
                        statTxtLblDb.NoPrintAddress = true;
                    checkErrRslt = statTxtLblDb.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblDb = null;
                    break;

                case 115: //69) First Bank of Nigeria  >> Prepaid Default Text 1/m
                    //clsMaintainData maintainDatafbnpre = new clsMaintainData();
                    //maintainDatafbnpre.fixAddress(bankCode);
                    //maintainDatafbnpre = null;
                    clsStatTxtLblDbFBN statTxtLblDbfbn = new clsStatTxtLblDbFBN();
                    statTxtLblDbfbn.setFrm = this;
                    statTxtLblDbfbn.emailService = true;
                    checkErrRslt = statTxtLblDbfbn.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblDbfbn = null;
                    break;

                case 280: //114) IDBE Industrial Development & Workers Bank of Egypt >> Credit XML 1/m
                    //clsMaintainData maintainDatafbnpre = new clsMaintainData();
                    //maintainDatafbnpre.fixAddress(bankCode);
                    //maintainDatafbnpre = null;
                    clsStatXML_IDBE xmlidbe = new clsStatXML_IDBE();
                    xmlidbe.setFrm = this;
                    checkErrRslt = xmlidbe.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    xmlidbe = null;
                    break;

                case 182: //87) WEMA BANK PLC NIGERIA >> Prepaid Default 15/m
                    //clsMaintainData maintainDatafbnpre = new clsMaintainData();
                    //maintainDatafbnpre.fixAddress(bankCode);
                    //maintainDatafbnpre = null;
                    clsStatTxtLblDb_ICBG statTxtLblDbwema = new clsStatTxtLblDb_ICBG();
                    statTxtLblDbwema.setFrm = this;
                    statTxtLblDbwema.emailService = true;
                    statTxtLblDbwema.PrepaidCondition = "('Visa Prepaid Classic (USD)')";
                    checkErrRslt = statTxtLblDbwema.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblDbwema = null;
                    break;

                case 137: //75) UMB UNIVERSAL MERCHANT BANK GHANA >> Default Prepaid Text 1/m
                    //clsMaintainData maintainDatafbnpre = new clsMaintainData();
                    //maintainDatafbnpre.fixAddress(bankCode);
                    //maintainDatafbnpre = null;
                    clsStatTxtLblDb_ICBG statTxtLblDbumb = new clsStatTxtLblDb_ICBG();
                    statTxtLblDbumb.setFrm = this;
                    //statTxtLblDbumb.emailService = true;
                    statTxtLblDbumb.PrepaidCondition = "('Visa Prepaid Classic')";
                    checkErrRslt = statTxtLblDbumb.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblDbumb = null;
                    break;

                case 69:   // 27) BCA BANCO COMERCIAL DO ATLANTICO >> Debit >> Default Debit Text 1/m
                    //clsMaintainData maintainDatabcadb = new clsMaintainData();
                    //maintainDatabcadb.fixAddress(bankCode);
                    //maintainDatabcadb = null;
                    clsStatTxtLblDb_BCA statTxtLblDb_bca = new clsStatTxtLblDb_BCA();
                    statTxtLblDb_bca.setFrm = this;
                    checkErrRslt = statTxtLblDb_bca.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblDb_bca = null;
                    break;

                //case 228: //  98) GTBK Guaranty trust bank Kenya  >> Credit PDF 1/m
                //    clsStatement_ExportRpt stmntGTBK = new clsStatement_ExportRpt();
                //    stmntGTBK.setFrm = this;
                //    stmntGTBK.SplitByProduct(txtFileName.Text, strStatementType, bankCode, strFileName, reportFleName, ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                //    stmntGTBK.CreateZip();
                //    stmntGTBK = null;
                //    break;

                case 212:  //73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 7th/m
                case 213:  //73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 12th/m
                case 214:  //73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 17th/m
                case 133:  //73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 20th/m
                case 215:  //73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 23rd/m
                case 216:  //73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 27th/m
                case 267:  //73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 7th/m
                case 268:  //73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 12th/m
                case 269:  //73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 17th/m
                case 270:  //73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 20th/m
                case 271:  //73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 23rd/m
                case 272:  //73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 27th/m
                    clsStatTxtLbl_FCMB statTxtLbl_fcmb = new clsStatTxtLbl_FCMB();
                    statTxtLbl_fcmb.setFrm = this;
                    if (pCmbProducts == 212)
                    {
                        statTxtLbl_fcmb.isSplitted = true;
                        statTxtLbl_fcmb.PaymentSystem = "Visa";
                        statTxtLbl_fcmb.BillingCycle = "SD 7th";
                    }
                    if (pCmbProducts == 213)
                    {
                        statTxtLbl_fcmb.isSplitted = true;
                        statTxtLbl_fcmb.PaymentSystem = "Visa";
                        statTxtLbl_fcmb.BillingCycle = "SD 12th";
                    }
                    if (pCmbProducts == 214)
                    {
                        statTxtLbl_fcmb.isSplitted = true;
                        statTxtLbl_fcmb.PaymentSystem = "Visa";
                        statTxtLbl_fcmb.BillingCycle = "SD 17th";
                    }
                    if (pCmbProducts == 133)
                    {
                        statTxtLbl_fcmb.isSplitted = true;
                        statTxtLbl_fcmb.PaymentSystem = "Visa";
                        statTxtLbl_fcmb.BillingCycle = "SD 20th";
                    }
                    if (pCmbProducts == 215)
                    {
                        statTxtLbl_fcmb.isSplitted = true;
                        statTxtLbl_fcmb.PaymentSystem = "Visa";
                        statTxtLbl_fcmb.BillingCycle = "SD 23rd";
                    }
                    if (pCmbProducts == 216)
                    {
                        statTxtLbl_fcmb.isSplitted = true;
                        statTxtLbl_fcmb.PaymentSystem = "Visa";
                        statTxtLbl_fcmb.BillingCycle = "SD 27th";
                    }
                    if (pCmbProducts == 267)
                    {
                        statTxtLbl_fcmb.isSplitted = true;
                        statTxtLbl_fcmb.PaymentSystem = "MasterCard";
                        statTxtLbl_fcmb.BillingCycle = "SD 7th";
                    }
                    if (pCmbProducts == 268)
                    {
                        statTxtLbl_fcmb.isSplitted = true;
                        statTxtLbl_fcmb.PaymentSystem = "MasterCard";
                        statTxtLbl_fcmb.BillingCycle = "SD 12th";
                    }
                    if (pCmbProducts == 269)
                    {
                        statTxtLbl_fcmb.isSplitted = true;
                        statTxtLbl_fcmb.PaymentSystem = "MasterCard";
                        statTxtLbl_fcmb.BillingCycle = "SD 17th";
                    }
                    if (pCmbProducts == 270)
                    {
                        statTxtLbl_fcmb.isSplitted = true;
                        statTxtLbl_fcmb.PaymentSystem = "MasterCard";
                        statTxtLbl_fcmb.BillingCycle = "SD 20th";
                    }
                    if (pCmbProducts == 271)
                    {
                        statTxtLbl_fcmb.isSplitted = true;
                        statTxtLbl_fcmb.PaymentSystem = "MasterCard";
                        statTxtLbl_fcmb.BillingCycle = "SD 23rd";
                    }
                    if (pCmbProducts == 272)
                    {
                        statTxtLbl_fcmb.isSplitted = true;
                        statTxtLbl_fcmb.PaymentSystem = "MasterCard";
                        statTxtLbl_fcmb.BillingCycle = "SD 27th";
                    }
                    checkErrRslt = statTxtLbl_fcmb.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLbl_fcmb = null;
                    break;

                case 447:  //73) FCMB First City Monument Bank Nigeria >> Visa Credit Emails 7/m
                case 448:  //73) FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 7/m
                case 449:  //73) FCMB First City Monument Bank Nigeria >> Visa Credit Emails 12/m
                case 450:  //73) FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 12/m
                case 451:  //73) FCMB First City Monument Bank Nigeria >> Visa Credit Emails 17/m
                case 452:  //73) FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 17/m
                case 453:  //73) FCMB First City Monument Bank Nigeria >> Visa Credit Emails 20/m
                case 454:  //73) FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 20/m
                case 455:  //73) FCMB First City Monument Bank Nigeria >> Visa Credit Emails 23/m
                case 456:  //73) FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 23/m
                case 457:  //73) FCMB First City Monument Bank Nigeria >> Visa Credit Emails 27/m
                case 458:  //73) FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 27/m
                    clsStatHtmlFCMB statHtmlFCMB = new clsStatHtmlFCMB();
                    statHtmlFCMB.setFrm = this;
                    if (pCmbProducts == 447)
                    {
                        statHtmlFCMB.PaymentSystem = "Visa";
                        statHtmlFCMB.BillingCycle = "SD 7th";
                    }
                    if (pCmbProducts == 449)
                    {
                        statHtmlFCMB.PaymentSystem = "Visa";
                        statHtmlFCMB.BillingCycle = "SD 12th";
                    }
                    if (pCmbProducts == 451)
                    {
                        statHtmlFCMB.PaymentSystem = "Visa";
                        statHtmlFCMB.BillingCycle = "SD 17th";
                    }
                    if (pCmbProducts == 453)
                    {
                        statHtmlFCMB.PaymentSystem = "Visa";
                        statHtmlFCMB.BillingCycle = "SD 20th";
                    }
                    if (pCmbProducts == 455)
                    {
                        statHtmlFCMB.PaymentSystem = "Visa";
                        statHtmlFCMB.BillingCycle = "SD 23rd";
                    }
                    if (pCmbProducts == 457)
                    {
                        statHtmlFCMB.PaymentSystem = "Visa";
                        statHtmlFCMB.BillingCycle = "SD 27th";
                    }
                    if (pCmbProducts == 448)
                    {
                        statHtmlFCMB.PaymentSystem = "MasterCard";
                        statHtmlFCMB.BillingCycle = "SD 7th";
                    }
                    if (pCmbProducts == 450)
                    {
                        statHtmlFCMB.PaymentSystem = "MasterCard";
                        statHtmlFCMB.BillingCycle = "SD 12th";
                    }
                    if (pCmbProducts == 452)
                    {
                        statHtmlFCMB.PaymentSystem = "MasterCard";
                        statHtmlFCMB.BillingCycle = "SD 17th";
                    }
                    if (pCmbProducts == 454)
                    {
                        statHtmlFCMB.PaymentSystem = "MasterCard";
                        statHtmlFCMB.BillingCycle = "SD 20th";
                    }
                    if (pCmbProducts == 456)
                    {
                        statHtmlFCMB.PaymentSystem = "MasterCard";
                        statHtmlFCMB.BillingCycle = "SD 23rd";
                    }
                    if (pCmbProducts == 458)
                    {
                        statHtmlFCMB.PaymentSystem = "MasterCard";
                        statHtmlFCMB.BillingCycle = "SD 27th";
                    }
                    statHtmlFCMB.emailFromName = "FCMB";
                    statHtmlFCMB.emailFrom = "ebusiness@fcmb.com";
                    statHtmlFCMB.bankWebLink = "www.fcmb.com";
                    statHtmlFCMB.bankWebLinkService = "ebusiness@fcmb.com";
                    statHtmlFCMB.bankLogo = @"D:\pC#\ProjData\Statement\FCMB\FCMB-logo_for-DM.gif";
                    statHtmlFCMB.topPicture = @"D:\pC#\ProjData\Statement\FCMB\FCMB-Credit-Card-Statement-.jpg";
                    statHtmlFCMB.facebookLogo = @"D:\pC#\ProjData\Statement\FCMB\fb_nw_.png";
                    statHtmlFCMB.instagramLogo = @"D:\pC#\ProjData\Statement\FCMB\Ig_nw_.png";
                    statHtmlFCMB.twitterLogo = @"D:\pC#\ProjData\Statement\FCMB\twt_nw_.png";
                    statHtmlFCMB.linkinLogo = @"D:\pC#\ProjData\Statement\FCMB\LkdIn_nw_.png";
                    statHtmlFCMB.whatsappLogo = @"D:\pC#\ProjData\Statement\FCMB\WA_nw_.png";
                    statHtmlFCMB.waitPeriod = 500;
                    checkErrRslt = statHtmlFCMB.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlFCMB = null;
                    break;

                case 276: //115) UNBN UNION BANK NIGERIA  >> Credit Text 15/m
                case 277: //115) UNBN UNION BANK NIGERIA  >> Credit Text 20/m
                case 278: //115) UNBN UNION BANK NIGERIA  >> Credit Text 27/m
                case 279: //115) UNBN UNION BANK NIGERIA  >> Credit Text 30/m
                    clsStatTxtLbl statTxtLbl_unbn = new clsStatTxtLbl();
                    statTxtLbl_unbn.setFrm = this;
                    if (pCmbProducts == 276)
                    {
                        statTxtLbl_unbn.isSplitted = true;
                        //statTxtLbl_unbn.PaymentSystem = "Visa";
                        statTxtLbl_unbn.BillingCycle = "SD 15";
                    }
                    if (pCmbProducts == 277)
                    {
                        statTxtLbl_unbn.isSplitted = true;
                        //statTxtLbl_unbn.PaymentSystem = "Visa";
                        statTxtLbl_unbn.BillingCycle = "SD 20";
                    }
                    if (pCmbProducts == 278)
                    {
                        statTxtLbl_unbn.isSplitted = true;
                        //statTxtLbl_unbn.PaymentSystem = "Visa";
                        statTxtLbl_unbn.BillingCycle = "SD 27";
                    }
                    if (pCmbProducts == 279)
                    {
                        statTxtLbl_unbn.isSplitted = true;
                        //statTxtLbl_unbn.PaymentSystem = "Visa";
                        statTxtLbl_unbn.BillingCycle = "SD 30";
                    }
                    checkErrRslt = statTxtLbl_unbn.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLbl_unbn = null;
                    break;

                case 288: //[115] UNBN UNION BANK NIGERIA  >> Credit Emails 15/m
                case 289: //[115] UNBN UNION BANK NIGERIA  >> Credit Emails 20/m
                case 290: //[115] UNBN UNION BANK NIGERIA  >> Credit Emails 27/m
                case 291: //[115] UNBN UNION BANK NIGERIA  >> Credit Emails 30/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlUNBN statHtmlunbn = new clsStatHtmlUNBN();
                    if (pCmbProducts == 288)
                    {
                        statHtmlunbn.isSplitted = true;
                        //statHtmlunbn.PaymentSystem = "Visa";
                        statHtmlunbn.BillingCycle = "SD 15";
                    }
                    if (pCmbProducts == 289)
                    {
                        statHtmlunbn.isSplitted = true;
                        //statHtmlunbn.PaymentSystem = "Visa";
                        statHtmlunbn.BillingCycle = "SD 20";
                    }
                    if (pCmbProducts == 290)
                    {
                        statHtmlunbn.isSplitted = true;
                        //statHtmlunbn.PaymentSystem = "Visa";
                        statHtmlunbn.BillingCycle = "SD 27";
                    }
                    if (pCmbProducts == 291)
                    {
                        statHtmlunbn.isSplitted = true;
                        //statHtmlunbn.PaymentSystem = "Visa";
                        statHtmlunbn.BillingCycle = "SD 30";
                    }
                    statHtmlunbn.emailFromName = "Union Bank Credit Card Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlunbn.emailFrom = "cardservices@unionbankng.com";//mmohammed@emp-group.com
                    statHtmlunbn.bankWebLink = "www.unionbankng.com";
                    statHtmlunbn.bankLogo = @"D:\pC#\ProjData\Statement\UNBN\logo.gif";
                    statHtmlunbn.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    statHtmlunbn.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.gif";
                    //statHtmlunbn.bottomBanner = @"D:\pC#\ProjData\Statement\UNBN\bottom.jpg";
                    statHtmlunbn.waitPeriod = 3000;
                    statHtmlunbn.HasAttachement = true;
                    statHtmlunbn.isReward = true;
                    statHtmlunbn.rewardCond = "('Credit Card Rewards Program')";
                    statHtmlunbn.setFrm = this;
                    checkErrRslt = statHtmlunbn.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlunbn = null;
                    break;

                case 188:   // 5) AIB Debit >> Text 1/m
                    //clsMaintainData maintainDataaibdb = new clsMaintainData();
                    //maintainDataaibdb.fixAddress(bankCode);
                    //maintainDataaibdb = null;
                    clsStatTxtLblDb_AIB statTxtLblDb_AIB = new clsStatTxtLblDb_AIB();
                    statTxtLblDb_AIB.setFrm = this;
                    statTxtLblDb_AIB.PrepaidCondition = "('Electron (Debit)- EGP','Prepaid A/C (EGP)')";
                    checkErrRslt = statTxtLblDb_AIB.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblDb_AIB = null;
                    break;

                case 303: // 5) AIB Prepaid EGP >> Prepaid Raw 1/m
                    //clsMaintainData maintainDataaibdb = new clsMaintainData();
                    //maintainDataaibdb.fixAddress(bankCode);
                    //maintainDataaibdb = null;
                    clsStatementAibRaw_PRE AibRaw_PRE = new clsStatementAibRaw_PRE();
                    AibRaw_PRE.setFrm = this;
                    AibRaw_PRE.PrepaidCondition = "('Prepaid A/C (EGP)','Prepaid A/C (USD)')";
                    checkErrRslt = AibRaw_PRE.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    AibRaw_PRE = null;
                    break;
                case 1009: // 174) ABPK Prepaid >> Prepaid Raw 1/m
                    clsStatementAbpkRaw_PRE AbpkRaw_PRE2 = new clsStatementAbpkRaw_PRE();
                    AbpkRaw_PRE2.setFrm = this;
                    //  AibRaw_PRE2.PrepaidCondition = "('Prepaid A/C (EGP)','Prepaid A/C (USD)')";
                    AbpkRaw_PRE2.PrepaidCondition = "('Prepaid ','Prepaid A/C')";
                    checkErrRslt = AbpkRaw_PRE2.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    AbpkRaw_PRE2 = null;
                    break;

                // Create total supplementary -- clsStatTxtLblSupTot
                case 97:   // 36) DBN Diamond Bank Nigeria - VIP >> Reward Default Text 16/m
                    clsStatTxtLblSupTot statTxtLblSupTotvip = new clsStatTxtLblSupTot();//clsStatementDBN_Reward + "ABP_Statement_File.txt"
                    statTxtLblSupTotvip.setFrm = this;
                    statTxtLblSupTotvip.isRewardVal = true;
                    statTxtLblSupTotvip.rewardCondition = "('Reward Program')";
                    //statTxtLblSupTotvip.VIPCondition = "('Visa Gold VIP3 UnSecuredNoCashNoCRFreeDays','Visa Gold VIP4 UnSecuredNoCashNoInterestFees','VISA Platinum - EXCO VIP','Visa Platinum Private Banking Customer')";
                    statTxtLblSupTotvip.VIPCondition = "('Visa Gold VIP3 UnSecuredNoCashNoCRFreeDays','Visa Gold VIP4 UnSecuredCashNoInterestFees')";
                    //statTxtLblSupTotvip.VIPCondition = "('Visa Gold VIP3 UnSecuredNoCashNoCRFreeDays','Visa Gold VIP4 UnSecuredNoCashNoInterestFees')";
                    checkErrRslt = statTxtLblSupTotvip.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblSupTotvip = null;
                    break;
                case 83:   // 36) DBN Diamond Bank Nigeria - VIP 1,2,5 >> Reward Default Text 1/m
                    //clsMaintainData maintainDatadbnvip12 = new clsMaintainData();
                    //maintainDatadbnvip12.fixAddress(bankCode);
                    //maintainDatadbnvip12 = null;
                    clsStatTxtLblSupTot statTxtLblSupTotvip12 = new clsStatTxtLblSupTot();//clsStatementDBN_Reward + "ABP_Statement_File.txt"
                    statTxtLblSupTotvip12.setFrm = this;
                    statTxtLblSupTotvip12.isRewardVal = true;
                    statTxtLblSupTotvip12.rewardCondition = "('Reward Program')";
                    statTxtLblSupTotvip12.VIPCondition = "('Visa Gold VIP1 SecuredNoCashCRFreeDays','Visa Gold VIP2 UnSecuredNoCashCRFreeDays','Visa Gold VIP5 UnSecuredNoCashCRFreeDays')";
                    //clsMaintainData maintainDatavip12 = new clsMaintainData();
                    //maintainDatavip12.notRward = false;
                    //maintainDatavip12.curRewardCond = statTxtLblSupTotvip12.rewardCondition;
                    //maintainDatavip12.fixReward(bankCode, statTxtLblSupTotvip12.rewardCondition);
                    //maintainDatavip12.matchCardBranch4Account(bankCode);
                    //maintainDatavip12 = null;
                    checkErrRslt = statTxtLblSupTotvip12.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblSupTotvip12 = null;
                    break;
                case 46:   // 36) DBN Diamond Bank Nigeria >> Reward Default Text 16/m
                    //clsMaintainData maintainDatadbn = new clsMaintainData();
                    //maintainDatadbn.fixAddress(bankCode);
                    //maintainDatadbn = null;
                    clsStatTxtLblSupTot statTxtLblSupTot = new clsStatTxtLblSupTot();//clsStatementDBN_Reward + "ABP_Statement_File.txt"
                    statTxtLblSupTot.setFrm = this;
                    statTxtLblSupTot.isRewardVal = true;
                    statTxtLblSupTot.rewardCondition = "('Reward Program')";
                    statTxtLblSupTot.ProductCondition = "VISA SPECIAL EXCO VIP";
                    statTxtLblSupTot.VIPCondition = "('Visa Gold VIP3 UnSecuredNoCashNoCRFreeDays','Visa Gold VIP4 UnSecuredCashNoInterestFees','VISA Platinum - ParkNShop Co-Brand','MasterCard Platinum Staff','MasterCard Platinum Savings Xtra Card','MasterCard Gold Savings Xtra Card','MasterCard Gold Xclusive','MasterCard Platinum Xclusive','MasterCard Gold Staff','MasterCard Platinum Standard','MasterCard Gold Standard','MasterCard Platinum Credit','MasterCard Gold Credit','Visa Platinum CR - USD Main','Visa Platinum CR - USD Supp.')";
                    //clsMaintainData maintainDataall = new clsMaintainData();
                    //maintainDataall.notRward = false;
                    //maintainDataall.curRewardCond = statTxtLblSupTot.rewardCondition;
                    //maintainDataall.fixReward(bankCode, statTxtLblSupTot.rewardCondition);
                    //maintainDataall.matchCardBranch4Account(bankCode);
                    //maintainDataall = null;
                    checkErrRslt = statTxtLblSupTot.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblSupTot = null;
                    break;

                case 111:   // 36) DBN Diamond Bank Nigeria >> VISA Platinum - ParkNShop Co-Brand 16/m
                    clsStatTxtLblSupTot statTxtLblSupTotps = new clsStatTxtLblSupTot();//clsStatementDBN_Reward + "ABP_Statement_File.txt"
                    statTxtLblSupTotps.setFrm = this;
                    statTxtLblSupTotps.isRewardVal = true;
                    statTxtLblSupTotps.rewardCondition = "('ParkNShop Reward Program')";
                    statTxtLblSupTotps.VIPCondition = "('VISA Platinum - ParkNShop Co-Brand')";
                    checkErrRslt = statTxtLblSupTotps.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblSupTotps = null;
                    break;

                case 149:   // 36) 36) DBN Diamond Bank Nigeria >> VISA Platinum - EXCO_VIP-Brand 16/m Text
                    clsStatTxtLblSupTot statTxtLblSupTotexco = new clsStatTxtLblSupTot();//clsStatementDBN_Reward + "ABP_Statement_File.txt"
                    statTxtLblSupTotexco.setFrm = this;
                    statTxtLblSupTotexco.isRewardVal = true;
                    statTxtLblSupTotexco.rewardCondition = "('Reward Program')";
                    statTxtLblSupTotexco.ProductCondition = "VISA SPECIAL EXCO VIP";
                    statTxtLblSupTotexco.VIPCondition = "('Visa Platinum Private Banking Customer','VISA Platinum - EXCO VIP')";
                    checkErrRslt = statTxtLblSupTotexco.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblSupTotexco = null;
                    break;

                case 220:   // 36) 36) DBN Diamond Bank Nigeria >> MasterCard Credit 16/m Text
                    clsStatTxtLblSupTot statTxtLblSupTotmc = new clsStatTxtLblSupTot();//clsStatementDBN_Reward + "ABP_Statement_File.txt"
                    statTxtLblSupTotmc.setFrm = this;
                    statTxtLblSupTotmc.isRewardVal = true;
                    statTxtLblSupTotmc.rewardCondition = "('Reward Program')";
                    statTxtLblSupTotmc.VIPCondition = "('MasterCard Platinum Staff','MasterCard Platinum Savings Xtra Card','MasterCard Gold Savings Xtra Card','MasterCard Gold Xclusive','MasterCard Platinum Xclusive','MasterCard Gold Staff','MasterCard Platinum Standard','MasterCard Gold Standard','MasterCard Platinum Credit','MasterCard Gold Credit')";
                    checkErrRslt = statTxtLblSupTotmc.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblSupTotmc = null;
                    break;

                // Create Raw Data
                case 75:   // 51) TMB Trust Merchant Bank >> Raw VISA Text 20/m
                case 204:   // 51) TMB Trust Merchant Bank >> Raw MasterCard Text 20/m
                    //clsMaintainData maintainDatatmb = new clsMaintainData();
                    //maintainDatatmb.fixAddress(bankCode);
                    //maintainDatatmb = null;
                    clsStatRawData clsBasStatementRawDataTMB = new clsStatRawData();// + "AIB_Statement_Raw_File.txt"
                    if (pCmbProducts == 75)
                        clsBasStatementRawDataTMB.productCond = "('Classic Contract','Gold Contract','Platinum Contract','VISA Credit Gold-ET-Co-Branded','Visa Infinite Credit')";
                    if (pCmbProducts == 204)
                        clsBasStatementRawDataTMB.productCond = "('MasterCard Standard - Credit','MasterCard Gold - Credit','MasterCard Platinum - Credit','MasterCard World - Credit','MasterCard Standard Credit (EURO)')";
                    clsBasStatementRawDataTMB.bankName = "TMB";
                    clsBasStatementRawDataTMB.setFrm = this;
                    checkErrRslt = clsBasStatementRawDataTMB.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "AIB_Statement_Raw_File.txt"
                    clsBasStatementRawDataTMB = null;
                    break;

                // Create Raw Data
                case 205:   // 6) AAIB Arab African International Bank >> Raw Text 1/m
                    clsStatRawData_AAIB clsBasStatementRawDataAAIB = new clsStatRawData_AAIB();// + "AIB_Statement_Raw_File.txt"
                    //clsBasStatementRawDataAAIB.productCond = "('MasterCard Titanium - STANDARD - EGP','MasterCard Titanium - STANDARD - USD','MasterCard Titanium - STAFF - EGP','MasterCard Titanium - STAFF - USD','MasterCard Titanium - SECURED - EGP','MasterCard Titanium - SECURED - USD','Reward program','MasterCard Classic - STANDARD - EGP','MasterCard Classic - STANDARD - USD','MasterCard Classic - STAFF - EGP','MasterCard Classic - STAFF - USD','MasterCard Classic - SECURED - EGP','MasterCard Classic - SECURED - USD', 'Visa Signature - MAIN - SECURED CLIENT - EGP', 'Visa Signature - MAIN - UNSECURED CLIENT - EGP', 'Visa Signature - STAFF - MAIN SECURED - EGP', 'Visa Signature - STAFF - MAIN UNSECURED - EGP', 'Visa Signature - SUPP - SECURED CLIENT - EGP', 'Visa Signature - SUPP - UNSECURED CLIENT - EGP', 'Visa Signature - STAFF - SUPP SECURED - EGP', 'Visa Signature - STAFF - SUPP UNSECURED - EGP', 'Visa Signature - MAIN - SECURED CLIENT - USD', 'Visa Signature - MAIN - UNSECURED CLIENT - USD', 'Visa Signature - STAFF - MAIN SECURED - USD', 'Visa Signature - STAFF - MAIN UNSECURED - USD', 'Visa Signature - SUPP - SECURED CLIENT - USD', 'Visa Signature - SUPP - UNSECURED CLIENT - USD', 'Visa Signature - STAFF - SUPP SECURED - USD', 'Visa Signature - STAFF - SUPP UNSECURED - USD',  'MC World Elite - STANDARD - EGP', 'MC World Elite - STANDARD - USD', 'MC World Elite - STAFF - EGP', 'MC World Elite - STAFF - USD', 'MC World Elite - SECURED - EGP', 'MC World Elite - SECURED - USD')";
                    clsBasStatementRawDataAAIB.productCond = "";
                    // New Contract types    
                    //clsBasStatementRawDataAAIB.productCond = "('MC World Elite - STANDARD - EGP', 'MC World Elite - STANDARD - USD', 'MC World Elite - STAFF - EGP', 'MC World Elite - STAFF - USD', 'MC World Elite - SECURED - EGP', 'MC World Elite - SECURED - USD')";
                    //clsBasStatementRawDataAAIB.productCond = "('Visa Signature - MAIN - SECURED CLIENT - EGP', 'Visa Signature - MAIN - UNSECURED CLIENT - EGP', 'Visa Signature - STAFF - MAIN SECURED - EGP', 'Visa Signature - STAFF - MAIN UNSECURED - EGP', 'Visa Signature - SUPP - SECURED CLIENT - EGP', 'Visa Signature - SUPP - UNSECURED CLIENT - EGP', 'Visa Signature - STAFF - SUPP SECURED - EGP', 'Visa Signature - STAFF - SUPP UNSECURED - EGP', 'Visa Signature - MAIN - SECURED CLIENT - USD', 'Visa Signature - MAIN - UNSECURED CLIENT - USD', 'Visa Signature - STAFF - MAIN SECURED - USD', 'Visa Signature - STAFF - MAIN UNSECURED - USD', 'Visa Signature - SUPP - SECURED CLIENT - USD', 'Visa Signature - SUPP - UNSECURED CLIENT - USD', 'Visa Signature - STAFF - SUPP SECURED - USD', 'Visa Signature - STAFF - SUPP UNSECURED - USD')";
                    clsBasStatementRawDataAAIB.bankName = "AAIB";
                    clsBasStatementRawDataAAIB.setFrm = this;
                    checkErrRslt = clsBasStatementRawDataAAIB.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "AIB_Statement_Raw_File.txt"
                    clsBasStatementRawDataAAIB = null;
                    break;

                case 431:   // 130) UNB Union National Bank >> Raw Text 1/m
                    clsStatRawData_UNB clsBasStatementRawDataUNB = new clsStatRawData_UNB();
                    clsBasStatementRawDataUNB.productCond = "";//"('MasterCard Titanium - STANDARD - EGP','MasterCard Titanium - STANDARD - USD','MasterCard Titanium - STAFF - EGP','MasterCard Titanium - STAFF - USD','MasterCard Titanium - SECURED - EGP','MasterCard Titanium - SECURED - USD','Reward program','MasterCard Classic - STANDARD - EGP','MasterCard Classic - STANDARD - USD','MasterCard Classic - STAFF - EGP','MasterCard Classic - STAFF - USD','MasterCard Classic - SECURED - EGP','MasterCard Classic - SECURED - USD')";
                    clsBasStatementRawDataUNB.bankName = "UNB";
                    clsBasStatementRawDataUNB.setFrm = this;
                    checkErrRslt = clsBasStatementRawDataUNB.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "AIB_Statement_Raw_File.txt"
                    clsBasStatementRawDataUNB = null;
                    break;



                case 469:   // 153) BRKA AL Baraka Bank of Egypt >> Raw Text 30/m
                    clsStatRawData_BRKA clsBasStatementRawDataBRKA = new clsStatRawData_BRKA();
                    clsBasStatementRawDataBRKA.productCond = "";//"('MasterCard Titanium - STANDARD - EGP','MasterCard Titanium - STANDARD - USD','MasterCard Titanium - STAFF - EGP','MasterCard Titanium - STAFF - USD','MasterCard Titanium - SECURED - EGP','MasterCard Titanium - SECURED - USD','Reward program','MasterCard Classic - STANDARD - EGP','MasterCard Classic - STANDARD - USD','MasterCard Classic - STAFF - EGP','MasterCard Classic - STAFF - USD','MasterCard Classic - SECURED - EGP','MasterCard Classic - SECURED - USD')";
                    clsBasStatementRawDataBRKA.bankName = "BRKA";
                    clsBasStatementRawDataBRKA.isRewardVal = true;
                    clsBasStatementRawDataBRKA.rewardCondition = "('Reward Program')";
                    clsBasStatementRawDataBRKA.setFrm = this;
                    strFileName = "_Statement_File_";
                    checkErrRslt = clsBasStatementRawDataBRKA.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "AIB_Statement_Raw_File.txt"
                    clsBasStatementRawDataBRKA = null;
                    break;



                //// Create Raw Data
                case 412:   // 128) EGB Egyptian Gulf Bank of Egypt >> Credit Raw 30/m
                    //clsStatRawData_AIBK clsBasStatementRawDataAIBK = new clsStatRawData_AIBK();// + "AIB_Statement_Raw_File.txt"
                    //clsBasStatementRawDataAIBK.productCond = "('MasterCard Classic - Regular','MasterCard Gold - Regular','MasterCard Platinum - Regular','MasterCard Classic - Staff','MasterCard Gold - Staff','MasterCard Platinum - Staff')";
                    clsStatRawDataEGB clsBasStatementRawDataEGB = new clsStatRawDataEGB();   // + "AIB_Statement_Raw_File.txt"
                    clsBasStatementRawDataEGB.bankName = "EGB";
                    clsBasStatementRawDataEGB.setFrm = this;
                    checkErrRslt = clsBasStatementRawDataEGB.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "AIB_Statement_Raw_File.txt"
                    clsBasStatementRawDataEGB = null;
                    break;

                case 298:   // 127) ALXB ALEXBANK  >> Credit Raw 30/m
                case 344:   // 127) ALXB ALEXBANK  >> Credit MF Raw 30/m
                    clsStatRawData_ALXB clsBasStatementRawDataALXB = new clsStatRawData_ALXB();// + "AIB_Statement_Raw_File.txt"
                    //clsBasStatementRawDataALXB.productCond = "('VISA Classic Credit - Secured','VISA Classic Credit - Unsecured','VISA Classic Credit - No interest','VISA Classic Credit - Staff','VISA Gold Credit - Secured','VISA Gold Credit - Unsecured','VISA Gold Credit - No interest','VISA Gold Credit - Staff','VISA Gold Credit - 24% - Secured')";
                    //clsBasStatementRawDataALXB.productCond = "('VISA Classic Credit - Secured', 'VISA Classic Credit - Unsecured', 'VISA Classic Credit - No interest', 'VISA Classic Credit - Staff', 'VISA Gold Credit - Secured', 'VISA Gold Credit - Unsecured', 'VISA Gold Credit - No interest', 'VISA Gold Credit - Staff', 'MC Gold Credit - Secured', 'MC Gold Credit - UnSecured', 'MC Gold Credit - No Interest', 'MC Gold Credit - Staff', 'MC Platinum Credit - Secured', 'MC Platinum Credit – Unsecured', 'MC Platinum Credit - No interest', 'MC Platinum Credit - Staff', 'MC Titanium Credit - Secured', 'MC Titanium Credit – Unsecured', 'MC Titanium Credit - No interest', 'MC Titanium Credit - Staff', 'MC World Credit - Secured', 'MC World Credit – Unsecured', 'MC World Credit - No interest', 'MC World Credit - Staff')";
                    if (pCmbProducts == 344)
                    {
                        clsBasStatementRawDataALXB.productCond = "('MC MF Individual Credit','MC MF Corporate Credit Cardholder')";
                    }
                    else
                    {
                        clsBasStatementRawDataALXB.productCond = "('MC World Credit - UnSecured','MC Platinum Credit - UnSecured','MC Titanium Credit - UnSecured','VISA Classic Credit - Secured', 'VISA Classic Credit - Unsecured', 'VISA Classic Credit - No interest', 'VISA Classic Credit - Staff', 'VISA Gold Credit - Secured', 'VISA Gold Credit - Unsecured', 'VISA Gold Credit - No interest', 'VISA Gold Credit - Staff', 'MC Gold Credit - Secured', 'MC Gold Credit - UnSecured', 'MC Gold Credit - No Interest', 'MC Gold Credit - Staff', 'MC Platinum Credit - Secured', 'MC Platinum Credit – Unsecured', 'MC Platinum Credit - No interest', 'MC Platinum Credit - Staff', 'MC Titanium Credit - Secured', 'MC Titanium Credit – Unsecured', 'MC Titanium Credit - No interest', 'MC Titanium Credit - Staff', 'MC World Credit - Secured', 'MC World Credit – Unsecured', 'MC World Credit - No interest', 'MC World Credit - Staff')";
                    }
                    //clsBasStatementRawDataALXB.productCond = "";
                    clsBasStatementRawDataALXB.bankName = "ALXB";
                    clsBasStatementRawDataALXB.setFrm = this;
                    checkErrRslt = clsBasStatementRawDataALXB.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "AIB_Statement_Raw_File.txt"
                    clsBasStatementRawDataALXB = null;
                    break;

                case 304:   // 127) ALXB ALEXBANK  >> Corporate Raw 30/m
                    clsStatRawDataCorp_ALXB clsBasStatementRawDataCorpALXB = new clsStatRawDataCorp_ALXB();// + "AIB_Statement_Raw_File.txt"
                    clsBasStatementRawDataCorpALXB.productCond = "('VISA Corporate Cardholder', 'VISA Corporate', 'MC Corporate Executive', 'MC Corporate Executive Cardholder')";
                    clsBasStatementRawDataCorpALXB.bankName = "ALXB";
                    clsBasStatementRawDataCorpALXB.setFrm = this;
                    clsBasStatementRawDataCorpALXB.CreateCorporate = true;
                    checkErrRslt = clsBasStatementRawDataCorpALXB.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "AIB_Statement_Raw_File.txt"
                    clsBasStatementRawDataCorpALXB = null;
                    break;

                // Create Raw Data => VCBK
                case 235: //102) VCBK Victoria Commercial Bank Ltd. Kenya  >> Raw Text 1/m
                    clsStatRawData_VCBK clsBasStatementRawDataVCBK = new clsStatRawData_VCBK();
                    clsBasStatementRawDataVCBK.productCond = "('MasterCard Platinum Credit','MasterCard WMR Credit')";
                    clsBasStatementRawDataVCBK.bankName = "VCBK";
                    clsBasStatementRawDataVCBK.setFrm = this;
                    checkErrRslt = clsBasStatementRawDataVCBK.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);
                    clsBasStatementRawDataVCBK = null;
                    break;

                //Create Access MDB file
                case 17:   // Zenith Bank PLC >> ZEN >> 21
                    //clsMaintainData maintainDatazen = new clsMaintainData();
                    //maintainDatazen.fixAddress(bankCode);
                    //maintainDatazen = null;
                    clsStatMdb statMdb = new clsStatMdb(); //clsBasStatementExportMDB
                    statMdb.isExcluded = true;
                    statMdb.ExcludeCond = "('VISA PRIORITY PASS')";
                    checkErrRslt = statMdb.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);
                    statMdb.CreateZip();
                    statMdb = null;
                    break;

                //Create Access MDB file
                case 430:   // GTB >> 11
                    clsStatMdb statMdb_gtb = new clsStatMdb(); //clsBasStatementExportMDB
                    statMdb_gtb.isExcluded = false;
                    //statMdb_gtb.ExcludeCond = "('VISA PRIORITY PASS')";
                    checkErrRslt = statMdb_gtb.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);
                    statMdb_gtb.CreateZip();
                    statMdb_gtb = null;
                    break;

                case 72:   // Common Corporate Detail English Text - No Charge
                    //clsMaintainData maintainDatacorp = new clsMaintainData();
                    //maintainDatacorp.fixAddress(bankCode);
                    //maintainDatacorp = null;
                    clsStatement_CommonCompanyDtl stmntCorporateDtl = new clsStatement_CommonCompanyDtl();// + "ABP_Statement_File.txt"
                    stmntCorporateDtl.setFrm = this;
                    checkErrRslt = stmntCorporateDtl.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    stmntCorporateDtl = null;
                    break;

                case 1:   // NSGB credit Splitted by product
                    //clsMaintainData maintainDataqnb = new clsMaintainData();
                    //maintainDataqnb.fixAddress(bankCode);
                    //maintainDataqnb = null;
                    clsStatementNSGBcredit stmntNSGBcredit = new clsStatementNSGBcredit(); // + "NSGB_Credit_Statement.txt"
                    stmntNSGBcredit.setFrm = this;
                    stmntNSGBcredit.isRewardVal = true;
                    stmntNSGBcredit.rewardCondition = "('Reward Program - VISA Classic', 'Reward Program - Visa Gold', 'Reward Program - Visa Platinum' , 'Reward Program - Visa infinite','Reward Program - MasterCard Standard','Reward Program - MasterCard TITANIUM','Reward Program - MasterCard World Elite','Reward Program - Visa Signature')";
                    stmntNSGBcredit.ExcludeCond = "Visa Business - Individuals";
                    //stmntNSGBcredit.isInstallmentVal = true;
                    //stmntNSGBcredit.installmentCondition = "'Gold Installment Program'";
                    checkErrRslt = stmntNSGBcredit.SplitByProduct(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData); // + "NSGB_Credit_Statement.txt"
                    stmntNSGBcredit = null;
                    break;

                //case 19:   // clsStatementNSGB individual
                //clsMaintainData maintainDataqnbind = new clsMaintainData();
                //maintainDataqnbind.fixAddress(bankCode);
                //maintainDataqnbind = null;
                //clsStatementNSGBcredit stmntNSGBindividual = new clsStatementNSGBcredit(); // + "NSGB_Credit_Statement.txtSME
                //stmntNSGBindividual.setFrm = this;
                //checkErrRslt = stmntNSGBindividual.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData); // + "NSGB_Credit_Statement.txt"
                //stmntNSGBindividual = null;
                //break;

                case 2:   // NSGB business
                case 78:  // NSGB business SME
                case 148: // NSGB Mastercard Business - SME 
                case 432: // NSGB Corporate Contract - FEDCOC 
                case 433: // NSGB Cardholder Corporate - B2B 
                case 434: // NSGB Corporate Contract - B2B
                case 4000:
                case 4001:
                case 4002:
                case 4003:
                    //clsMaintainData maintainDataqnbbus = new clsMaintainData();
                    //maintainDataqnbbus.fixAddress(bankCode);
                    //maintainDataqnbbus = null;
                    clsStatementNSGBbusiness stmntNSGBbusiness = new clsStatementNSGBbusiness();
                    stmntNSGBbusiness.setFrm = this;
                    stmntNSGBbusiness.BankCode = 1;
                    if (pCmbProducts == 2)
                        stmntNSGBbusiness.productCond = "Corporate Contract - Single Limit in EGP";
                    else if (pCmbProducts == 78)
                        //stmntNSGBbusiness.productCond = "Corporate Contracte - SME";
                        stmntNSGBbusiness.productCond = "Corporate Contract - SME";
                    else if (pCmbProducts == 148)
                        stmntNSGBbusiness.productCond = "MC Corporate Contract - Business SME";
                    else if (pCmbProducts == 432)
                        stmntNSGBbusiness.productCond = "Corporate Contract - FEDCOC";
                    else if (pCmbProducts == 433)
                        stmntNSGBbusiness.productCond = "Cardholder Corporate - B2B";
                    else if (pCmbProducts == 434)
                        stmntNSGBbusiness.productCond = "Corporate Contract - B2B";
                    else if (pCmbProducts == 4000)
                        stmntNSGBbusiness.productCond = "Visa Platinum MVSE";
                    else if (pCmbProducts == 4001)
                        stmntNSGBbusiness.productCond = "Visa Platinum Single Limit in EGP - Corporate Contract";
                    else if (pCmbProducts == 4002)
                        stmntNSGBbusiness.productCond = "Visa Platinum SME";
                    else if (pCmbProducts == 4003)
                        stmntNSGBbusiness.productCond = "Visa Platinum B2B - Corporate Contract";
                    checkErrRslt = stmntNSGBbusiness.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);
                    stmntNSGBbusiness = null;
                    break;

                case 3:   // 3) NBK VISA Classic >> Text 1/m
                case 60:   // 3) NBK VISA Gold >> Text 1/m
                    //clsMaintainData maintainDataawb = new clsMaintainData();
                    //maintainDataawb.fixAddress(bankCode);
                    //maintainDataawb = null;
                    clsStatementAwbPlus stmntAwb = new clsStatementAwbPlus();// + "AWB_Plus_Statement_File.txt"
                    stmntAwb.setFrm = this;
                    if (pCmbProducts == 3)
                        stmntAwb.productCond = "Visa Classic";
                    else if (pCmbProducts == 60)
                        stmntAwb.productCond = "Visa Gold";
                    stmntAwb.rewardCondition = "('Reward Contract')";
                    checkErrRslt = stmntAwb.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "AWB_Plus_Statement_File.txt"
                    stmntAwb = null;
                    break;

                //case 54:   // 1) NSGB Credit >> Email 1/m
                //case 55:   // 1) NSGB Business >> Email 1/m
                //case 56:   // 1) NSGB Individual >> Email 1/m
                ////chkDontPrompt.Checked = true;
                //clsStatementNSGBcreditHtml emailStmntNSGB = new clsStatementNSGBcreditHtml();
                //emailStmntNSGB.emailFromName = "NSGB - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                //emailStmntNSGB.emailFrom = "Cardcenter.NSGB@socgen.com";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                //emailStmntNSGB.bankWebLink = "qnbalahli.com";//"www.emp-group.com""www.socgen.com"
                //emailStmntNSGB.bankWebLinkService = "cardcenter.nsgb@socgen.com";//"www.socgen.com"
                //emailStmntNSGB.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                //emailStmntNSGB.logoAlignment = "left";
                ////emailStmntNSGB.waitPeriod = 9000;
                //emailStmntNSGB.statMessageFile = @"D:\pC#\ProjData\Statement\NSGB\EmailMessage.txt";
                ////emailStmntNSGB.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\tawfeer.jpg";
                ////emailStmntNSGB.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\hotelclub.jpg";
                ////emailStmntNSGB.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\dubai.jpg";
                ////emailStmntNSGB.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\family2.jpg";
                ////emailStmntNSGB.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\Thailand.jpg";
                //emailStmntNSGB.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\back.jpg";
                ////emailStmntNSGB.bottomBannerDefault = @"D:\pC#\ProjData\Statement\NSGB\Tech.jpg";
                ////emailStmntNSGB.bottomBannerPlatinum = @"D:\pC#\ProjData\Statement\NSGB\Tech.jpg";
                ////emailStmntNSGB.bottomBannerDefault = @"D:\pC#\ProjData\Statement\NSGB\back.jpg";
                ////emailStmntNSGB.bottomBannerPlatinum = @"D:\pC#\ProjData\Statement\NSGB\back.jpg";
                //emailStmntNSGB.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\EmailMessageMonthly.txt";
                //emailStmntNSGB.IsSplitted = false;
                ////emailStmntNSGB.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\Flyers";
                //emailStmntNSGB.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.gif";

                //emailStmntNSGB.CreateCorporate = pCmbProducts == 55 ? true : false;
                ////emailStmntNSGB.waitPeriod = 500;
                //emailStmntNSGB.waitPeriod = 500;
                //emailStmntNSGB.setFrm = this;
                //checkErrRslt = emailStmntNSGB.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                //emailStmntNSGB = null;
                //break;

                case 4:   // 5) AIB Raw Data VISA USD>> Text 1/m
                case 64:   // 5) AIB Raw Data VISA EUR>> Text 1/m
                case 187:   // 5) AIB Raw Data VISAEGP>> Text 1/m
                case 237:   // 5) AIB Raw Data MasterCard USD>> Text 1/m
                case 296:   // 5) AIB Raw Data Corporate>> Text 1/m
                    //clsMaintainData maintainDataaib = new clsMaintainData();
                    //maintainDataaib.fixAddress(bankCode);
                    //maintainDataaib = null;
                    clsStatementAibRaw stmntAibRaw = new clsStatementAibRaw();// + "AIB_Statement_Raw_File.txt"
                    //clsStatement_Common stmntAibRaw = new clsStatement_Common();// + "AIB_Statement_Raw_File.txt"
                    stmntAibRaw.setFrm = this;
                    if (pCmbProducts == 4) //AIB-237 => EDT-1018
                    {
                        //stmntAibRaw.mainTblCond = " m.accountcurrency = 'USD' ";
                        //stmntAibRaw.supTblCond = " d.accountcurrency = 'USD' ";
                        stmntAibRaw.mainTblCond = " m.cardproduct like 'V%' and m.accountcurrency = 'USD' ";
                        stmntAibRaw.supTblCond = " d.accountcurrency = 'USD' and d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + bankCode + " and x.cardproduct like 'V%' and x.accountcurrency = 'USD')";
                    }
                    else if (pCmbProducts == 64)//AIB-237 => EDT-1018
                    {
                        //stmntAibRaw.mainTblCond = " m.accountcurrency = 'EUR' ";
                        //stmntAibRaw.supTblCond = " d.accountcurrency = 'EUR' ";
                        stmntAibRaw.mainTblCond = " m.cardproduct like 'V%' and m.accountcurrency = 'EUR' ";
                        stmntAibRaw.supTblCond = " d.accountcurrency = 'EUR' and d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + bankCode + " and x.cardproduct like 'V%' and x.accountcurrency = 'EUR')";
                    }
                    else if (pCmbProducts == 187)//AIB-237 => EDT-1018
                    {
                        //stmntAibRaw.mainTblCond = " m.accountcurrency = 'EGP' ";
                        //stmntAibRaw.supTblCond = " d.accountcurrency = 'EGP' ";
                        stmntAibRaw.mainTblCond = " m.cardproduct like 'V%' and m.accountcurrency = 'EGP' ";
                        stmntAibRaw.supTblCond = " d.accountcurrency = 'EGP' and d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + bankCode + " and x.cardproduct like 'V%' and x.accountcurrency = 'EGP')";
                    }
                    else if (pCmbProducts == 237) //AIB-237 => EDT-1018
                    {
                        //stmntAibRaw.mainTblCond = " m.accountcurrency = 'USD' ";
                        //stmntAibRaw.supTblCond = " d.accountcurrency = 'USD' ";
                        stmntAibRaw.mainTblCond = " m.cardproduct like 'M%' and m.accountcurrency = 'USD' ";
                        stmntAibRaw.supTblCond = " d.accountcurrency = 'USD' and d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + bankCode + " and x.cardproduct like 'M%' and x.accountcurrency = 'USD')";
                    }
                    else if (pCmbProducts == 296) //AIB-237 => EDT-1018
                    {
                        stmntAibRaw.CreateCorporate = true;
                        stmntAibRaw.mainTblCond = " m.cardproduct like '%Corp%'";
                        stmntAibRaw.supTblCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + bankCode + " and x.cardproduct like '%Corp%')";
                    }
                    checkErrRslt = stmntAibRaw.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "AIB_Statement_Raw_File.txt"
                    stmntAibRaw = null;
                    break;

                case 7:   // BAI
                case 464:   // 7) BAI Visa Classic >> PDF 1/m
                case 465:   // 7) BAI Visa Gold >> PDF 1/m
                case 466:   // 7) BAI Visa Platinum >> PDF 1/m
                    //clsMaintainData maintainDatabaicr = new clsMaintainData();
                    //maintainDatabaicr.fixAddress(bankCode);
                    //maintainDatabaicr = null;
                    clsStatement_ExportRpt stmntBAI = new clsStatement_ExportRpt();
                    stmntBAI.setFrm = this;
                    stmntBAI.mantainBank(bankCode);
                    if (pCmbProducts == 464)
                    {
                        stmntBAI.whereCond = " and cardproduct in ('Visa Classic')";
                        stmntBAI.productCond = "'Visa Classic'";
                    }
                    if (pCmbProducts == 465)
                    {
                        stmntBAI.whereCond = " and cardproduct in ('Visa Gold')";
                        stmntBAI.productCond = "'Visa Gold'";
                    }
                    if (pCmbProducts == 466)
                    {
                        stmntBAI.whereCond = " and cardproduct in ('Visa Platinum')";
                        stmntBAI.productCond = "'Visa Platinum'";
                    }
                    stmntBAI.SplitByBranch(txtFileName.Text, strStatementType, bankCode, strFileName, reportFleName, ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                    stmntBAI.CreateZip();
                    stmntBAI = null;
                    break;

                case 127:   // 7) BAI Prepaid >> PDF 1/m
                    clsStatement_ExportRpt stmntBAIDB = new clsStatement_ExportRpt();
                    stmntBAIDB.setFrm = this;
                    stmntBAIDB.mantainBank(bankCode);
                    stmntBAIDB.SplitByBranch(txtFileName.Text, strStatementType, bankCode, strFileName, reportFleName, ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                    stmntBAIDB.CreateZip();
                    stmntBAIDB = null;
                    break;

                case 74:   // 15) SSB SG-SSB Ltd Debit >> PDF 1/m
                    //clsMaintainData maintainDataSSB = new clsMaintainData();
                    //maintainDataSSB.makeBranchAsMainCard(bankCode);
                    //maintainDataSSB = null;
                    //clsMaintainData maintainDatassb = new clsMaintainData();
                    //maintainDatassb.fixAddress(bankCode);
                    //maintainDatassb = null;
                    clsStatement_ExportRpt stmntSSB = new clsStatement_ExportRpt();
                    stmntSSB.setFrm = this;
                    stmntSSB.SplitByBranch(txtFileName.Text, strStatementType, bankCode, strFileName, reportFleName, ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                    stmntSSB.CreateZip();
                    stmntSSB = null;
                    break;

                case 10:   // 16) ABP Access Bank Plc with Label Credit >> Text 1/m
                case 35:   // ABP Text Dual Currency
                case 61:   // 16) ABP Access Bank Plc with Label Infinite Dual Currency >> Text 1/m
                case 471:  // 16) ABP Access Bank Plc with Label AMEX Dual Currency >> Text 1/m
                case 81:   // 16) ABP Access Bank Plc with Label Credit >> Reward Default Text 1/m
                case 43:   // 16) ABP Access Bank Plc with Label Corporate >> Text 1/m
                    //clsMaintainData maintainDataabp = new clsMaintainData();
                    //maintainDataabp.fixAddress(bankCode);
                    //maintainDataabp = null;
                    clsStatementABPlbl stmntABPlbl = new clsStatementABPlbl();// + "ABP_Statement_File.txt"
                    stmntABPlbl.setFrm = this;
                    if (pCmbProducts == 81)
                    {
                        stmntABPlbl.isRewardVal = true;
                        stmntABPlbl.rewardCondition = "'Reward Program'";
                        stmntABPlbl.productCond = "('Access Visa Classic','Access Visa Gold','Access Visa Platinum','Access Visa Platinum Extra')";
                        stmntABPlbl.emailService = true;
                    }
                    if (pCmbProducts == 10)
                    {
                        stmntABPlbl.productCond = "('Access Visa Classic','Access Visa Gold','Access Visa Platinum','Access Visa Platinum Extra')";
                        stmntABPlbl.emailService = true;
                    }
                    if (pCmbProducts == 35)
                    {
                        stmntABPlbl.productCond = "('Access Visa Classic (Dual Currency)','Access Visa Gold (Dual Currency)','Access Visa Platinum (Dual Currency)','Access Visa Classic Exec (Dual Currency)','Access Visa Gold Exec(Dual Currency)','Access Classic E-Tour (Dual Currency)','Access Classic City of David (Dual Currency)','Access Gold E-Tour (Dual Currency)','Access Gold City of David (Dual Currency)','Access Platinum E-Tour (Dual Currency)','Access Platinum City of David (Dual Currency)','Access Visa Gold EXEC 2 (Dual Currency)')";
                        stmntABPlbl.emailService = true;
                    }
                    if (pCmbProducts == 61)
                    {
                        stmntABPlbl.productCond = "('Access Visa Infinite(Dual)')";
                        stmntABPlbl.emailService = true;
                    }
                    if (pCmbProducts == 471)
                    {
                        stmntABPlbl.productCond = "('AMEX GOLD (Dual Currency)','AMEX Platinum (Dual Currency)')";
                        stmntABPlbl.emailService = true;
                    }
                    if (pCmbProducts == 43)
                    {
                        stmntABPlbl.CreateCorporate = true;
                        stmntABPlbl.productCond = "('Business Cardholder Contract (DUAL CCY)','Business Corporate Contract (DUAL CCY)','Business Cardholder Contract Silver (DUAL CCY)','Business Corporate Contract Silver (DUAL CCY)','Access Visa Cardholder Exec Contract (DUAL CCY)','Access Visa Corporate Exec Contract (DUAL CCY)','Business Silver (DUAL CCY)-Level 1 (Additional)','Visa Credit Cardholder Classic (OANDO)','Visa Credit Corporate (OANDO)','Visa Credit Cardholder Platinum (OANDO)','Visa Credit Cardholder Gold (OANDO)','Visa Credit Corporate(OANDO)-Level 1 (Additional)','Business Corporate Contract')";
                    }
                    checkErrRslt = stmntABPlbl.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);
                    stmntABPlbl = null;
                    break;

                case 211:   //[16] ABP Access Bank Plc  >> Credit Emails 1/m
                    //clsMaintainData maintainabpestmt = new clsMaintainData();
                    //maintainabpestmt.fixAddress(bankCode);
                    //maintainabpestmt = null;
                    clsStatHtmlABP statHtmlABP = new clsStatHtmlABP();
                    statHtmlABP.emailFromName = "ABP Access Bank Plc - Statement";
                    statHtmlABP.emailFrom = "noreplyinfo@accessbankplc.com";
                    statHtmlABP.bankWebLink = "www.accessbankplc.com";
                    //statHtmlABP.bankLogo = @"D:\pC#\ProjData\Statement\ABP\BankLogo.jpg";
                    //statHtmlABP.BackGround = @"D:\pC#\ProjData\Statement\ABP\frmBackground.jpg";
                    //statHtmlABP.bottomBanner = @"D:\pC#\ProjData\Statement\ABP\footer1.jpg";
                    //statHtmlABP.bottomBanner1 = @"D:\pC#\ProjData\Statement\ABP\footer2.jpg";
                    statHtmlABP.waitPeriod = 3000;
                    statHtmlABP.HasAttachement = true;
                    statHtmlABP.setFrm = this;
                    checkErrRslt = statHtmlABP.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlABP = null;
                    break;

                case 230:   //[98] GTBK Guaranty trust bank Kenya  >> Credit Emails 15/m
                    //clsMaintainData maintaingtbkestmt = new clsMaintainData();
                    //maintaingtbkestmt.fixAddress(bankCode);
                    //maintaingtbkestmt = null;
                    clsStatHtmlGTBK statHtmlGTBK = new clsStatHtmlGTBK();
                    statHtmlGTBK.emailFromName = "GTBK Guaranty Trust Bank Limited - Statement";
                    statHtmlGTBK.emailFrom = "cardserviceske@gtbank.com";
                    statHtmlGTBK.bankWebLink = "www.gtbank.co.ke";
                    statHtmlGTBK.waitPeriod = 3000;
                    statHtmlGTBK.HasAttachement = true;
                    statHtmlGTBK.setFrm = this;
                    checkErrRslt = statHtmlGTBK.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlGTBK = null;
                    break;

                case 257:   //[98] GTBK Guaranty trust bank Kenya  >> Prepaid Emails 15/m
                    //clsMaintainData maintaingtbkestmt = new clsMaintainData();
                    //maintaingtbkestmt.fixAddress(bankCode);
                    //maintaingtbkestmt = null;
                    clsStatHtmlGTBKDB statHtmlGTBKPre = new clsStatHtmlGTBKDB();
                    statHtmlGTBKPre.emailFromName = "GTBK Guaranty Trust Bank Limited - Statement";
                    statHtmlGTBKPre.emailFrom = "cardserviceske@gtbank.com";
                    statHtmlGTBKPre.bankWebLink = "www.gtbank.co.ke";
                    statHtmlGTBKPre.IsSplitted = true;
                    statHtmlGTBKPre.productCond = "('MasterCard Prepaid General Spend Card')";
                    statHtmlGTBKPre.waitPeriod = 3000;
                    statHtmlGTBKPre.HasAttachement = true;
                    statHtmlGTBKPre.setFrm = this;
                    checkErrRslt = statHtmlGTBKPre.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlGTBKPre = null;
                    break;

                case 225:   //[16] ABP Access Bank Plc  >> Corporate Cardholder Emails 1/m
                    //clsMaintainData maintainabpestmt = new clsMaintainData();
                    //maintainabpestmt.fixAddress(bankCode);
                    //maintainabpestmt = null;
                    clsStatHtmlABP statHtmlABPCorp = new clsStatHtmlABP();
                    statHtmlABPCorp.emailFromName = "ABP Access Bank Plc - Statement";
                    statHtmlABPCorp.emailFrom = "noreplyinfo@accessbankplc.com";
                    statHtmlABPCorp.bankWebLink = "www.accessbankplc.com";
                    //statHtmlABPCorp.bankLogo = @"D:\pC#\ProjData\Statement\ABP\BankLogo.jpg";
                    //statHtmlABPCorp.BackGround = @"D:\pC#\ProjData\Statement\ABP\frmBackground.jpg";
                    //statHtmlABPCorp.bottomBanner = @"D:\pC#\ProjData\Statement\ABP\footer1.jpg";
                    //statHtmlABPCorp.bottomBanner1 = @"D:\pC#\ProjData\Statement\ABP\footer2.jpg";
                    statHtmlABPCorp.CreateCorporate = true;
                    statHtmlABPCorp.waitPeriod = 3000;
                    statHtmlABPCorp.HasAttachement = true;
                    statHtmlABPCorp.setFrm = this;
                    checkErrRslt = statHtmlABPCorp.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlABPCorp = null;
                    break;

                case 11:   // clsStatementAAIB
                    //clsMaintainData maintainDataaaib = new clsMaintainData();
                    //maintainDataaaib.fixAddress(bankCode);
                    //maintainDataaaib = null;
                    clsStatementAAIB stmntAAIB = new clsStatementAAIB();// + "AAIB_Statement_File.txt"
                    stmntAAIB.setFrm = this;
                    checkErrRslt = stmntAAIB.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "AAIB_Statement_File.txt"
                    stmntAAIB = null;
                    break;

                //case 12:   // clsStatementBMSRsave
                //    //clsMaintainData maintainDatabmsrs = new clsMaintainData();
                //    //maintainDatabmsrs.fixAddress(bankCode);
                //    //maintainDatabmsrs = null;
                //    //clsStatementBMSRsave stmntBMSR = new clsStatementBMSRsave(txtFileName.Text, bankName, bankCode, strFileName, stmntType ,appendData);
                //    clsStatementBMSRsave stmntBMSR = new clsStatementBMSRsave();
                //    stmntBMSR.setFrm = this;
                //    checkErrRslt = stmntBMSR.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);
                //    stmntBMSR = null;
                //    break;

                //case 13:   // 25 Ahli United Bank AUB >> PDF
                //    //clsMaintainData maintainDataaub = new clsMaintainData();
                //    //maintainDataaub.fixAddress(bankCode);
                //    //maintainDataaub = null;
                //    clsStatement_ExportRewardRpt stmntAUB = new clsStatement_ExportRewardRpt();
                //    stmntAUB.export(txtFileName.Text, strStatementType, bankCode, strFileName, Application.StartupPath + @"\Reports\Statement_Delta_AUB.rpt", ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                //    stmntAUB.exportContactData(txtFileName.Text, strStatementType, bankCode, strFileName, Application.StartupPath + @"\Reports\Statement_Delta_AUB_Contacts_Data.rpt", ExportFormatType.Excel, stmntDate, stmntType, appendData);
                //    stmntAUB.CreateZip();
                //    stmntAUB = null;
                //    break;

                case 15:   // 4) BMSR Bank Misr Moga >> Text 16/m
                case 16:   // 4) BMSR Bank Misr Classic >> Text 1/m
                case 73:   // 30) BMG Banque Misr Gulf >> Text 1/m
                case 88:   // 4) BMSR Bank Misr Credit Gold >> Text 1/m
                case 89:   // 4) BMSR Bank Misr Credit Mobinet >> Text 1/m
                    //clsStatementBMSRcredit stmntBMSRcredit = new clsStatementBMSRcredit();
                    //clsMaintainData maintainDatabmsr = new clsMaintainData();
                    //maintainDatabmsr.fixAddress(bankCode);
                    //maintainDatabmsr = null;
                    clsStatementBMSRcredit stmntBMSRcredit = new clsStatementBMSRcredit();
                    stmntBMSRcredit.setFrm = this;
                    stmntBMSRcredit.savedDataset = true;
                    stmntBMSRcredit.whereCond = whereCond;// " and contracttype = 'Visa Gold'";
                    if (pCmbProducts == 88)
                        stmntBMSRcredit.isVIP = true;
                    checkErrRslt = stmntBMSRcredit.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);
                    stmntBMSRcredit = null;
                    break;



                case 49:  // 41) BDCA Banque Du Caire Classic >> Text 5/m
                case 58:  // 41) BDCA Banque Du Caire Gold >> Text 5/m
                case 138: // 41) BDCA Banque Du Caire MasterCard Standard >> Text 5/m
                case 139: // 41) BDCA Banque Du Caire MasterCard Gold >> Text 5/m
                case 173: // 41) BDCA Banque Du Caire MasterCard Installment >> Text 5/m
                case 273: // 41) BDCA Banque Du Caire MasterCard Titanium >> Text 5/m
                case 318: // 41) BDCA Banque Du Caire MasterCard World Elite >> Text 5/m
                case 413: // 41) BDCA Banque Du Caire Classic staff >> Text 1/m
                case 513: //41) BDCA Banque Du Caire MasterCard Platinum >> Text 5 / m
                case 415: // 41) BDCA Banque Du Caire MasterCard Standard staff >> Text 1/m
                case 414: // 41) BDCA Banque Du Caire Gold staff >> Text 5/m
                case 416: // 41) BDCA Banque Du Caire MasterCard Gold staff >> Text 1/m
                case 417: // 41) BDCA Banque Du Caire MasterCard Installment staff >> Text 5/m
                case 418: // 41) BDCA Banque Du Caire MasterCard Titanium staff >> Text 5m
                case 419: // 41) BDCA Banque Du Caire MasterCard World Elite staff >> Text 5/m
                case 514: //41) BDCA Banque Du Caire MasterCard Platinum STAFF >> Text 5 / m
                case 515: //41) BDCA Banque Du Caire MasterCard Corporate >> Text 1 / m
                case 516: //41) BDCA Banque Du Caire MasterCard Corporate STAFF >> Text 1 / m
                    clsStatementBDCA stmntBDCA = new clsStatementBDCA();
                    stmntBDCA.setFrm = this;
                    stmntBDCA.savedDataset = true;
                    stmntBDCA.statMessageBox = @"D:\pC#\ProjData\Statement\BDCA\Message.txt";
                    if (pCmbProducts == 49)
                    {
                        stmntBDCA.productCond = "('Visa Classic')";
                        stmntBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 58)
                    {
                        stmntBDCA.productCond = "('Visa Gold')";
                        stmntBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 138)
                    {
                        stmntBDCA.productCond = "('MasterCard Standard','MasterCard Standard 1')";
                        stmntBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 139)
                    {
                        stmntBDCA.productCond = "('MasterCard Gold', 'MasterCard Gold 1' ,'MasterCard Gold old')";
                        stmntBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 173)
                    {
                        stmntBDCA.productCond = "('MasterCard Installment', 'MasterCard Installment 1' ,'MasterCard Installment old')";
                        stmntBDCA.InstallmentCondition = "('BDC Installment Card Program')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 273)
                    {
                        stmntBDCA.productCond = "('MasterCard Titanium', 'MasterCard Titanium 1' ,'MasterCard Co-Brand Titanium Credit')";
                        stmntBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 318)
                    {
                        stmntBDCA.productCond = "('MasterCard World Elite')";
                        stmntBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 513)
                    {
                        stmntBDCA.productCond = "('MC Platinum Credit Card')";
                        stmntBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 515)
                    {
                        stmntBDCA.productCond = "('MC Corporate Credit Card')";
                        stmntBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        stmntBDCA.isInstallmentVal = true;
                    }




                    if (pCmbProducts == 413)
                    {
                        stmntBDCA.productCond = "('Visa Classic - STAFF')";
                        stmntBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 414)
                    {
                        stmntBDCA.productCond = "('Visa Gold - STAFF')";
                        stmntBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 415)
                    {
                        stmntBDCA.productCond = "('MasterCard Standard - STAFF','MasterCard Standard 1 - STAFF')";
                        stmntBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 416)
                    {
                        stmntBDCA.productCond = "('MasterCard Gold - STAFF', 'MasterCard Gold 1 - STAFF' ,'MasterCard Gold old - STAFF')";
                        stmntBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 417)
                    {
                        stmntBDCA.productCond = "('MasterCard Installment - STAFF', 'MasterCard Installment 1 - STAFF' ,'MasterCard Installment old - STAFF')";
                        stmntBDCA.InstallmentCondition = "('BDC Installment Card Program')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 418)
                    {
                        stmntBDCA.productCond = "('MasterCard Titanium - STAFF', 'MasterCard Titanium 1 - STAFF' ,'MasterCard Co-Brand Titanium Credit - STAFF')";
                        stmntBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 419)
                    {
                        stmntBDCA.productCond = "('MasterCard World Elite - STAFF')";
                        stmntBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 514)
                    {
                        stmntBDCA.productCond = "('MC Platinum Credit Card - STAFF')";
                        stmntBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 516)
                    {
                        stmntBDCA.productCond = "('MC Corporate Credit Card - STAFF')";
                        stmntBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        stmntBDCA.isInstallmentVal = true;
                    }
                    checkErrRslt = stmntBDCA.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);
                    stmntBDCA = null;
                    break;

                //case 18:   // 15) SSB SG-SSB Ltd Corporate >> PDF 1/m
                //    clsStatement_CommonCorpExp stmntSSB = new clsStatement_CommonCorpExp();
                //    stmntSSB.reportCompany = Application.StartupPath + @"\Reports\Statement_Corp_Com_SSB.rpt";
                //    stmntSSB.reportIndividual = Application.StartupPath + @"\Reports\Statement_Corp_Ind_SSB.rpt";
                //    stmntSSB.export(txtFileName.Text, strStatementType, bankCode, strFileName, Application.StartupPath + @"\Reports\Statement_Corp_Com_SSB.rpt", ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                //    stmntSSB.CreateZip();
                //    stmntSSB = null;
                //    break;

                case 555:
                    var statHtmlBDCApdf = new clsStatHtmlBDCAPDF();
                    statHtmlBDCApdf.emailFromName = "BDC E-Statement";
                    statHtmlBDCApdf.emailFrom = "bdc.estatment@bdc.com.eg";
                    statHtmlBDCApdf.bankWebLink = "https://bit.ly/33peXqd";
                    statHtmlBDCApdf.bankLogo = @"D:\pC#\ProjData\Statement\BDCA\BDCALogo.png";
                    statHtmlBDCApdf.BackGround = @"D:\pC#\ProjData\Statement\BDCA\EmailBody.jpg";
                    //statHtmlBDCApdfCA.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\BDCA\EmailMessageBox.txt";
                    statHtmlBDCApdf.waitPeriod = 3000;
                    statHtmlBDCApdf.HasAttachement = true;
                    statHtmlBDCApdf.setFrm = this;
                    statHtmlBDCApdf.productCond = "('Visa Classic','Visa Classic - STAFF')";
                    checkErrRslt = statHtmlBDCApdf.Statement(txtFileName.Text,bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlBDCApdf = null;
                    break;


                case 26:    // 23) Suez Canal Bank(SCB) Visa Gold - Standard>> PDF 1/m
                case 27:    // 23) Suez Canal Bank(SCB) Visa Classic - Standard>> PDF 1/m
                case 196:   // 23) Suez Canal Bank(SCB) MC Titanium - Standard>> PDF 1/m
                case 197:   // 23) Suez Canal Bank(SCB) MC Classic - Standard>> PDF 1/m
                case 242:   // 23) Suez Canal Bank(SCB) Visa Gold - Staff>> PDF 1/m
                case 243:   // 23) Suez Canal Bank(SCB) Visa Classic - Staff>> PDF 1/m
                case 244:   // 23) Suez Canal Bank(SCB) MC Titanium - Staff>> PDF 1/m
                case 245:   // 23) Suez Canal Bank(SCB) MC Classic - Staff>> PDF 1/m
                case 426:   // 23) Suez Canal Bank(SCB) MC Platnium - Standard>> PDF 1/m
                case 428:   // 23) Suez Canal Bank(SCB) MC Platnium - Staff>> PDF 1/m
                    clsMaintainData maintainDataSuez = new clsMaintainData();
                    //maintainDataSuez.makeMainCardNum(bankCode);
                    //maintainDataSuez.fixAddress(bankCode);
                    maintainDataSuez.makeBranchAsMainCard(bankCode);
                    maintainDataSuez = null;

                    clsStatement_ExportRpt stmntSuez = new clsStatement_ExportRpt();
                    stmntSuez.setFrm = this;
                    stmntSuez.mantainBank(bankCode);
                    stmntSuez.whereCond = whereCond;
                    if (pCmbProducts == 26)
                        stmntSuez.productCond = "'Visa Gold - Standard'";
                    if (pCmbProducts == 27)
                        stmntSuez.productCond = "'Visa Classic - Standard'";
                    if (pCmbProducts == 196)
                        stmntSuez.productCond = "'MC Titanium - Standard'";
                    if (pCmbProducts == 197)
                        stmntSuez.productCond = "'MC Classic - Standard'";
                    if (pCmbProducts == 242)
                        stmntSuez.productCond = "'Visa Gold - Staff'";
                    if (pCmbProducts == 243)
                        stmntSuez.productCond = "'Visa Classic - Staff'";
                    if (pCmbProducts == 244)
                        stmntSuez.productCond = "'MC Titanium - Staff'";
                    if (pCmbProducts == 245)
                        stmntSuez.productCond = "'MC Classic - Staff'";
                    if (pCmbProducts == 426)
                        stmntSuez.productCond = "'MC Platnium - Standard'";
                    if (pCmbProducts == 428)
                        stmntSuez.productCond = "'MC Platnium - Staff'";
                    stmntSuez.SplitByBranchByProfile(txtFileName.Text, strStatementType, bankCode, strFileName, reportFleName, ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                    stmntSuez.SplitByDbCr(txtFileName.Text, strStatementType, bankCode, strFileName, reportFleName, ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                    stmntSuez.CreateZip();
                    stmntSuez = null;
                    break;

                case 249:   // 23) Suez Canal Bank(SCB) Visa Gold - Standard>> Text 1/m
                case 260:   // 23) Suez Canal Bank(SCB) Visa Classic - Standard>> Text 1/m
                case 261:   // 23) Suez Canal Bank(SCB) MC Titanium - Standard>> Text 1/m
                case 262:   // 23) Suez Canal Bank(SCB) MC Classic - Standard>> Text 1/m
                case 263:   // 23) Suez Canal Bank(SCB) Visa Gold - Staff>> Text 1/m
                case 264:   // 23) Suez Canal Bank(SCB) Visa Classic - Staff>> Text 1/m
                case 265:   // 23) Suez Canal Bank(SCB) MC Titanium - Staff>> Text 1/m
                case 266:   // 23) Suez Canal Bank(SCB) MC Classic - Staff>> Text 1/m
                case 427:   // 23) Suez Canal Bank(SCB) MC Platnium - Standard>> Text 1/m
                case 429:   // 23) Suez Canal Bank(SCB) MC Platnium - Staff>> Text 1/m
                    clsStatTxtLbl_Suez suezcrtxt = new clsStatTxtLbl_Suez();
                    suezcrtxt.setFrm = this;
                    if (pCmbProducts == 249)
                        suezcrtxt.productCond = "'Visa Gold - Standard'";
                    if (pCmbProducts == 260)
                        suezcrtxt.productCond = "'Visa Classic - Standard'";
                    if (pCmbProducts == 261)
                        suezcrtxt.productCond = "'MC Titanium - Standard'";
                    if (pCmbProducts == 262)
                        suezcrtxt.productCond = "'MC Classic - Standard'";
                    if (pCmbProducts == 263)
                        suezcrtxt.productCond = "'Visa Gold - Staff'";
                    if (pCmbProducts == 264)
                        suezcrtxt.productCond = "'Visa Classic - Staff'";
                    if (pCmbProducts == 265)
                        suezcrtxt.productCond = "'MC Titanium - Staff'";
                    if (pCmbProducts == 266)
                        suezcrtxt.productCond = "'MC Classic - Staff'";
                    if (pCmbProducts == 427)
                        suezcrtxt.productCond = "'MC Platnium - Standard'";
                    if (pCmbProducts == 429)
                        suezcrtxt.productCond = "'MC Platnium - Staff'";
                    checkErrRslt = suezcrtxt.SplitByBranch(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    suezcrtxt = null;
                    break;

                case 33:   // Common Corporate
                    //clsMaintainData maintainDatacomcorp = new clsMaintainData();
                    //maintainDatacomcorp.fixAddress(bankCode);
                    //maintainDatacomcorp = null;
                    clsStatement_CommonCorpExp stmntCorp = new clsStatement_CommonCorpExp();
                    stmntCorp.reportCompany = Application.StartupPath + @"\Reports\Statement_Common_Corporate_Company.rpt";
                    stmntCorp.reportIndividual = Application.StartupPath + @"\Reports\Statement_Common_Corporate_Individual.rpt";
                    stmntCorp.export(txtFileName.Text, strStatementType, bankCode, strFileName, Application.StartupPath + @"\Reports\Statement_Corp_Com_SSB.rpt", ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                    stmntCorp.CreateZip();
                    stmntCorp = null;
                    break;

                //case 34:   // Common English Reward >> PDF
                //    //clsMaintainData maintainDatarew = new clsMaintainData();
                //    //maintainDatarew.fixAddress(bankCode);
                //    //maintainDatarew = null;
                //    clsStatement_ExportRewardRpt stmntReward = new clsStatement_ExportRewardRpt();
                //    stmntReward.export(txtFileName.Text, strStatementType, bankCode, strFileName, Application.StartupPath + @"\Reports\Statement_Common_English_Reward.rpt", ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                //    stmntReward.CreateZip();
                //    stmntReward = null;
                //    break;

                case 6:   // XXX 10) BNP Credit >> Text 1/m
                    //clsStatementBNP stmnt_BNP = new clsStatementBNP();
                    //stmnt_BNP.setFrm = this;
                    //checkErrRslt = stmnt_BNP.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);
                    //stmnt_BNP = null;
                    break;

                //case 36:   // 10) BNP Corporate >> Text 1/m
                //case 37:   // 10) BNP Credit Reward >> Text 1/m
                //case 86:   // 10) BNP Debit >> Text 1/m
                //    //clsMaintainData maintainDatabnp = new clsMaintainData();
                //    //maintainDatabnp.fixAddress(bankCode);
                //    //maintainDatabnp = null;
                //    clsStatementBNP_Reward stmntBNP = new clsStatementBNP_Reward();
                //    stmntBNP.setFrm = this;
                //    if (pCmbProducts == 36)
                //        {
                //        stmntBNP.maxDetailInPage = 30;
                //        stmntBNP.CreateCorporate = true;
                //        }
                //    checkErrRslt = stmntBNP.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);
                //    stmntBNP = null;
                //    break;

                case 41:   // Zenith Bank PLC >> ZEN >> 21
                    //clsMaintainData maintainDataZENemail = new clsMaintainData();
                    //maintainDataZENemail.fixDecimalData(bankCode);
                    //maintainDataZENemail = null;

                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrl00 emailStmntZen = new clsStatHtmlGnrl00();
                    emailStmntZen.emailFromName = "Zenith Bank PLC - Visa Cards";
                    emailStmntZen.emailFrom = "cardstatements@zenithbank.com";
                    emailStmntZen.bankWebLink = "www.ZenithBank.com";
                    emailStmntZen.bankLogo = @"D:\pC#\ProjData\Statement\ZEN\logo.gif";
                    emailStmntZen.visaLogo = @"D:\pC#\ProjData\Statement\_Background\VisaLogo.gif";
                    emailStmntZen.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\ZEN\EmailMessageMonthly.txt";
                    //emailStmntZen.waitPeriod = 500;bai
                    emailStmntZen.waitPeriod = 500;
                    emailStmntZen.isExcluded = true;
                    emailStmntZen.ExcludeCond = "('VISA PRIORITY PASS')";
                    emailStmntZen.setFrm = this;
                    //emailStmntZen.waitPeriod = 10000;
                    checkErrRslt = emailStmntZen.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntZen = null;
                    break;

                case 241: //[110] HBLN Heritage Banking Company Nigeria >> MasterCard Credit Emails 1/m
                    //clsMaintainData maintainDatahblnemail = new clsMaintainData();
                    //maintainDatahblnemail.fixDecimalData(bankCode);
                    //maintainDatahblnemail = null;

                    //chkDontPrompt.Checked = true;
                    clsStatHtmlHBLN emailStmnthbln = new clsStatHtmlHBLN();
                    emailStmnthbln.emailFromName = "Heritage Banking Company - Statement";
                    emailStmnthbln.emailFrom = "contactCentre@hbng.com";
                    emailStmnthbln.bankWebLink = "www.hbng.com";
                    emailStmnthbln.bankLogo = @"D:\pC#\ProjData\Statement\HBLN\header.png";
                    emailStmnthbln.BottomBanner = @"D:\pC#\ProjData\Statement\HBLN\footer.png";
                    emailStmnthbln.waitPeriod = 500;
                    emailStmnthbln.setFrm = this;
                    checkErrRslt = emailStmnthbln.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmnthbln = null;
                    break;

                case 247: //[110] HBLN Heritage Banking Company Nigeria >> MasterCard Prepaid Emails 1/m
                    //clsMaintainData maintainDatahblnemail = new clsMaintainData();
                    //maintainDatahblnemail.fixDecimalData(bankCode);
                    //maintainDatahblnemail = null;

                    //chkDontPrompt.Checked = true;
                    clsStatHtmlHBLNPre emailStmnthblnp = new clsStatHtmlHBLNPre();
                    emailStmnthblnp.emailFromName = "Heritage Banking Company - Statement";
                    emailStmnthblnp.emailFrom = "contactCentre@hbng.com";
                    emailStmnthblnp.bankWebLink = "www.hbng.com";
                    emailStmnthblnp.bankLogo = @"D:\pC#\ProjData\Statement\HBLN\header.png";
                    emailStmnthblnp.BottomBanner = @"D:\pC#\ProjData\Statement\HBLN\footer.png";
                    emailStmnthblnp.waitPeriod = 500;
                    emailStmnthblnp.setFrm = this;
                    checkErrRslt = emailStmnthblnp.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmnthblnp = null;
                    break;

                case 45:   // 29) UBA United Bank for Africa Plc Nigeria >> Emails
                    clsStatHtmlUBA emailStmntUBA = new clsStatHtmlUBA();
                    //emailStmntUBA.emailFromName = "United Bank for Africa Plc Nigeria - Statement";
                    emailStmntUBA.emailFromName = "Cards Products";
                    //emailStmntUBA.emailFrom = "credit.cards@ubagroup.com";
                    emailStmntUBA.emailFrom = "creditcardstatement@ubagroup.com";
                    emailStmntUBA.bankWebLink = "www.ubagroup.com";
                    emailStmntUBA.bankLogo = @"D:\pC#\ProjData\Statement\UBA\logo.png";
                    //emailStmntUBAN.visaLogo = @"D:\pC#\ProjData\Statement\_Background\VisaLogo.gif";
                    emailStmntUBA.logoAlignment = "right";
                    emailStmntUBA.BottomBanner = @"D:\pC#\ProjData\Statement\UBA\Footer.png";
                    //emailStmntUBAN.waitPeriod = 500;
                    emailStmntUBA.waitPeriod = 500;
                    emailStmntUBA.setFrm = this;
                    //checkErrRslt = emailStmntUBA.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    checkErrRslt = emailStmntUBA.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntUBA = null;
                    break;

                case 435:   // 29) UBA United Bank for Africa Plc Nigeria >> Emails
                    clsStatHtmlUBA emailStmntUBA_15 = new clsStatHtmlUBA();
                    emailStmntUBA_15.emailFromName = "Cards Products";
                    emailStmntUBA_15.emailFrom = "creditcardstatement@ubagroup.com";
                    emailStmntUBA_15.bankWebLink = "www.ubagroup.com";
                    emailStmntUBA_15.bankLogo = @"D:\pC#\ProjData\Statement\UBA\logo.png";
                    emailStmntUBA_15.logoAlignment = "right";
                    emailStmntUBA_15.BottomBanner = @"D:\pC#\ProjData\Statement\UBA\Footer.png";
                    emailStmntUBA_15.waitPeriod = 500;
                    emailStmntUBA_15.setFrm = this;
                    //checkErrRslt = emailStmntUBA_15.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    checkErrRslt = emailStmntUBA_15.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntUBA = null;
                    break;


                case 84:   // 36] DBN Diamond Bank Nigeria - VIP 1,2,5 >> Email 1/m
                    //clsStatementDBN_Html_New emailStmntDBNvip12 = new clsStatementDBN_Html_New();
                    ////emailStmntDBNvip12.emailFromName = "Diamond Bank Nigeria - Statement";
                    ////DBN-2628
                    //emailStmntDBNvip12.emailFromName = "Diamond Bank PLC - Statement";
                    ////emailStmntDBNvip12.emailFrom = "diamondvisacard@diamondbank.com";
                    ////DBN-5403
                    //emailStmntDBNvip12.emailFrom = "cards@diamondbank.com";
                    //emailStmntDBNvip12.bankWebLink = "www.DiamondBank.com";
                    //emailStmntDBNvip12.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    ////emailStmntDBNvip12.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\logo.gif";
                    ////emailStmntDBNvip12.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\Header.jpg";
                    ////emailStmntDBNvip12.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\VISA_E_STATEMENT_header.jpg";
                    //emailStmntDBNvip12.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_header.jpg";
                    ////emailStmntDBNvip12.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\VisaCreditCard_MidBanner.jpg";
                    ////emailStmntDBNvip12.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\important_note.jpg";
                    //emailStmntDBNvip12.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_footer.jpg";
                    ////emailStmntDBNvip12.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\VisaCreditCard_BottomBanner.jpg";
                    ////emailStmntDBNvip12.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\VISA_E_STATEMENT_banner.jpg";
                    //emailStmntDBNvip12.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\Statement_banner.jpg";
                    //emailStmntDBNvip12.bankCard = @"D:\pC#\ProjData\Statement\DBN Diamond\Gold.jpg";
                    //emailStmntDBNvip12.RewardLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\Reward_Banner.jpg";
                    //emailStmntDBNvip12.logoAlignment = "left";
                    //emailStmntDBNvip12.isReward = true;
                    //emailStmntDBNvip12.rewardCond = "'Reward Program'";
                    //emailStmntDBNvip12.ProductCondition = "VISA Gold";
                    //emailStmntDBNvip12.VIPCondition = "('Visa Gold VIP1 SecuredNoCashCRFreeDays','Visa Gold VIP2 UnSecuredNoCashCRFreeDays','Visa Gold VIP5 UnSecuredNoCashCRFreeDays')";
                    //emailStmntDBNvip12.setFrm = this;
                    ////emailStmntDBNvip12.waitPeriod = 500;
                    //emailStmntDBNvip12.waitPeriod = 500;
                    //checkErrRslt = emailStmntDBNvip12.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    //emailStmntDBNvip12 = null;


                    var emailStmntDBNvip12 = new clsStatHtmlDBN();
                    emailStmntDBNvip12.emailFromName = "ABP Access Bank Plc - Statement";
                    emailStmntDBNvip12.emailFrom = "noreplyinfo@accessbankplc.com";
                    emailStmntDBNvip12.bankWebLink = "www.accessbankplc.com";
                    emailStmntDBNvip12.waitPeriod = 3000;
                    emailStmntDBNvip12.ProductCondition = "VISA Gold";
                    emailStmntDBNvip12.VIPCondition = "('Visa Gold VIP1 SecuredNoCashCRFreeDays','Visa Gold VIP2 UnSecuredNoCashCRFreeDays','Visa Gold VIP5 UnSecuredNoCashCRFreeDays')";
                    emailStmntDBNvip12.HasAttachement = true;
                    emailStmntDBNvip12.setFrm = this;
                    checkErrRslt = emailStmntDBNvip12.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntDBNvip12 = null;

                    break;

                case 102:   // [36] DBN Diamond Bank Nigeria - VIP 1,2,5 Supplementary>> Email 1/m
                    clsStatementDBNSup_Html emailStmntDBNSupvip12 = new clsStatementDBNSup_Html();
                    //emailStmntDBNSupvip12.emailFromName = "Diamond Bank Nigeria - Transaction Details";
                    //DBN-2628
                    emailStmntDBNSupvip12.emailFromName = "Diamond Bank PLC - Transaction Details";
                    //emailStmntDBNSupvip12.emailFrom = "diamondvisacard@diamondbank.com";
                    //DBN-5403
                    emailStmntDBNSupvip12.emailFrom = "cards@diamondbank.com";
                    emailStmntDBNSupvip12.bankWebLink = "www.DiamondBank.com";
                    emailStmntDBNSupvip12.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    //emailStmntDBNSupvip12.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\logo.gif";
                    //emailStmntDBNSupvip12.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\Header.jpg";
                    //emailStmntDBNSupvip12.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\VISA_E_STATEMENT_header.jpg";
                    emailStmntDBNSupvip12.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_header.jpg";
                    //emailStmntDBNSupvip12.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\VisaCreditCard_MidBanner.jpg";
                    //emailStmntDBNSupvip12.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\important_note.jpg";
                    emailStmntDBNSupvip12.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_footer.jpg";
                    //emailStmntDBNSupvip12.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\VisaCreditCard_BottomBanner.jpg";
                    //emailStmntDBNSupvip12.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\VISA_E_STATEMENT_banner.jpg";
                    emailStmntDBNSupvip12.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\Statement_banner.jpg";
                    emailStmntDBNSupvip12.logoAlignment = "left";
                    emailStmntDBNSupvip12.isReward = true;
                    emailStmntDBNSupvip12.rewardCond = "'Reward Program'";
                    emailStmntDBNSupvip12.VIPCondition = "('Visa Gold VIP1 SecuredNoCashCRFreeDays','Visa Gold VIP2 UnSecuredNoCashCRFreeDays','Visa Gold VIP5 UnSecuredNoCashCRFreeDays')";
                    emailStmntDBNSupvip12.setFrm = this;
                    //emailStmntDBNSupvip12.waitPeriod = 500;
                    emailStmntDBNSupvip12.waitPeriod = 500;
                    checkErrRslt = emailStmntDBNSupvip12.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntDBNSupvip12 = null;

                    //clsStatHtmlABP_Sup emailStmntDBNSupvip12 = new clsStatHtmlABP_Sup();
                    //emailStmntDBNSupvip12.emailFromName = "ABP Access Bank Plc - Statement";
                    //emailStmntDBNSupvip12.emailFrom = "noreplyinfo@accessbankplc.com";
                    //emailStmntDBNSupvip12.bankWebLink = "www.accessbankplc.com";
                    //emailStmntDBNSupvip12.waitPeriod = 3000;
                    ////emailStmntDBNSupvip12.ProductCondition = "('VISA CLASSIC','VISA Classic - One Credit')";
                    //emailStmntDBNSupvip12.VIPCondition = "('Visa Gold VIP1 SecuredNoCashCRFreeDays','Visa Gold VIP2 UnSecuredNoCashCRFreeDays','Visa Gold VIP5 UnSecuredNoCashCRFreeDays')";
                    //emailStmntDBNSupvip12.HasAttachement = true;
                    //emailStmntDBNSupvip12.setFrm = this;
                    //checkErrRslt = emailStmntDBNSupvip12.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    //emailStmntDBNSupvip12 = null;

                    break;

                case 48:   // 36) DBN Diamond Bank Nigeria >> Email
                    //chkDontPrompt.Checked = true;
                    clsStatementDBN_Html_New emailStmntDBN = new clsStatementDBN_Html_New();
                    //emailStmntDBN.emailFromName = "Diamond Bank Nigeria - Statement";
                    //DBN-2628
                    emailStmntDBN.emailFromName = "Diamond Bank PLC - Statement";
                    //emailStmntDBN.emailFrom = "diamondvisacard@diamondbank.com";
                    //DBN-5403
                    emailStmntDBN.emailFrom = "cards@diamondbank.com";
                    emailStmntDBN.bankWebLink = "www.DiamondBank.com";
                    emailStmntDBN.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    //emailStmntDBN.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\logo.gif";
                    //emailStmntDBN.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\Header.jpg";
                    //emailStmntDBN.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\VISA_E_STATEMENT_header.jpg";
                    emailStmntDBN.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_header.jpg";
                    //emailStmntDBN.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\VisaCreditCard_MidBanner.jpg";
                    //emailStmntDBN.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\important_note.jpg";
                    emailStmntDBN.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_footer.jpg";
                    //emailStmntDBN.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\VisaCreditCard_BottomBanner.jpg";
                    //emailStmntDBN.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\VISA_E_STATEMENT_banner.jpg";
                    emailStmntDBN.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\Statement_banner.jpg";
                    emailStmntDBN.logoAlignment = "left";
                    emailStmntDBN.isReward = true;
                    emailStmntDBN.rewardCond = "'Reward Program'";
                    emailStmntDBN.ProductCondition = "VISA SPECIAL EXCO VIP";
                    emailStmntDBN.VIPCondition = "('Visa Gold VIP3 UnSecuredNoCashNoCRFreeDays','Visa Gold VIP4 UnSecuredCashNoInterestFees','VISA Platinum - ParkNShop Co-Brand','MasterCard Platinum Staff','MasterCard Platinum Savings Xtra Card','MasterCard Gold Savings Xtra Card','MasterCard Gold Xclusive','MasterCard Platinum Xclusive','MasterCard Gold Staff','MasterCard Platinum Standard','MasterCard Gold Standard','MasterCard Platinum Credit','MasterCard Gold Credit','Visa Platinum CR - USD Main','Visa Platinum CR - USD Supp.')";
                    emailStmntDBN.setFrm = this;
                    //emailStmntDBN.waitPeriod = 500;
                    emailStmntDBN.waitPeriod = 500;
                    checkErrRslt = emailStmntDBN.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntDBN = null;
                    break;

                case 284:   // 36) DBN Diamond Bank Nigeria >> Email Classic 16/m
                    //clsStatementDBN_Html_New emailStmntDBNClassic = new clsStatementDBN_Html_New();
                    //emailStmntDBNClassic.emailFromName = "Diamond Bank PLC - Statement";
                    //emailStmntDBNClassic.emailFrom = "cards@diamondbank.com";
                    //emailStmntDBNClassic.bankWebLink = "www.DiamondBank.com";
                    //emailStmntDBNClassic.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    //emailStmntDBNClassic.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_header.jpg";
                    //emailStmntDBNClassic.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_footer.jpg";
                    //emailStmntDBNClassic.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\Statement_banner.jpg";
                    //emailStmntDBNClassic.bankCard = @"D:\pC#\ProjData\Statement\DBN Diamond\Classic.jpg";
                    //emailStmntDBNClassic.RewardLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\Reward_Banner.jpg";
                    //emailStmntDBNClassic.logoAlignment = "left";
                    //emailStmntDBNClassic.isReward = true;
                    //emailStmntDBNClassic.rewardCond = "'Reward Program'";
                    //emailStmntDBNClassic.ProductCondition = "('VISA CLASSIC','VISA Classic - One Credit')";
                    //emailStmntDBNClassic.VIPCondition = "('Visa Classic','Visa Classic Savings Xtra Card','Visa Classic-One Credit','Visa Classic-One Credit Savings Xtra Card')";
                    //emailStmntDBNClassic.setFrm = this;
                    //emailStmntDBNClassic.waitPeriod = 500;
                    //checkErrRslt = emailStmntDBNClassic.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    //emailStmntDBNClassic = null;


                    var emailStmntDBNClassic = new clsStatHtmlDBN();
                    emailStmntDBNClassic.emailFromName = "ABP Access Bank Plc - Statement";
                    emailStmntDBNClassic.emailFrom = "noreplyinfo@accessbankplc.com";
                    emailStmntDBNClassic.bankWebLink = "www.accessbankplc.com";
                    emailStmntDBNClassic.waitPeriod = 3000;
                    emailStmntDBNClassic.ProductCondition = "('VISA CLASSIC','VISA Classic - One Credit')";
                    emailStmntDBNClassic.VIPCondition = "('Visa Classic','Visa Classic Savings Xtra Card','Visa Classic-One Credit','Visa Classic-One Credit Savings Xtra Card')";
                    emailStmntDBNClassic.HasAttachement = true;
                    emailStmntDBNClassic.setFrm = this;
                    checkErrRslt = emailStmntDBNClassic.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntDBNClassic = null;

                    break;

                case 285:   // 36) DBN Diamond Bank Nigeria >> Email Gold 16/m
                    //clsStatementDBN_Html_New emailStmntDBNGold = new clsStatementDBN_Html_New();
                    //emailStmntDBNGold.emailFromName = "Diamond Bank PLC - Statement";
                    //emailStmntDBNGold.emailFrom = "cards@diamondbank.com";
                    //emailStmntDBNGold.bankWebLink = "www.DiamondBank.com";
                    //emailStmntDBNGold.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    //emailStmntDBNGold.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_header.jpg";
                    //emailStmntDBNGold.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_footer.jpg";
                    //emailStmntDBNGold.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\Statement_banner.jpg";
                    //emailStmntDBNGold.bankCard = @"D:\pC#\ProjData\Statement\DBN Diamond\Gold.jpg";
                    //emailStmntDBNGold.RewardLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\Reward_Banner.jpg";
                    //emailStmntDBNGold.logoAlignment = "left";
                    //emailStmntDBNGold.isReward = true;
                    //emailStmntDBNGold.rewardCond = "'Reward Program'";
                    //emailStmntDBNGold.ProductCondition = "VISA Gold";
                    //emailStmntDBNGold.VIPCondition = "('Visa Gold','Visa Gold Staff','Visa Gold Xclusive','Visa Gold Savings Xtra Card')";
                    //emailStmntDBNGold.setFrm = this;
                    //emailStmntDBNGold.waitPeriod = 500;
                    //checkErrRslt = emailStmntDBNGold.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    //emailStmntDBNGold = null;


                    var emailStmntDBNGold = new clsStatHtmlDBN();
                    emailStmntDBNGold.emailFromName = "ABP Access Bank Plc - Statement";
                    emailStmntDBNGold.emailFrom = "noreplyinfo@accessbankplc.com";
                    emailStmntDBNGold.bankWebLink = "www.accessbankplc.com";
                    emailStmntDBNGold.waitPeriod = 3000;
                    emailStmntDBNGold.ProductCondition = "VISA Gold";
                    emailStmntDBNGold.VIPCondition = "('Visa Gold','Visa Gold Staff','Visa Gold Xclusive','Visa Gold Savings Xtra Card')";
                    emailStmntDBNGold.HasAttachement = true;
                    emailStmntDBNGold.setFrm = this;
                    checkErrRslt = emailStmntDBNGold.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntDBNGold = null;

                    break;

                case 286:   // 36) DBN Diamond Bank Nigeria >> Email Platinum 16/m
                    //clsStatementDBN_Html_New emailStmntDBNPlatinum = new clsStatementDBN_Html_New();
                    //emailStmntDBNPlatinum.emailFromName = "Diamond Bank PLC - Statement";
                    //emailStmntDBNPlatinum.emailFrom = "cards@diamondbank.com";
                    //emailStmntDBNPlatinum.bankWebLink = "www.DiamondBank.com";
                    //emailStmntDBNPlatinum.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    //emailStmntDBNPlatinum.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_header.jpg";
                    //emailStmntDBNPlatinum.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_footer.jpg";
                    //emailStmntDBNPlatinum.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\Statement_banner.jpg";
                    //emailStmntDBNPlatinum.bankCard = @"D:\pC#\ProjData\Statement\DBN Diamond\Platinum.jpg";
                    //emailStmntDBNPlatinum.RewardLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\Reward_Banner.jpg";
                    //emailStmntDBNPlatinum.logoAlignment = "left";
                    //emailStmntDBNPlatinum.isReward = true;
                    //emailStmntDBNPlatinum.rewardCond = "'Reward Program'";
                    //emailStmntDBNPlatinum.ProductCondition = "('VISA Platinum')";
                    //emailStmntDBNPlatinum.VIPCondition = "('Visa Platinum','Visa Platinum Staff','Visa Platinum Xclusive','Visa Platinum Savings Xtra Card')";
                    //emailStmntDBNPlatinum.setFrm = this;
                    //emailStmntDBNPlatinum.waitPeriod = 500;
                    //checkErrRslt = emailStmntDBNPlatinum.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    //emailStmntDBNPlatinum = null;

                    var emailStmntDBNPlatinum = new clsStatHtmlDBN();
                    emailStmntDBNPlatinum.emailFromName = "ABP Access Bank Plc - Statement";
                    emailStmntDBNPlatinum.emailFrom = "noreplyinfo@accessbankplc.com";
                    emailStmntDBNPlatinum.bankWebLink = "www.accessbankplc.com";
                    emailStmntDBNPlatinum.waitPeriod = 3000;
                    emailStmntDBNPlatinum.ProductCondition = "('VISA Platinum')";
                    emailStmntDBNPlatinum.VIPCondition = "('Visa Platinum','Visa Platinum Staff','Visa Platinum Xclusive','Visa Platinum Savings Xtra Card')";
                    emailStmntDBNPlatinum.HasAttachement = true;
                    emailStmntDBNPlatinum.setFrm = this;
                    checkErrRslt = emailStmntDBNPlatinum.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntDBNPlatinum = null;

                    break;

                case 340:   // 36) DBN Diamond Bank Nigeria >> Email Platinum-USD 16/m
                    //clsStatementDBN_Html_New emailStmntDBNPlatinumUSD = new clsStatementDBN_Html_New();
                    //emailStmntDBNPlatinumUSD.emailFromName = "Diamond Bank PLC - Statement";
                    //emailStmntDBNPlatinumUSD.emailFrom = "cards@diamondbank.com";
                    //emailStmntDBNPlatinumUSD.bankWebLink = "www.DiamondBank.com";
                    //emailStmntDBNPlatinumUSD.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    //emailStmntDBNPlatinumUSD.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_header.jpg";
                    //emailStmntDBNPlatinumUSD.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_footer.jpg";
                    //emailStmntDBNPlatinumUSD.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\Statement_banner.jpg";
                    //emailStmntDBNPlatinumUSD.bankCard = @"D:\pC#\ProjData\Statement\DBN Diamond\Platinum-USD.jpg";
                    //emailStmntDBNPlatinumUSD.RewardLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\Reward_Banner.jpg";
                    //emailStmntDBNPlatinumUSD.logoAlignment = "left";
                    //emailStmntDBNPlatinumUSD.isReward = true;
                    //emailStmntDBNPlatinumUSD.rewardCond = "'Reward Program'";
                    //emailStmntDBNPlatinumUSD.ProductCondition = "('VISA Platinum Credit - USD')";
                    //emailStmntDBNPlatinumUSD.VIPCondition = "('Visa Platinum CR - USD Main','Visa Platinum CR - USD Supp.')";
                    //emailStmntDBNPlatinumUSD.setFrm = this;
                    //emailStmntDBNPlatinumUSD.waitPeriod = 500;
                    //checkErrRslt = emailStmntDBNPlatinumUSD.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    //emailStmntDBNPlatinumUSD = null;

                    var emailStmntDBNPlatinumUSD = new clsStatHtmlDBN();
                    emailStmntDBNPlatinumUSD.emailFromName = "ABP Access Bank Plc - Statement";
                    emailStmntDBNPlatinumUSD.emailFrom = "noreplyinfo@accessbankplc.com";
                    emailStmntDBNPlatinumUSD.bankWebLink = "www.accessbankplc.com";
                    emailStmntDBNPlatinumUSD.waitPeriod = 3000;
                    emailStmntDBNPlatinumUSD.ProductCondition = "('VISA Platinum Credit - USD')";
                    emailStmntDBNPlatinumUSD.VIPCondition = "('Visa Platinum CR - USD Main','Visa Platinum CR - USD Supp.')";
                    emailStmntDBNPlatinumUSD.HasAttachement = true;
                    emailStmntDBNPlatinumUSD.setFrm = this;
                    checkErrRslt = emailStmntDBNPlatinumUSD.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntDBNPlatinumUSD = null;

                    break;

                case 98:   // 36] DBN Diamond Bank Nigeria - VIP >> Email 16/m
                    //clsStatementDBN_Html_New emailStmntDBNvip = new clsStatementDBN_Html_New();
                    ////emailStmntDBNvip.emailFromName = "Diamond Bank Nigeria - Statement";
                    ////DBN-2628
                    //emailStmntDBNvip.emailFromName = "Diamond Bank PLC - Statement";
                    ////emailStmntDBNvip.emailFrom = "diamondvisacard@diamondbank.com";
                    ////DBN-5403
                    //emailStmntDBNvip.emailFrom = "cards@diamondbank.com";
                    //emailStmntDBNvip.bankWebLink = "www.DiamondBank.com";
                    //emailStmntDBNvip.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    ////emailStmntDBNvip.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\logo.gif";
                    ////emailStmntDBNvip.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\Header.jpg";
                    ////emailStmntDBNvip.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\VISA_E_STATEMENT_header.jpg";
                    //emailStmntDBNvip.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_header.jpg";
                    ////emailStmntDBNvip.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\VisaCreditCard_MidBanner.jpg";
                    ////emailStmntDBNvip.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\important_note.jpg";
                    //emailStmntDBNvip.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_footer.jpg";
                    ////emailStmntDBNvip.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\VisaCreditCard_BottomBanner.jpg";
                    ////emailStmntDBNvip.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\VISA_E_STATEMENT_banner.jpg";
                    //emailStmntDBNvip.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\Statement_banner.jpg";
                    //emailStmntDBNvip.bankCard = @"D:\pC#\ProjData\Statement\DBN Diamond\Gold.jpg";
                    //emailStmntDBNvip.RewardLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\Reward_Banner.jpg";
                    //emailStmntDBNvip.logoAlignment = "left";
                    //emailStmntDBNvip.isReward = true;
                    //emailStmntDBNvip.rewardCond = "'Reward Program'";
                    //emailStmntDBNvip.ProductCondition = "VISA Gold";
                    ////emailStmntDBNvip.VIPCondition = "('Visa Gold VIP3 UnSecuredNoCashNoCRFreeDays','Visa Gold VIP4 UnSecuredNoCashNoInterestFees','VISA Platinum - EXCO VIP','Visa Platinum Private Banking Customer')";
                    //emailStmntDBNvip.VIPCondition = "('Visa Gold VIP3 UnSecuredNoCashNoCRFreeDays','Visa Gold VIP4 UnSecuredCashNoInterestFees')";
                    ////emailStmntDBNvip.VIPCondition = "('Visa Gold VIP3 UnSecuredNoCashNoCRFreeDays','Visa Gold VIP4 UnSecuredNoCashNoInterestFees')";
                    //emailStmntDBNvip.setFrm = this;
                    ////emailStmntDBNvip.waitPeriod = 500;
                    //emailStmntDBNvip.waitPeriod = 500;
                    //checkErrRslt = emailStmntDBNvip.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    //emailStmntDBNvip = null;


                    var emailStmntDBNvip = new clsStatHtmlDBN();
                    emailStmntDBNvip.emailFromName = "ABP Access Bank Plc - Statement";
                    emailStmntDBNvip.emailFrom = "noreplyinfo@accessbankplc.com";
                    emailStmntDBNvip.bankWebLink = "www.accessbankplc.com";
                    emailStmntDBNvip.waitPeriod = 3000;
                    emailStmntDBNvip.ProductCondition = "VISA Gold";
                    emailStmntDBNvip.VIPCondition = "('Visa Gold VIP3 UnSecuredNoCashNoCRFreeDays','Visa Gold VIP4 UnSecuredCashNoInterestFees')";
                    emailStmntDBNvip.HasAttachement = true;
                    emailStmntDBNvip.setFrm = this;
                    checkErrRslt = emailStmntDBNvip.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntDBNvip = null;

                    break;

                case 112:   // [36] DBN Diamond Bank Nigeria >> VISA Platinum - ParkNShop Co-Brand Email 16/m
                    //clsStatementDBN_Html_New emailStmntDBNps = new clsStatementDBN_Html_New();
                    ////emailStmntDBNps.emailFromName = "Diamond Bank Nigeria - Statement";
                    ////DBN-2628
                    //emailStmntDBNps.emailFromName = "Diamond Bank PLC - Statement";
                    ////emailStmntDBNps.emailFrom = "diamondvisacard@diamondbank.com";
                    ////DBN-5403
                    //emailStmntDBNps.emailFrom = "cards@diamondbank.com";
                    //emailStmntDBNps.bankWebLink = "www.DiamondBank.com";
                    //emailStmntDBNps.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    ////emailStmntDBNps.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\logo.gif";
                    ////emailStmntDBNps.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\Header.jpg";
                    //emailStmntDBNps.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_header.jpg";
                    ////emailStmntDBNps.ADBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\PaknShop_MidBanner.jpg";
                    ////emailStmntDBNps.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\PaknShop_BottomBanner.jpg";
                    //emailStmntDBNps.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\Statement_banner.jpg";
                    //emailStmntDBNps.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_footer.jpg";
                    //emailStmntDBNps.bankCard = @"D:\pC#\ProjData\Statement\DBN Diamond\PNS.jpg";
                    //emailStmntDBNps.RewardLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\Reward_Banner.jpg";
                    //emailStmntDBNps.logoAlignment = "left";
                    //emailStmntDBNps.isReward = true;
                    //emailStmntDBNps.rewardCond = "('ParkNShop Reward Program')";
                    //emailStmntDBNps.ProductCondition = "VISA ParkNShop";
                    //emailStmntDBNps.VIPCondition = "('VISA Platinum - ParkNShop Co-Brand')";
                    //emailStmntDBNps.setFrm = this;
                    ////emailStmntDBNps.waitPeriod = 500;
                    //emailStmntDBNps.waitPeriod = 500;
                    //checkErrRslt = emailStmntDBNps.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    //emailStmntDBNps = null;


                    var emailStmntDBNps = new clsStatHtmlDBN();
                    emailStmntDBNps.emailFromName = "ABP Access Bank Plc - Statement";
                    emailStmntDBNps.emailFrom = "noreplyinfo@accessbankplc.com";
                    emailStmntDBNps.bankWebLink = "www.accessbankplc.com";
                    emailStmntDBNps.waitPeriod = 3000;
                    emailStmntDBNps.ProductCondition = "VISA ParkNShop";
                    emailStmntDBNps.VIPCondition = "('VISA Platinum - ParkNShop Co-Brand')";
                    emailStmntDBNps.HasAttachement = true;
                    emailStmntDBNps.setFrm = this;
                    checkErrRslt = emailStmntDBNps.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntDBNps = null;

                    break;

                case 150:   // [36] DBN Diamond Bank Nigeria >> VISA SPECIAL EXCO VIP - Privilege Email 16/m
                    //clsStatementDBN_Html_New emailStmntDBNexco = new clsStatementDBN_Html_New();
                    ////emailStmntDBNexco.emailFromName = "Diamond Bank Nigeria - Statement";
                    ////DBN-2628
                    //emailStmntDBNexco.emailFromName = "Diamond Bank PLC - Statement";
                    ////emailStmntDBNexco.emailFrom = "diamondvisacard@diamondbank.com";
                    ////DBN-5403
                    //emailStmntDBNexco.emailFrom = "cards@diamondbank.com";
                    //emailStmntDBNexco.bankWebLink = "www.DiamondBank.com";
                    //emailStmntDBNexco.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    ////emailStmntDBNexco.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\Header.jpg";
                    ////emailStmntDBNexco.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\VISA_E_STATEMENT_header.jpg";
                    //emailStmntDBNexco.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_header.jpg";
                    ////emailStmntDBNexco.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\Diamond_Card_MidBanner.jpg";
                    ////emailStmntDBNexco.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\important_note.jpg";
                    //emailStmntDBNexco.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_footer.jpg";
                    ////emailStmntDBNexco.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\Diamond_Card_BottomBanner.jpg";
                    ////emailStmntDBNexco.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\VISA_E_STATEMENT_banner.jpg";
                    //emailStmntDBNexco.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\Statement_banner.jpg";
                    //emailStmntDBNexco.bankCard = @"D:\pC#\ProjData\Statement\DBN Diamond\Privilege.jpg";
                    //emailStmntDBNexco.RewardLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\Reward_Banner.jpg";
                    //emailStmntDBNexco.logoAlignment = "left";
                    //emailStmntDBNexco.isReward = true;
                    //emailStmntDBNexco.rewardCond = "('Reward Program')";
                    //emailStmntDBNexco.ProductCondition = "VISA SPECIAL EXCO VIP";
                    //emailStmntDBNexco.VIPCondition = "('Visa Platinum Private Banking Customer','VISA Platinum - EXCO VIP')";
                    //emailStmntDBNexco.setFrm = this;
                    ////emailStmntDBNexco.waitPeriod = 500;
                    //emailStmntDBNexco.waitPeriod = 500;
                    //checkErrRslt = emailStmntDBNexco.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    //emailStmntDBNexco = null;


                    var emailStmntDBNexco = new clsStatHtmlDBN();
                    emailStmntDBNexco.emailFromName = "ABP Access Bank Plc - Statement";
                    emailStmntDBNexco.emailFrom = "noreplyinfo@accessbankplc.com";
                    emailStmntDBNexco.bankWebLink = "www.accessbankplc.com";
                    emailStmntDBNexco.waitPeriod = 3000;
                    emailStmntDBNexco.ProductCondition = "VISA SPECIAL EXCO VIP";
                    emailStmntDBNexco.VIPCondition = "('Visa Platinum Private Banking Customer','VISA Platinum - EXCO VIP')";
                    emailStmntDBNexco.HasAttachement = true;
                    emailStmntDBNexco.setFrm = this;
                    checkErrRslt = emailStmntDBNexco.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntDBNexco = null;

                    break;

                case 221:   // 36) DBN Diamond Bank Nigeria >> MasterCard Credit Email 16/m
                    //clsStatementDBN_Html_New emailStmntDBNmc = new clsStatementDBN_Html_New();
                    //emailStmntDBNmc.emailFromName = "Diamond Bank PLC - Statement";
                    ////emailStmntDBNmc.emailFrom = "diamondvisacard@diamondbank.com";
                    ////DBN-5403
                    //emailStmntDBNmc.emailFrom = "cards@diamondbank.com";
                    //emailStmntDBNmc.bankWebLink = "www.DiamondBank.com";
                    //emailStmntDBNmc.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    //emailStmntDBNmc.bankLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_header.jpg";
                    //emailStmntDBNmc.bankMidBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\statement_footer.jpg";
                    //emailStmntDBNmc.bankBottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\MasterCard_Footer.jpg";
                    //emailStmntDBNmc.bankCardGold = @"D:\pC#\ProjData\Statement\DBN Diamond\MasterCard_Gold.jpg";
                    //emailStmntDBNmc.bankCardPlatinum = @"D:\pC#\ProjData\Statement\DBN Diamond\MasterCard_Platinum.jpg";
                    //emailStmntDBNmc.RewardLogo = @"D:\pC#\ProjData\Statement\DBN Diamond\Reward_Banner.jpg";
                    //emailStmntDBNmc.isReward = true;
                    //emailStmntDBNmc.rewardCond = "('Reward Program')";
                    //emailStmntDBNmc.logoAlignment = "left";
                    //emailStmntDBNmc.VIPCondition = "('MasterCard Platinum Staff','MasterCard Platinum Savings Xtra Card','MasterCard Gold Savings Xtra Card','MasterCard Gold Xclusive','MasterCard Platinum Xclusive','MasterCard Gold Staff','MasterCard Platinum Standard','MasterCard Gold Standard','MasterCard Platinum Credit','MasterCard Gold Credit')";
                    //emailStmntDBNmc.ProductCondition = "('MasterCard Gold Credit','MasterCard Platinum Credit')";
                    //emailStmntDBNmc.setFrm = this;
                    //emailStmntDBNmc.waitPeriod = 500;
                    //checkErrRslt = emailStmntDBNmc.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    //emailStmntDBNmc = null;


                    var emailStmntDBNmc = new clsStatHtmlDBN();
                    emailStmntDBNmc.emailFromName = "ABP Access Bank Plc - Statement";
                    emailStmntDBNmc.emailFrom = "noreplyinfo@accessbankplc.com";
                    emailStmntDBNmc.bankWebLink = "www.accessbankplc.com";
                    emailStmntDBNmc.waitPeriod = 3000;
                    emailStmntDBNmc.VIPCondition = "('MasterCard Platinum Staff','MasterCard Platinum Savings Xtra Card','MasterCard Gold Savings Xtra Card','MasterCard Gold Xclusive','MasterCard Platinum Xclusive','MasterCard Gold Staff','MasterCard Platinum Standard','MasterCard Gold Standard','MasterCard Platinum Credit','MasterCard Gold Credit')";
                    emailStmntDBNmc.ProductCondition = "('MasterCard Gold Credit','MasterCard Platinum Credit')";
                    emailStmntDBNmc.HasAttachement = true;
                    emailStmntDBNmc.setFrm = this;
                    checkErrRslt = emailStmntDBNmc.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntDBNmc = null;

                    break;

                case 51:   // 25) AUB Ahli United Bank >> Text 1/m
                    //clsMaintainData maintainDataaubc = new clsMaintainData();
                    //maintainDataaubc.fixAddress(bankCode);
                    //maintainDataaubc = null;
                    clsStatementAUB stmntTextAUB = new clsStatementAUB();
                    stmntTextAUB.setFrm = this;
                    stmntTextAUB.InstallmentCondition = "('Purchase Installment With Interest Rate','Cash Installment')";
                    stmntTextAUB.isInstallmentVal = true;
                    checkErrRslt = stmntTextAUB.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);
                    stmntTextAUB = null;
                    break;

                //case 59:   // Test Statement
                //    //clsMaintainData maintainDatatst = new clsMaintainData();
                //    //maintainDatatst.fixAddress(bankCode);
                //    //maintainDatatst = null;
                //    clsStatementBNPcompIndv stmntTest = new clsStatementBNPcompIndv();
                //    stmntTest.setFrm = this;
                //    //stmntTest.CreateCorporate = true;
                //    checkErrRslt = stmntTest.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);
                //    stmntTest = null;
                //    break;

                case 62:   // 37) KBL PLATINUM HABIB BANK PLC >> Emails 15/m
                    //chkDontPrompt.Checked = true;
                    //clsStatHtmlGnrl01 emailStmntPHB = new clsStatHtmlGnrl01();
                    clsStatHtmlGnrl00 emailStmntPHB = new clsStatHtmlGnrl00();
                    emailStmntPHB.emailFromName = "Keystone Bank - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    //emailStmntPHB.emailFrom = "Ebankinghelpdesk@bankphb.com";//phbebankinghelpdesk@emp-group.com
                    emailStmntPHB.emailFrom = "Ebankingcustomerservice@keystonebankng.com";//phbebankinghelpdesk@emp-group.com
                    emailStmntPHB.bankWebLink = "www.keystonebankng.com/";
                    emailStmntPHB.bankLogo = @"D:\pC#\ProjData\Statement\PHB\logo.jpg";
                    emailStmntPHB.visaLogo = @"D:\pC#\ProjData\Statement\_Background\VisaLogo.gif";
                    emailStmntPHB.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\PHB\EmailMessageMonthly.txt";
                    //emailStmntPHB.waitPeriod = 500;
                    emailStmntPHB.waitPeriod = 500;
                    emailStmntPHB.setFrm = this;
                    checkErrRslt = emailStmntPHB.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntPHB = null;
                    break;

                case 320: //  [77] I&M Bank Rwanda Limited  >> Credit Emails 15/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlIMB emailStmntIMB = new clsStatHtmlIMB();
                    emailStmntIMB.emailFromName = "Ebanking@imbank.co.rw";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    //emailStmntIMB.emailFrom = "Ebankinghelpdesk@bankphb.com";//phbebankinghelpdesk@emp-group.com
                    emailStmntIMB.emailFrom = "cardnotifierRW@imbank.co.rw";//phbebankinghelpdesk@emp-group.com
                    emailStmntIMB.bankWebLink = "www.imbank.com/rwanda";
                    emailStmntIMB.bankLogo = @"D:\pC#\ProjData\Statement\IMB\logo.jpg";
                    emailStmntIMB.visaLogo = @"D:\pC#\ProjData\Statement\_Background\VisaLogo.gif";
                    emailStmntIMB.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\IMB\EmailMessageMonthly.txt";
                    //emailStmntPHB.waitPeriod = 500;
                    emailStmntIMB.waitPeriod = 500;
                    emailStmntIMB.setFrm = this;
                    checkErrRslt = emailStmntIMB.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntIMB = null;
                    break;

                case 325: //  [77] I&M Bank Rwanda Limited  >> Prepaid Emails 15/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrIMBPrepaid emailStmntIMBPre = new clsStatHtmlGnrIMBPrepaid();
                    emailStmntIMBPre.emailFromName = "Ebanking@imbank.co.rw";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    //emailStmntIMBPre.emailFrom = "Ebankinghelpdesk@bankphb.com";//phbebankinghelpdesk@emp-group.com
                    emailStmntIMBPre.emailFrom = "cardnotifierRW@imbank.co.rw";//phbebankinghelpdesk@emp-group.com
                    emailStmntIMBPre.bankWebLink = "www.imbank.com/rwanda";
                    emailStmntIMBPre.bankLogo = @"D:\pC#\ProjData\Statement\IMB\logo.jpg";
                    emailStmntIMBPre.visaLogo = @"D:\pC#\ProjData\Statement\_Background\VisaLogo.gif";
                    emailStmntIMBPre.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\IMB\EmailMessageMonthly.txt";
                    //emailStmntIMBPre.waitPeriod = 500;
                    emailStmntIMBPre.PrepaidCondition = "('Visa Classic Prepaid')";
                    emailStmntIMBPre.waitPeriod = 500;
                    emailStmntIMBPre.setFrm = this;
                    checkErrRslt = emailStmntIMBPre.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntIMBPre = null;
                    break;

                case 70:   // 27) BCA BANCO COMERCIAL DO ATLANTICO >> Corporate >> Default PDF 1/m
                    //clsMaintainData maintainDatabcacorp = new clsMaintainData();
                    //maintainDatabcacorp.fixAddress(bankCode);
                    //maintainDatabcacorp = null;
                    clsStatement_CommonCorpExp stmntCorporateBCA = new clsStatement_CommonCorpExp();
                    stmntCorporateBCA.mantainBank(bankCode);
                    stmntCorporateBCA.reportCompany = Application.StartupPath + @"\Reports\Statement_Corp_Com_BCA.rpt";
                    stmntCorporateBCA.reportIndividual = Application.StartupPath + @"\Reports\Statement_Corp_Ind_BCA.rpt";
                    stmntCorporateBCA.export(txtFileName.Text, strStatementType, bankCode, strFileName, Application.StartupPath + @"\Reports\Statement_Corp_Com_BCA.rpt", ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                    stmntCorporateBCA.CreateZip();
                    stmntCorporateBCA = null;
                    break;

                case 140:  //7) BAI Corporate >> PDF 1/m
                    clsStatement_CommonCorpExp stmntCorporateBAI = new clsStatement_CommonCorpExp();
                    stmntCorporateBAI.mantainBank(bankCode);
                    stmntCorporateBAI.reportCompany = Application.StartupPath + @"\Reports\Statement_BAI_Portuguese_English_Corporate.rpt";
                    stmntCorporateBAI.reportIndividual = Application.StartupPath + @"\Reports\Statement_BAI_Portuguese_Business.rpt";
                    stmntCorporateBAI.export(txtFileName.Text, strStatementType, bankCode, strFileName, Application.StartupPath + @"\Reports\Statement_BAI_Portuguese_English_Corporate.rpt", ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                    stmntCorporateBAI.CreateZip();
                    stmntCorporateBAI = null;
                    break;

                case 160:   // 16) ABP Access Bank Plc with Label Corporate >> PDF 1/m
                    //clsMaintainData maintainDataabpcorp = new clsMaintainData();
                    //maintainDataabpcorp.fixAddress(bankCode);
                    //maintainDataabpcorp = null;
                    clsStatement_CommonCorpExp stmntCorporateABP = new clsStatement_CommonCorpExp();
                    stmntCorporateABP.mantainBank(bankCode);
                    stmntCorporateABP.reportCompany = Application.StartupPath + @"\Reports\Statement_Common_Corporate_Company.rpt";
                    stmntCorporateABP.reportIndividual = Application.StartupPath + @"\Reports\Statement_Common_Corporate_Individual.rpt";
                    stmntCorporateABP.SplitByCurrency(txtFileName.Text, strStatementType, bankCode, strFileName, reportFleName, ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                    //stmntCorporateABP.export(txtFileName.Text, strStatementType, bankCode, strFileName, Application.StartupPath + @"\Reports\Statement_Common_Corporate_Company.rpt", ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                    stmntCorporateABP.CreateZip();
                    stmntCorporateABP = null;
                    break;

                case 93:   // 58] SBP SKYE BANK PLC >> Emails 1/m
                    clsStatHtmlSBP statHtmlSBPUSD = new clsStatHtmlSBP();
                    statHtmlSBPUSD.emailFromName = "Polaris BANK - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlSBPUSD.emailFrom = "skyevisa@skyebankng.com";//mmohammed@emp-group.com
                    statHtmlSBPUSD.bankWebLink = "www.polarisbanklimited.com";
                    //statHtmlSBPUSD.bankLogo = @"D:\pC#\ProjData\Statement\SBP\logo.gif";
                    statHtmlSBPUSD.bankLogo = @"D:\pC#\ProjData\Statement\SBP\Logo.png";
                    //statHtmlSBPUSD.topPicture = @"D:\pC#\ProjData\Statement\SBP\Top.gif";
                    statHtmlSBPUSD.ProductCondition = "('Visa Gold','Visa Platinum')";
                    //statHtmlSBPUSD.waitPeriod = 500;
                    statHtmlSBPUSD.waitPeriod = 500;
                    statHtmlSBPUSD.setFrm = this;
                    checkErrRslt = statHtmlSBPUSD.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlSBPUSD = null;
                    break;

                case 80:   // 58] SBP SKYE BANK PLC >> Emails 16/m
                case 253:   // 58] SBP SKYE BANK PLC >> MasterCard Platinum Credit Emails 16/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlSBP statHtmlSBPNGN = new clsStatHtmlSBP();
                    statHtmlSBPNGN.emailFromName = "Polaris BANK - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlSBPNGN.emailFrom = "creditcard-services@skyebankng.com";//mmohammed@emp-group.com
                    statHtmlSBPNGN.bankWebLink = "www.polarisbanklimited.com";
                    //statHtmlSBPNGN.bankLogo = @"D:\pC#\ProjData\Statement\SBP\logo.gif";
                    statHtmlSBPNGN.bankLogo = @"D:\pC#\ProjData\Statement\SBP\Logo.png";
                    if (pCmbProducts == 253)
                    {
                        //statHtmlSBPNGN.topPicture = @"D:\pC#\ProjData\Statement\SBP\Top3.jpg";
                        statHtmlSBPNGN.ProductCondition = "('MasterCard Credit Gold','MasterCard Credit Platinum')";
                    }
                    else
                    {
                        //statHtmlSBPNGN.topPicture = @"D:\pC#\ProjData\Statement\SBP\Top.gif";
                        statHtmlSBPNGN.ProductCondition = "('Visa Gold Premium','Visa Classic','Visa Classic - MEGALEK','VISA Platinum Credit NGN','VISA Platinum Credit NGN 419226')"; //VISA Platinum Credit NGN 419226 NP:"VISA Platinum Credit NGN" is not exist
                        //statHtmlSBPNGN.ProductCondition = "('Visa Gold Premium','Visa Classic','Visa Classic - MEGALEK','VISA Platinum Credit NGN')";
                    }
                    statHtmlSBPNGN.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\SBP\EmailMessageMonthly.txt";
                    //statHtmlSBPNGN.emailCC = "mscc_emails@yahoo.com";
                    //statHtmlSBPNGN.emailBCC = "mmohammed@emp-group.com";
                    //statHtmlSBPNGN.waitPeriod = 500;
                    statHtmlSBPNGN.waitPeriod = 500;
                    statHtmlSBPNGN.setFrm = this;
                    checkErrRslt = statHtmlSBPNGN.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlSBPNGN = null;
                    break;

                case 126:   // [58] SBP SKYE BANK PLC >> Corporate Emails 1/m
                    clsStatHtmlSBPDebit statHtmlSBPDB = new clsStatHtmlSBPDebit();
                    statHtmlSBPDB.emailFromName = "Polaris BANK - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlSBPDB.emailFrom = "skyevisa@skyebankng.com";//mmohammed@emp-group.com
                    statHtmlSBPDB.bankWebLink = "www.polarisbanklimited.com";
                    statHtmlSBPDB.bankLogo = @"D:\pC#\ProjData\Statement\SBP\Logo.png";
                    //statHtmlSBPDB.topPicture = @"D:\pC#\ProjData\Statement\SBP\Top.jpg";
                    statHtmlSBPDB.ProductCondition = "('Visa Cardholder Corporate USD')";
                    //statHtmlSBPDB.waitPeriod = 500;
                    statHtmlSBPDB.waitPeriod = 500;
                    statHtmlSBPDB.setFrm = this;
                    checkErrRslt = statHtmlSBPDB.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlSBPDB = null;
                    break;

                case 176:   // [58] SBP SKYE BANK PLC >> Prepaid VISA Emails 1/m
                case 184:   // [58] SBP SKYE BANK PLC >> Prepaid VISA-NTDC Emails 1/m
                case 177:   // [58] SBP SKYE BANK PLC >> Prepaid MasterCard Emails 1/m
                    clsStatHtmlSBPPrepaid statHtmlSBPPre = new clsStatHtmlSBPPrepaid();
                    statHtmlSBPPre.emailFromName = "Polaris BANK - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlSBPPre.emailFrom = "skyevisa@skyebankng.com";//mmohammed@emp-group.com
                    statHtmlSBPPre.bankWebLink = "www.polarisbanklimited.com";
                    statHtmlSBPPre.bankLogo = @"D:\pC#\ProjData\Statement\SBP\Logo.png";
                    //if (pCmbProducts == 176)
                    //{} //statHtmlSBPPre.topPicture = @"D:\pC#\ProjData\Statement\SBP\Top2.jpg";
                    statHtmlSBPPre.isPrepaid = true;
                    if (pCmbProducts == 176)
                        statHtmlSBPPre.PrepaidCondition = "('Visa Prepaid Travel Allowance Card')";
                    else if (pCmbProducts == 177)
                    {
                        //statHtmlSBPPre.topPicture = @"D:\pC#\ProjData\Statement\SBP\Top2.jpg";
                        statHtmlSBPPre.PrepaidCondition = "('MasterCard Prepaid')";
                    }
                    else if (pCmbProducts == 184)
                        statHtmlSBPPre.PrepaidCondition = "('Visa Prepaid-NTDC')";
                    //statHtmlSBPPre.waitPeriod = 500;
                    statHtmlSBPPre.waitPeriod = 500;
                    statHtmlSBPPre.setFrm = this;
                    checkErrRslt = statHtmlSBPPre.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlSBPPre = null;
                    break;

                case 255: //[113] GTBU Guaranty trust bank Uganada >> MasterCard Debit Emails 1/m
                    clsStatHtmlGTBUPrepaid statHtmlgtbudb = new clsStatHtmlGTBUPrepaid();
                    statHtmlgtbudb.emailFromName = "Guaranty trust bank Uganada - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlgtbudb.emailFrom = "cardservicesug@gtbank.com";//mmohammed@emp-group.com
                    statHtmlgtbudb.bankWebLink = "gtbank.co.ug";
                    statHtmlgtbudb.bankLogo = @"D:\pC#\ProjData\Statement\GTBU\logo.gif";
                    statHtmlgtbudb.isPrepaid = true;
                    statHtmlgtbudb.PrepaidCondition = "('MasterCard Prepaid General Spend Card')";
                    statHtmlgtbudb.waitPeriod = 500;
                    statHtmlgtbudb.setFrm = this;
                    checkErrRslt = statHtmlgtbudb.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlgtbudb = null;
                    break;

                case 192:   // 22) BPC BANCO DE POUPANCA E CREDITO, SARL Debit >> Email 1/m
                    clsStatHtmlBPCPrepaid statHtmlBPCpre = new clsStatHtmlBPCPrepaid();
                    statHtmlBPCpre.emailFromName = "Extracto Cartão BPC";//"BAI - Statement""mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBPCpre.emailFrom = "CartoesdeCredito@bpc.ao";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBPCpre.bankWebLink = "www.bpc.ao.";//"www.emp-group.com""www.socgen.com"
                    statHtmlBPCpre.bankWebLinkService = "www.bpc.ao.";//"www.socgen.com"
                    statHtmlBPCpre.bankLogo = @"D:\pC#\ProjData\Statement\BPC\Logo.jpg";
                    statHtmlBPCpre.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\BPC\EmailMessageMonthly.txt";
                    //statHtmlBPCpre.waitPeriod = 500;
                    statHtmlBPCpre.waitPeriod = 500;
                    statHtmlBPCpre.setFrm = this;
                    checkErrRslt = statHtmlBPCpre.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlBPCpre = null;
                    break;

                case 194:   // //32)UBAG UNITED BANK FOR AFRICA GHANA LIMITED >> Debit Email 16/m
                    clsStatHtmlUBAGPrepaid statHtmlUBAGPre = new clsStatHtmlUBAGPrepaid();
                    statHtmlUBAGPre.emailFromName = "UBAG - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlUBAGPre.emailFrom = "estatements@ubagroup.com";//mmohammed@emp-group.com
                    statHtmlUBAGPre.bankWebLink = "www.ubagroup.com";
                    statHtmlUBAGPre.bankLogo = @"D:\pC#\ProjData\Statement\UBAG\logo.jpg";
                    //statHtmlUBAGPre.waitPeriod = 500;
                    statHtmlUBAGPre.waitPeriod = 500;
                    statHtmlUBAGPre.setFrm = this;
                    checkErrRslt = statHtmlUBAGPre.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlUBAGPre = null;
                    break;

                //case 85:   // 7) BAI Products >> Email 1/m
                //    //chkDontPrompt.Checked = true;
                //    clsStatHtmlBAI statHtmlBAI = new clsStatHtmlBAI();
                //    statHtmlBAI.emailFromName = "Extractos VISA BAI";//"BAI - Statement""mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                //    statHtmlBAI.emailFrom = "bancobai@bancobai.ao";  //"baicartao@bancobai.ao";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                //    statHtmlBAI.bankWebLink = "www.bancobai.ao";//"www.emp-group.com""www.socgen.com"
                //    statHtmlBAI.bankWebLinkService = "bancobai@bancobai.ao";  //"baicartao@bancobai.ao";//"www.socgen.com"
                //    //          statHtmlBAI.bankLogo = @"D:\pC#\ProjData\Statement\BAI\Logo.gif";
                //    statHtmlBAI.bankLogo = @"D:\pC#\ProjData\Statement\BAI\Logo.bmp";
                //    //statHtmlBAI.backGround = @"D:\Web\Email\_Background\Background06.jpg";
                //    statHtmlBAI.topPicture = @"D:\pC#\ProjData\Statement\BAI\TopCards.gif";
                //    statHtmlBAI.downPicture = @"D:\pC#\ProjData\Statement\BAI\DownPic.gif";
                //    //statHtmlBAI.waitPeriod = 500;
                //    statHtmlBAI.waitPeriod = 500;
                //    //statHtmlBAI.emailBCC = "mmohammed@emp-group.com";
                //    statHtmlBAI.setFrm = this;
                //    //statHtmlBAI.emailBCC = "mamr@network.com.eg";

                //    checkErrRslt = statHtmlBAI.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                //    statHtmlBAI = null;
                //    break;

                case 460: //    [7] BAI Visa Classic >> Email 1/m
                case 461: //    [7] BAI Visa Gold >> Email 1/m
                case 462: //    [7] BAI Visa Platinum >> Email 1/m
                case 463: //    [7] BAI Visa Business >> Email 1/m
                    clsStatHtmlBAICredit statHtmlBAIPDF = new clsStatHtmlBAICredit();
                    if (pCmbProducts == 460)
                    {
                        statHtmlBAIPDF.ProductCondition = "('Visa Classic')";
                    }
                    if (pCmbProducts == 461)
                    {
                        statHtmlBAIPDF.ProductCondition = "('Visa Gold')";
                    }
                    if (pCmbProducts == 462)
                    {
                        statHtmlBAIPDF.ProductCondition = "('Visa Platinum')";
                    }
                    if (pCmbProducts == 463)
                    {
                        statHtmlBAIPDF.CreateCorporate = true;
                    }
                    // statHtmlBAIPDF.emailFromName = "Extractos VISA BAI";//"BAI - Statement""mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBAIPDF.emailFromName = "Extracto VISA BAI";//"BAI - Statement""mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"

                    statHtmlBAIPDF.emailFrom = "bancobai@bancobai.ao";  //"baicartao@bancobai.ao";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBAIPDF.bankWebLink = "www.bancobai.ao";//"www.emp-group.com""www.socgen.com"
                    statHtmlBAIPDF.bankWebLinkService = "bancobai@bancobai.ao";  //"baicartao@bancobai.ao";//"www.socgen.com"
                    statHtmlBAIPDF.bankLogo = @"D:\pC#\ProjData\Statement\BAI\Logo.bmp";
                    statHtmlBAIPDF.topPicture = @"D:\pC#\ProjData\Statement\BAI\topBanner.jpg";
                    statHtmlBAIPDF.downPicture = @"D:\pC#\ProjData\Statement\BAI\bottomBanner.jpg";
                    //statHtmlBAIPDF.waitPeriod = 500;
                    statHtmlBAIPDF.waitPeriod = 500;
                    statHtmlBAIPDF.setFrm = this;
                    checkErrRslt = statHtmlBAIPDF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlBAIPDF = null;
                    break;

                case 468: //158) RBGH Republic Bank (Ghana) Limited >> Default Credit Emails 23rd/m
                    clsStatHtmlRBGHCredit statHtmlRBGHPDF = new clsStatHtmlRBGHCredit();

                    statHtmlRBGHPDF.emailFromName = "RBGH Credit Card Statement";
                    statHtmlRBGHPDF.emailFrom = "alerts@republicghana.com";
                    statHtmlRBGHPDF.bankWebLink = "https://republicghana.com/";
                    statHtmlRBGHPDF.bankWebLinkService = "alerts@republicghana.com";
                    //statHtmlRBGHPDF.bankLogo = @"D:\pC#\ProjData\Statement\RBGH\Logo.bmp";
                    //statHtmlRBGHPDF.topPicture = @"D:\pC#\ProjData\Statement\RBGH\topBanner.jpg";
                    statHtmlRBGHPDF.downPicture = @"D:\pC#\ProjData\Statement\RBGH\bottomBanner.png";
                    //statHtmlRBGHPDF.waitPeriod = 500;
                    statHtmlRBGHPDF.waitPeriod = 500;
                    statHtmlRBGHPDF.setFrm = this;
                    checkErrRslt = statHtmlRBGHPDF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlRBGHPDF = null;
                    break;

                case 180:   // [87] WEMA BANK PLC NIGERIA >> Credit Emails 15/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlWEMA statHtmlWemaCredit = new clsStatHtmlWEMA();
                    //statHtmlWemaCredit.emailFromName = "WEMA Bank PLC - Statement";
                    statHtmlWemaCredit.emailFromName = "Purple Connect";
                    //statHtmlWemaCredit.emailFrom = "purpleconnect@wemabank.com";
                    statHtmlWemaCredit.emailFrom = "purpleconnect@emp-group.com";
                    //statHtmlWemaCredit.emailFrom = "mabouleila@emp-group.com";
                    statHtmlWemaCredit.bankWebLink = "www.wemabank.com";
                    statHtmlWemaCredit.bankWebLinkService = "purpleconnect@wemabank.com";
                    statHtmlWemaCredit.bankfacebookLink = "www.facebook.com/wemabankplc";
                    statHtmlWemaCredit.banktwitterLink = "www.twitter.com/wemabank";
                    statHtmlWemaCredit.bankyoutubeLink = "www.youtube.com/wematube";
                    statHtmlWemaCredit.bankgooglePlusLink = "www.google.com/+wematube";
                    statHtmlWemaCredit.bankLogo = @"D:\pC#\ProjData\Statement\WEMA\wemaLogo.jpg";
                    statHtmlWemaCredit.visaLogo = @"D:\pC#\ProjData\Statement\WEMA\visaLogo.jpg";
                    statHtmlWemaCredit.facebookLogo = @"D:\pC#\ProjData\Statement\WEMA\facebookLogo.png";
                    statHtmlWemaCredit.twitterLogo = @"D:\pC#\ProjData\Statement\WEMA\twitterLogo.png";
                    statHtmlWemaCredit.youtubeLogo = @"D:\pC#\ProjData\Statement\WEMA\youtubeLogo.png";
                    statHtmlWemaCredit.googlePlusLogo = @"D:\pC#\ProjData\Statement\WEMA\googlePlusLogo.png";
                    statHtmlWemaCredit.waitPeriod = 500;
                    statHtmlWemaCredit.setFrm = this;
                    checkErrRslt = statHtmlWemaCredit.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlWemaCredit = null;
                    break;

                //case 223:   // [47] Sterling BANK NIGERIA >> Credit Emails 1/m
                //    //chkDontPrompt.Checked = true;
                //    clsStatHtmlSBN_New statHtmlsbnCredit = new clsStatHtmlSBN_New();
                //    statHtmlsbnCredit.emailFromName = "Sterling Bank Nigeria - Statement";
                //    statHtmlsbnCredit.emailFrom = "VisaCreditCard@sterlingbankng.com";
                //    statHtmlsbnCredit.bankWebLink = "www.sterlingbankng.com";
                //    statHtmlsbnCredit.bankWebLinkService = "VisaCreditCard@sterlingbankng.com";
                //    statHtmlsbnCredit.bankLogo = @"D:\pC#\ProjData\Statement\SBN\SterlingBanklogo.gif";
                //    statHtmlsbnCredit.MidBanner = @"D:\pC#\ProjData\Statement\SBN\Main.jpg";
                //    statHtmlsbnCredit.BottomBanner = @"D:\pC#\ProjData\Statement\SBN\Bottom.jpg";
                //    statHtmlsbnCredit.ProductCondition = "('Visa Classic Credit Naira Contract','Visa Platinum Credit Naira Contract','VISA Gold Credit Naira Contract')";
                //    statHtmlsbnCredit.waitPeriod = 500;
                //    statHtmlsbnCredit.setFrm = this;
                //    checkErrRslt = statHtmlsbnCredit.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                //    statHtmlsbnCredit = null;
                //    break;

                //case 316:   // [47] Sterling BANK NIGERIA >> Credit Emails 15/m
                //    //chkDontPrompt.Checked = true;
                //    clsStatHtmlSBN_New statHtmlsbnCreditSD_15th = new clsStatHtmlSBN_New();
                //    statHtmlsbnCreditSD_15th.emailFromName = "Sterling Bank Nigeria - Statement";
                //    statHtmlsbnCreditSD_15th.emailFrom = "VisaCreditCard@sterlingbankng.com";
                //    statHtmlsbnCreditSD_15th.bankWebLink = "www.sterlingbankng.com";
                //    statHtmlsbnCreditSD_15th.bankWebLinkService = "VisaCreditCard@sterlingbankng.com";
                //    statHtmlsbnCreditSD_15th.bankLogo = @"D:\pC#\ProjData\Statement\SBN\SterlingBanklogo.gif";
                //    statHtmlsbnCreditSD_15th.MidBanner = @"D:\pC#\ProjData\Statement\SBN\Main.jpg";
                //    statHtmlsbnCreditSD_15th.BottomBanner = @"D:\pC#\ProjData\Statement\SBN\Bottom.jpg";
                //    statHtmlsbnCreditSD_15th.ProductCondition = "('Visa Infinite Credit (USD)','Visa Signature Credit (NGN)')";
                //    statHtmlsbnCreditSD_15th.waitPeriod = 500;
                //    statHtmlsbnCreditSD_15th.setFrm = this;
                //    checkErrRslt = statHtmlsbnCreditSD_15th.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                //    statHtmlsbnCreditSD_15th = null;
                //    break;


                case 223:   // [47] Sterling BANK NIGERIA >> Credit Emails 15/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlSBN_New_Signature statHtmlsbnCredit = new clsStatHtmlSBN_New_Signature();
                    //statHtmlsbnCredit.emailFromName = "Sterling Bank Nigeria - Statement";
                    statHtmlsbnCredit.emailFromName = "Credit card Statement";
                    //statHtmlsbnCredit.emailFrom = "VisaCreditCard@sterlingbankng.com";
                    //statHtmlsbnCredit.emailFrom = "e-statement@sterlingbankng.com";
                    statHtmlsbnCredit.emailFrom = "creditcardstatement@sterling.ng";
                    statHtmlsbnCredit.bankWebLink = "www.sterlingbankng.com";
                    statHtmlsbnCredit.bankWebLinkService = "VisaCreditCard@sterlingbankng.com";
                    statHtmlsbnCredit.bankLogo = @"D:\pC#\ProjData\Statement\SBN\OtherCRLogo.jpg";
                    statHtmlsbnCredit.MidBanner = @"D:\pC#\ProjData\Statement\SBN\Main.jpg";
                    statHtmlsbnCredit.BottomBanner = @"D:\pC#\ProjData\Statement\SBN\Bottom.jpg";
                    //statHtmlsbnCredit.ProductCondition = "('Visa Classic Credit Naira Contract','Visa Platinum Credit Naira Contract','VISA Gold Credit Naira Contract')";
                    statHtmlsbnCredit.ProductCondition = "('Visa Classic Credit Naira Contract','Visa Platinum Credit Naira Contract','VISA Gold Credit Naira Contract','Visa Corporate Credit - Naira','Visa Platinum Credit Dollar Contract','VISA Gold Credit Dollar Contract','Visa Corporate Credit - USD', 'Visa ULTRA Classic Credit Naira Contract', 'Visa ULTRA Platinum Credit Naira Contract', 'VISA ULTRA Gold Credit Naira Contract')";
                    statHtmlsbnCredit.waitPeriod = 500;
                    statHtmlsbnCredit.setFrm = this;
                    checkErrRslt = statHtmlsbnCredit.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlsbnCredit = null;
                    break;

                case 316:   // [47] Sterling BANK NIGERIA >> Credit Emails 15/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlSBN_New_Signature statHtmlsbnCreditSD_15th = new clsStatHtmlSBN_New_Signature();
                    //statHtmlsbnCreditSD_15th.emailFromName = "Sterling Bank Nigeria - Statement";
                    statHtmlsbnCreditSD_15th.emailFromName = "Credit card Statement";
                    //statHtmlsbnCreditSD_15th.emailFrom = "VisaCreditCard@sterlingbankng.com";
                    //statHtmlsbnCredit.emailFrom = "e-statement@sterlingbankng.com";
                    statHtmlsbnCreditSD_15th.emailFrom = "creditcardstatement@sterling.ng";
                    statHtmlsbnCreditSD_15th.bankWebLink = "www.sterlingbankng.com";
                    statHtmlsbnCreditSD_15th.bankWebLinkService = "VisaCreditCard@sterlingbankng.com";
                    statHtmlsbnCreditSD_15th.bankLogo = @"D:\pC#\ProjData\Statement\SBN\InfiniteLogo.jpg";
                    statHtmlsbnCreditSD_15th.MidBanner = @"D:\pC#\ProjData\Statement\SBN\Main.jpg";
                    statHtmlsbnCreditSD_15th.BottomBanner = @"D:\pC#\ProjData\Statement\SBN\Bottom.jpg";
                    statHtmlsbnCreditSD_15th.ProductCondition = "('Visa Infinite Credit (USD)')";
                    statHtmlsbnCreditSD_15th.waitPeriod = 500;
                    statHtmlsbnCreditSD_15th.setFrm = this;
                    checkErrRslt = statHtmlsbnCreditSD_15th.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlsbnCreditSD_15th = null;
                    break;
                case 401:   // [47] Sterling BANK NIGERIA >> Credit Emails 15/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlSBN_New_Signature statHtmlsbnCreditSD_15th_Signature = new clsStatHtmlSBN_New_Signature();
                    //statHtmlsbnCreditSD_15th_Signature.emailFromName = "Sterling Bank Nigeria - Statement";
                    statHtmlsbnCreditSD_15th_Signature.emailFromName = "Credit card Statement";
                    //statHtmlsbnCreditSD_15th_Signature.emailFrom = "VisaCreditCard@sterlingbankng.com";
                    //statHtmlsbnCredit.emailFrom = "e-statement@sterlingbankng.com";
                    statHtmlsbnCreditSD_15th_Signature.emailFrom = "creditcardstatement@sterling.ng";
                    statHtmlsbnCreditSD_15th_Signature.bankWebLink = "www.sterlingbankng.com";
                    statHtmlsbnCreditSD_15th_Signature.bankWebLinkService = "VisaCreditCard@sterlingbankng.com";
                    statHtmlsbnCreditSD_15th_Signature.bankLogo = @"D:\pC#\ProjData\Statement\SBN\SignatureLogo.jpg";
                    statHtmlsbnCreditSD_15th_Signature.MidBanner = @"D:\pC#\ProjData\Statement\SBN\Main.jpg";
                    statHtmlsbnCreditSD_15th_Signature.BottomBanner = @"D:\pC#\ProjData\Statement\SBN\Bottom.jpg";
                    statHtmlsbnCreditSD_15th_Signature.ProductCondition = "('Visa Signature Credit (NGN)', 'Visa ULTRA Signature Credit (NGN)')";
                    statHtmlsbnCreditSD_15th_Signature.waitPeriod = 500;
                    statHtmlsbnCreditSD_15th_Signature.setFrm = this;
                    checkErrRslt = statHtmlsbnCreditSD_15th_Signature.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlsbnCreditSD_15th_Signature = null;
                    break;

                case 224:   // 56) NCB National Commercial Bank >> Credit >> Email  1/m
                    clsStatementNCBcreditHtml statHtmlncbCredit = new clsStatementNCBcreditHtml();
                    statHtmlncbCredit.emailFromName = "National Commerical Bank Libya - Statement";
                    statHtmlncbCredit.emailFrom = "statement@ncb.ly";
                    statHtmlncbCredit.bankWebLink = "www.ncb.ly";
                    statHtmlncbCredit.bankWebLinkService = "statement@ncb.ly";
                    statHtmlncbCredit.bankLogo = @"D:\pC#\ProjData\Statement\NCB\logo1.gif";
                    statHtmlncbCredit.waitPeriod = 500;
                    statHtmlncbCredit.setFrm = this;
                    checkErrRslt = statHtmlncbCredit.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlncbCredit = null;
                    break;

                case 227:   // 56) NCB National Commercial Bank >> Prepaid >> Email  1/m
                    clsStatementNCBPreHtml statHtmlncbprepaid = new clsStatementNCBPreHtml();
                    statHtmlncbprepaid.emailFromName = "National Commerical Bank Libya - Statement";
                    statHtmlncbprepaid.emailFrom = "statement@ncb.ly";
                    statHtmlncbprepaid.bankWebLink = "www.ncb.ly";
                    statHtmlncbprepaid.bankWebLinkService = "statement@ncb.ly";
                    statHtmlncbprepaid.bankLogo = @"D:\pC#\ProjData\Statement\NCB\logo1.gif";
                    statHtmlncbprepaid.productCond = "('Electron Travel Prepaid')";
                    statHtmlncbprepaid.waitPeriod = 500;
                    statHtmlncbprepaid.setFrm = this;
                    checkErrRslt = statHtmlncbprepaid.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlncbprepaid = null;
                    break;

                case 183:   // [87] WEMA BANK PLC NIGERIA >> Debit/Prepaid Emails 1/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlWEMADebit statHtmlWemaDebit = new clsStatHtmlWEMADebit();
                    //statHtmlWemaDebit.emailFromName = "WEMA Bank PLC - Statement";
                    statHtmlWemaDebit.emailFromName = "Purple Connect";
                    //statHtmlWemaDebit.emailFrom = "purpleconnect@wemabank.com";
                    statHtmlWemaDebit.emailFrom = "purpleconnect@emp-group.com";
                    //statHtmlWemaDebit.emailFrom = "mabouleila@emp-group.com";
                    statHtmlWemaDebit.bankWebLink = "www.wemabank.com";
                    statHtmlWemaDebit.bankWebLinkService = "purpleconnect@wemabank.com";
                    statHtmlWemaDebit.bankfacebookLink = "www.facebook.com/wemabankplc";
                    statHtmlWemaDebit.banktwitterLink = "www.twitter.com/wemabank";
                    statHtmlWemaDebit.bankyoutubeLink = "www.youtube.com/wematube";
                    statHtmlWemaDebit.bankgooglePlusLink = "www.google.com/+wemabank";
                    statHtmlWemaDebit.bankLogo = @"D:\pC#\ProjData\Statement\WEMA\wemaLogo.jpg";
                    statHtmlWemaDebit.visaLogo = @"D:\pC#\ProjData\Statement\WEMA\visaLogo.jpg";
                    statHtmlWemaDebit.facebookLogo = @"D:\pC#\ProjData\Statement\WEMA\facebookLogo.png";
                    statHtmlWemaDebit.twitterLogo = @"D:\pC#\ProjData\Statement\WEMA\twitterLogo.png";
                    statHtmlWemaDebit.youtubeLogo = @"D:\pC#\ProjData\Statement\WEMA\youtubeLogo.png";
                    statHtmlWemaDebit.googlePlusLogo = @"D:\pC#\ProjData\Statement\WEMA\googlePlusLogo.png";
                    statHtmlWemaDebit.waitPeriod = 500;
                    statHtmlWemaDebit.isPrepaid = true;
                    statHtmlWemaDebit.PrepaidCondition = "('Visa Prepaid Classic (USD)')";
                    statHtmlWemaDebit.setFrm = this;
                    checkErrRslt = statHtmlWemaDebit.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlWemaDebit = null;
                    break;

                case 128: //    [7] BAI Prepaid >> Email 1/m
                    clsStatHtmlBAIPrepaid statHtmlBAIDB = new clsStatHtmlBAIPrepaid();
                    // statHtmlBAIDB.emailFromName = "Extractos VISA BAI";//"BAI - Statement""mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBAIDB.emailFromName = "Extracto VISA BAI";//"BAI - Statement""mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"

                    statHtmlBAIDB.emailFrom = "bancobai@bancobai.ao";  //"baicartao@bancobai.ao";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBAIDB.bankWebLink = "www.bancobai.ao";//"www.emp-group.com""www.socgen.com"
                    statHtmlBAIDB.bankWebLinkService = "bancobai@bancobai.ao";  //"baicartao@bancobai.ao";//"www.socgen.com"
                    statHtmlBAIDB.bankLogo = @"D:\pC#\ProjData\Statement\BAI\Logo.bmp";
                    statHtmlBAIDB.topPicture = @"D:\pC#\ProjData\Statement\BAI\topBanner.jpg";
                    statHtmlBAIDB.downPicture = @"D:\pC#\ProjData\Statement\BAI\bottomBanner.jpg";
                    //statHtmlBAIDB.waitPeriod = 500;
                    statHtmlBAIDB.waitPeriod = 500;
                    statHtmlBAIDB.setFrm = this;
                    checkErrRslt = statHtmlBAIDB.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    // var testDate=checkErrRslt[stmntDate];
                    statHtmlBAIDB = null;
                    break;

                //case 181: //    [7] BAI Corporate >> Email 1/m
                //    clsStatHtmlBAI statHtmlBAICorp = new clsStatHtmlBAI();
                //    statHtmlBAICorp.emailFromName = "Extractos VISA BAI";//"BAI - Statement""mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                //    statHtmlBAICorp.emailFrom = "bancobai@bancobai.ao";  //"baicartao@bancobai.ao";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                //    statHtmlBAICorp.bankWebLink = "www.bancobai.ao";//"www.emp-group.com""www.socgen.com"
                //    statHtmlBAICorp.bankWebLinkService = "bancobai@bancobai.ao";  //"baicartao@bancobai.ao";//"www.socgen.com"
                //    statHtmlBAICorp.bankLogo = @"D:\pC#\ProjData\Statement\BAI\Logo.bmp";
                //    statHtmlBAICorp.topPicture = @"D:\pC#\ProjData\Statement\BAI\corporate.jpg";
                //    statHtmlBAICorp.downPicture = @"D:\pC#\ProjData\Statement\BAI\DownPic.gif";
                //    statHtmlBAICorp.waitPeriod = 500;
                //    statHtmlBAICorp.CreateCorporate = true;
                //    statHtmlBAICorp.setFrm = this;
                //    checkErrRslt = statHtmlBAICorp.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                //    statHtmlBAICorp = null;
                //    break;

                case 96:   // [55] FBP Fidelity Bank PLC >> Emails 16/m
                case 219:   // [55] FBP Fidelity Bank PLC >> Emails 1/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlFBP statHtmlFBP = new clsStatHtmlFBP();
                    //FBP-986
                    statHtmlFBP.emailFromName = "Fidelity Bank";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    //FBP-986
                    statHtmlFBP.emailFrom = "ebanking@fidelitybank.ng";//mmohammed@emp-group.com
                    statHtmlFBP.bankWebLink = "fidelitybank.ng";
                    statHtmlFBP.bankLogo = @"D:\pC#\ProjData\Statement\FBP\logo.png";
                    statHtmlFBP.waitPeriod = 500;
                    statHtmlFBP.setFrm = this;
                    checkErrRslt = statHtmlFBP.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlFBP = null;
                    break;

                case 101:   // [13] BK Bank of Kigali >> Prepaid Emails 1/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrl04 statHtmlBDK = new clsStatHtmlGnrl04();
                    statHtmlBDK.emailFromName = " Bank of Kigali - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBDK.emailFrom = "VisaCardStatement@bk.rw";//mmohammed@emp-group.com
                    statHtmlBDK.bankWebLink = "www.bk.rw";
                    statHtmlBDK.bankLogo = @"D:\pC#\ProjData\Statement\BDK\logo.gif";
                    //statHtmlBDK.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    statHtmlBDK.backGround = @"D:\pC#\ProjData\Statement\_Background\Background01.jpg";
                    statHtmlBDK.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.gif";
                    //statHtmlBDK.waitPeriod = 500;
                    statHtmlBDK.waitPeriod = 3000;
                    statHtmlBDK.setFrm = this;
                    statHtmlBDK.isPrepaid = true;
                    statHtmlBDK.PrepaidCondition = "('Visa Prepaid')";
                    checkErrRslt = statHtmlBDK.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlBDK = null;
                    break;

                case 155:   // [13] BK Bank of Kigali >> VISA RWF Credit Emails 1/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrl04 statHtmlBDKCRRWF = new clsStatHtmlGnrl04();
                    statHtmlBDKCRRWF.emailFromName = " Bank of Kigali - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBDKCRRWF.emailFrom = "VisaCardStatement@bk.rw";//mmohammed@emp-group.com
                    //statHtmlBDKCRRWF.emailFrom = "bk@bk.rw";//mmohammed@emp-group.com
                    statHtmlBDKCRRWF.bankWebLink = "www.bk.rw";
                    statHtmlBDKCRRWF.bankLogo = @"D:\pC#\ProjData\Statement\BDK\logo.gif";
                    //statHtmlBDKCRRWF.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    statHtmlBDKCRRWF.backGround = @"D:\pC#\ProjData\Statement\_Background\Background01.jpg";
                    statHtmlBDKCRRWF.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.gif";
                    //statHtmlBDKCRRWF.waitPeriod = 500;
                    statHtmlBDKCRRWF.waitPeriod = 2000;
                    //statHtmlBDKCRRWF.CurrencyFilter = "RWF"; //iatta BDK
                    statHtmlBDKCRRWF.ProductCondition = "('Visa Gold Standard','Visa Classic Standard','Visa Platinum Paywave Standard','Visa Gold Staff','Visa Classic Staff','Visa Platinum Paywave Staff')";
                    //statHtmlBDKCRRWF.ProductCondition = "('Visa Gold Staff','Visa Classic Staff','Visa Platinum Paywave Staff')";
                    statHtmlBDKCRRWF.setFrm = this;
                    checkErrRslt = statHtmlBDKCRRWF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlBDKCRRWF = null;
                    break;

                case 287:   // [13] BK Bank of Kigali >> VISA USD Credit Emails 1/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrl04 statHtmlBDKCRUSD = new clsStatHtmlGnrl04();
                    statHtmlBDKCRUSD.emailFromName = " Bank of Kigali - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBDKCRUSD.emailFrom = "VisaCardStatement@bk.rw";//mmohammed@emp-group.com
                    //statHtmlBDKCRUSD.emailFrom = "bk@bk.rw";//mmohammed@emp-group.com
                    statHtmlBDKCRUSD.bankWebLink = "www.bk.rw";
                    statHtmlBDKCRUSD.bankLogo = @"D:\pC#\ProjData\Statement\BDK\logo.gif";
                    //statHtmlBDKCRUSD.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    statHtmlBDKCRUSD.backGround = @"D:\pC#\ProjData\Statement\_Background\Background01.jpg";
                    statHtmlBDKCRUSD.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.gif";
                    //statHtmlBDKCRUSD.waitPeriod = 500;
                    statHtmlBDKCRUSD.waitPeriod = 2000;
                    //statHtmlBDKCRUSD.CurrencyFilter = "USD"; //iatta BDK
                    statHtmlBDKCRUSD.ProductCondition = "('Visa Gold Standard','Visa Classic Standard','Visa Platinum Paywave Standard','Visa Gold Staff','Visa Classic Staff','Visa Platinum Paywave Staff')";
                    //statHtmlBDKCRUSD.ProductCondition = "('Visa Gold Staff','Visa Classic Staff','Visa Platinum Paywave Staff')";
                    statHtmlBDKCRUSD.setFrm = this;
                    checkErrRslt = statHtmlBDKCRUSD.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlBDKCRUSD = null;
                    break;

                case 313:   // [13] BK Bank of Kigali >> MasterCard RWF Credit Emails 1/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrl04 statHtmlBDKMCCRRWF = new clsStatHtmlGnrl04();
                    statHtmlBDKMCCRRWF.emailFromName = " Bank of Kigali - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBDKMCCRRWF.emailFrom = "VisaCardStatement@bk.rw";//mmohammed@emp-group.com
                    //statHtmlBDKMCCRRWF.emailFrom = "bk@bk.rw";//mmohammed@emp-group.com
                    statHtmlBDKMCCRRWF.bankWebLink = "www.bk.rw";
                    statHtmlBDKMCCRRWF.bankLogo = @"D:\pC#\ProjData\Statement\BDK\logo.gif";
                    //statHtmlBDKMCCRRWF.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    statHtmlBDKMCCRRWF.backGround = @"D:\pC#\ProjData\Statement\_Background\Background01.jpg";
                    statHtmlBDKMCCRRWF.MasterCardLogo = @"D:\pC#\ProjData\Statement\_Background\MasterCardLogo.jpg";
                    //statHtmlBDKMCCRRWF.waitPeriod = 500;
                    statHtmlBDKMCCRRWF.waitPeriod = 2000;
                    //statHtmlBDKMCCRRWF.CurrencyFilter = "RWF"; //iatta BDK
                    statHtmlBDKMCCRRWF.ProductCondition = "('MC Credit Classic Standard','MC Credit World Standard','MC Platinum Credit Standard', 'MC Credit Classic Staff', 'MC Credit World Staff', 'MC Platinum Credit Staff')";
                    //statHtmlBDKMCCRRWF.ProductCondition = "('MC Credit Classic Staff', 'MC Credit World Staff', 'MC Platinum Credit Staff')";
                    statHtmlBDKMCCRRWF.setFrm = this;
                    checkErrRslt = statHtmlBDKMCCRRWF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlBDKMCCRRWF = null;
                    break;

                case 314:   // [13] BK Bank of Kigali >> MasterCard USD Credit Emails 1/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrl04 statHtmlBDKMCCRUSD = new clsStatHtmlGnrl04();
                    statHtmlBDKMCCRUSD.emailFromName = " Bank of Kigali - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBDKMCCRUSD.emailFrom = "VisaCardStatement@bk.rw";//mmohammed@emp-group.com
                    //statHtmlBDKMCCRUSD.emailFrom = "bk@bk.rw";//mmohammed@emp-group.com
                    statHtmlBDKMCCRUSD.bankWebLink = "www.bk.rw";
                    statHtmlBDKMCCRUSD.bankLogo = @"D:\pC#\ProjData\Statement\BDK\logo.gif";
                    //statHtmlBDKMCCRUSD.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    statHtmlBDKMCCRUSD.backGround = @"D:\pC#\ProjData\Statement\_Background\Background01.jpg";
                    statHtmlBDKMCCRUSD.MasterCardLogo = @"D:\pC#\ProjData\Statement\_Background\MasterCardLogo.jpg";
                    //statHtmlBDKMCCRUSD.waitPeriod = 500;
                    statHtmlBDKMCCRUSD.waitPeriod = 2000;
                    //statHtmlBDKMCCRUSD.CurrencyFilter = "USD"; //iatta BDK
                    statHtmlBDKMCCRUSD.ProductCondition = "('MC Credit Classic Standard','MC Credit World Standard','MC Platinum Credit Standard', 'MC Credit Classic Staff', 'MC Credit World Staff', 'MC Platinum Credit Staff')";
                    //statHtmlBDKMCCRUSD.ProductCondition = "('MC Credit Classic Staff', 'MC Credit World Staff', 'MC Platinum Credit Staff')";
                    statHtmlBDKMCCRUSD.setFrm = this;
                    checkErrRslt = statHtmlBDKMCCRUSD.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlBDKMCCRUSD = null;
                    break;

                case 331:   // [13] BK Bank of Kigali >> VISA RWF Credit VIP Emails 5/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrl04 statHtmlBDKCRVIPRWF = new clsStatHtmlGnrl04();
                    statHtmlBDKCRVIPRWF.emailFromName = " Bank of Kigali - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBDKCRVIPRWF.emailFrom = "VisaCardStatement@bk.rw";//mmohammed@emp-group.com
                    //statHtmlBDKCRVIPRWF.emailFrom = "bk@bk.rw";//mmohammed@emp-group.com
                    statHtmlBDKCRVIPRWF.bankWebLink = "www.bk.rw";
                    statHtmlBDKCRVIPRWF.bankLogo = @"D:\pC#\ProjData\Statement\BDK\logo.gif";
                    //statHtmlBDKCRVIPRWF.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    statHtmlBDKCRVIPRWF.backGround = @"D:\pC#\ProjData\Statement\_Background\Background01.jpg";
                    statHtmlBDKCRVIPRWF.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.gif";
                    //statHtmlBDKCRVIPRWF.waitPeriod = 500;
                    statHtmlBDKCRVIPRWF.waitPeriod = 2000;
                    //statHtmlBDKCRVIPRWF.CurrencyFilter = "RWF"; //iatta BDK
                    statHtmlBDKCRVIPRWF.ProductCondition = "('Visa Gold VIP','Visa Classic VIP','Visa Platinum Paywave VIP')";
                    statHtmlBDKCRVIPRWF.setFrm = this;
                    checkErrRslt = statHtmlBDKCRVIPRWF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlBDKCRVIPRWF = null;
                    break;

                case 332:   // [13] BK Bank of Kigali >> VISA USD Credit VIP Emails 5/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrl04 statHtmlBDKCRVIPUSD = new clsStatHtmlGnrl04();
                    statHtmlBDKCRVIPUSD.emailFromName = " Bank of Kigali - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBDKCRVIPUSD.emailFrom = "VisaCardStatement@bk.rw";//mmohammed@emp-group.com
                    //statHtmlBDKCRVIPUSD.emailFrom = "bk@bk.rw";//mmohammed@emp-group.com
                    statHtmlBDKCRVIPUSD.bankWebLink = "www.bk.rw";
                    statHtmlBDKCRVIPUSD.bankLogo = @"D:\pC#\ProjData\Statement\BDK\logo.gif";
                    //statHtmlBDKCRVIPUSD.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    statHtmlBDKCRVIPUSD.backGround = @"D:\pC#\ProjData\Statement\_Background\Background01.jpg";
                    statHtmlBDKCRVIPUSD.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.gif";
                    //statHtmlBDKCRVIPUSD.waitPeriod = 500;
                    statHtmlBDKCRVIPUSD.waitPeriod = 2000;
                    //statHtmlBDKCRVIPUSD.CurrencyFilter = "USD"; //iatta BDK
                    statHtmlBDKCRVIPUSD.ProductCondition = "('Visa Gold VIP','Visa Classic VIP','Visa Platinum Paywave VIP')";
                    statHtmlBDKCRVIPUSD.setFrm = this;
                    checkErrRslt = statHtmlBDKCRVIPUSD.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlBDKCRVIPUSD = null;
                    break;

                case 333:   // [13] BK Bank of Kigali >> MasterCard RWF Credit VIP Emails 5/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrl04 statHtmlBDKMCCRVIPRWF = new clsStatHtmlGnrl04();
                    statHtmlBDKMCCRVIPRWF.emailFromName = " Bank of Kigali - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBDKMCCRVIPRWF.emailFrom = "VisaCardStatement@bk.rw";//mmohammed@emp-group.com
                    //statHtmlBDKMCCRVIPRWF.emailFrom = "bk@bk.rw";//mmohammed@emp-group.com
                    statHtmlBDKMCCRVIPRWF.bankWebLink = "www.bk.rw";
                    statHtmlBDKMCCRVIPRWF.bankLogo = @"D:\pC#\ProjData\Statement\BDK\logo.gif";
                    //statHtmlBDKMCCRVIPRWF.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    statHtmlBDKMCCRVIPRWF.backGround = @"D:\pC#\ProjData\Statement\_Background\Background01.jpg";
                    statHtmlBDKMCCRVIPRWF.MasterCardLogo = @"D:\pC#\ProjData\Statement\_Background\MasterCardLogo.jpg";
                    //statHtmlBDKMCCRVIPRWF.waitPeriod = 500;
                    statHtmlBDKMCCRVIPRWF.waitPeriod = 2000;
                    //statHtmlBDKMCCRVIPRWF.CurrencyFilter = "RWF"; //iatta BDK
                    statHtmlBDKMCCRVIPRWF.ProductCondition = "('MC Credit Classic VIP','MC Credit World VIP','MC Platinum Credit VIP')";
                    statHtmlBDKMCCRVIPRWF.setFrm = this;
                    checkErrRslt = statHtmlBDKMCCRVIPRWF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlBDKMCCRVIPRWF = null;
                    break;

                case 334:   // [13] BK Bank of Kigali >> MasterCard USD Credit VIP Emails 5/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrl04 statHtmlBDKMCCRVIPUSD = new clsStatHtmlGnrl04();
                    statHtmlBDKMCCRVIPUSD.emailFromName = " Bank of Kigali - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBDKMCCRVIPUSD.emailFrom = "VisaCardStatement@bk.rw";//mmohammed@emp-group.com
                    //statHtmlBDKMCCRVIPUSD.emailFrom = "bk@bk.rw";//mmohammed@emp-group.com
                    statHtmlBDKMCCRVIPUSD.bankWebLink = "www.bk.rw";
                    statHtmlBDKMCCRVIPUSD.bankLogo = @"D:\pC#\ProjData\Statement\BDK\logo.gif";
                    //statHtmlBDKMCCRVIPUSD.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    statHtmlBDKMCCRVIPUSD.backGround = @"D:\pC#\ProjData\Statement\_Background\Background01.jpg";
                    statHtmlBDKMCCRVIPUSD.MasterCardLogo = @"D:\pC#\ProjData\Statement\_Background\MasterCardLogo.jpg";
                    //statHtmlBDKMCCRVIPUSD.waitPeriod = 500;
                    statHtmlBDKMCCRVIPUSD.waitPeriod = 2000;
                    //statHtmlBDKMCCRVIPUSD.CurrencyFilter = "USD"; //iatta BDK
                    statHtmlBDKMCCRVIPUSD.ProductCondition = "('MC Credit Classic VIP','MC Credit World VIP','MC Platinum Credit VIP')";
                    statHtmlBDKMCCRVIPUSD.setFrm = this;
                    checkErrRslt = statHtmlBDKMCCRVIPUSD.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlBDKMCCRVIPUSD = null;
                    break;

                case 335:   // [13] BK Bank of Kigali >> MasterCard RWF Corporate cardholder VIP Emails 1/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrl04 statHtmlBDKMCCORPVIPRWF = new clsStatHtmlGnrl04();
                    statHtmlBDKMCCORPVIPRWF.emailFromName = " Bank of Kigali - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBDKMCCORPVIPRWF.emailFrom = "VisaCardStatement@bk.rw";//mmohammed@emp-group.com
                    //statHtmlBDKMCCORPVIPRWF.emailFrom = "bk@bk.rw";//mmohammed@emp-group.com
                    statHtmlBDKMCCORPVIPRWF.bankWebLink = "www.bk.rw";
                    statHtmlBDKMCCORPVIPRWF.bankLogo = @"D:\pC#\ProjData\Statement\BDK\logo.gif";
                    //statHtmlBDKMCCORPVIPRWF.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    statHtmlBDKMCCORPVIPRWF.backGround = @"D:\pC#\ProjData\Statement\_Background\Background01.jpg";
                    statHtmlBDKMCCORPVIPRWF.MasterCardLogo = @"D:\pC#\ProjData\Statement\_Background\MasterCardLogo.jpg";
                    //statHtmlBDKMCCORPVIPRWF.waitPeriod = 500;
                    statHtmlBDKMCCORPVIPRWF.waitPeriod = 2000;
                    //statHtmlBDKMCCORPVIPRWF.CurrencyFilter = "RWF"; //iatta BDK
                    statHtmlBDKMCCORPVIPRWF.ProductCondition = "('MC Cardholder Platinum Credit VIP')";
                    statHtmlBDKMCCORPVIPRWF.setFrm = this;
                    checkErrRslt = statHtmlBDKMCCORPVIPRWF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlBDKMCCORPVIPRWF = null;
                    break;

                case 336:   // [13] BK Bank of Kigali >> MasterCard USD Corporate cardholder VIP Emails 1/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrl04 statHtmlBDKMCCORPVIPUSD = new clsStatHtmlGnrl04();
                    statHtmlBDKMCCORPVIPUSD.emailFromName = " Bank of Kigali - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBDKMCCORPVIPUSD.emailFrom = "VisaCardStatement@bk.rw";//mmohammed@emp-group.com
                    //statHtmlBDKMCCORPVIPUSD.emailFrom = "bk@bk.rw";//mmohammed@emp-group.com
                    statHtmlBDKMCCORPVIPUSD.bankWebLink = "www.bk.rw";
                    statHtmlBDKMCCORPVIPUSD.bankLogo = @"D:\pC#\ProjData\Statement\BDK\logo.gif";
                    //statHtmlBDKMCCORPVIPUSD.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    statHtmlBDKMCCORPVIPUSD.backGround = @"D:\pC#\ProjData\Statement\_Background\Background01.jpg";
                    statHtmlBDKMCCORPVIPUSD.MasterCardLogo = @"D:\pC#\ProjData\Statement\_Background\MasterCardLogo.jpg";
                    //statHtmlBDKMCCORPVIPUSD.waitPeriod = 500;
                    statHtmlBDKMCCORPVIPUSD.waitPeriod = 2000;
                    //statHtmlBDKMCCORPVIPUSD.CurrencyFilter = "USD"; //iatta BDK
                    statHtmlBDKMCCORPVIPUSD.ProductCondition = "('MC Cardholder Platinum Credit VIP')";
                    statHtmlBDKMCCORPVIPUSD.setFrm = this;
                    checkErrRslt = statHtmlBDKMCCORPVIPUSD.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlBDKMCCORPVIPUSD = null;
                    break;

                case 337:   // [13] BK Bank of Kigali >> MasterCard RWF Corporate cardholder Emails 1/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrl04 statHtmlBDKMCCORPRWF = new clsStatHtmlGnrl04();
                    statHtmlBDKMCCORPRWF.emailFromName = " Bank of Kigali - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBDKMCCORPRWF.emailFrom = "VisaCardStatement@bk.rw";//mmohammed@emp-group.com
                    //statHtmlBDKMCCORPRWF.emailFrom = "bk@bk.rw";//mmohammed@emp-group.com
                    statHtmlBDKMCCORPRWF.bankWebLink = "www.bk.rw";
                    statHtmlBDKMCCORPRWF.bankLogo = @"D:\pC#\ProjData\Statement\BDK\logo.gif";
                    //statHtmlBDKMCCORPRWF.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    statHtmlBDKMCCORPRWF.backGround = @"D:\pC#\ProjData\Statement\_Background\Background01.jpg";
                    statHtmlBDKMCCORPRWF.MasterCardLogo = @"D:\pC#\ProjData\Statement\_Background\MasterCardLogo.jpg";
                    //statHtmlBDKMCCORPRWF.waitPeriod = 500;
                    statHtmlBDKMCCORPRWF.waitPeriod = 2000;
                    //statHtmlBDKMCCORPRWF.CurrencyFilter = "RWF"; //iatta BDK
                    statHtmlBDKMCCORPRWF.ProductCondition = "('MC Platinum Credit Standard')";
                    statHtmlBDKMCCORPRWF.setFrm = this;
                    checkErrRslt = statHtmlBDKMCCORPRWF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlBDKMCCORPRWF = null;
                    break;

                case 338:   // [13] BK Bank of Kigali >> MasterCard USD Corporate cardholder Emails 1/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrl04 statHtmlBDKMCCORPUSD = new clsStatHtmlGnrl04();
                    statHtmlBDKMCCORPUSD.emailFromName = " Bank of Kigali - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBDKMCCORPUSD.emailFrom = "VisaCardStatement@bk.rw";//mmohammed@emp-group.com
                    //statHtmlBDKMCCORPUSD.emailFrom = "bk@bk.rw";//mmohammed@emp-group.com
                    statHtmlBDKMCCORPUSD.bankWebLink = "www.bk.rw";
                    statHtmlBDKMCCORPUSD.bankLogo = @"D:\pC#\ProjData\Statement\BDK\logo.gif";
                    //statHtmlBDKMCCORPUSD.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    statHtmlBDKMCCORPUSD.backGround = @"D:\pC#\ProjData\Statement\_Background\Background01.jpg";
                    statHtmlBDKMCCORPUSD.MasterCardLogo = @"D:\pC#\ProjData\Statement\_Background\MasterCardLogo.jpg";
                    //statHtmlBDKMCCORPUSD.waitPeriod = 500;
                    statHtmlBDKMCCORPUSD.waitPeriod = 2000;
                    //statHtmlBDKMCCORPUSD.CurrencyFilter = "USD"; //iatta BDK
                    statHtmlBDKMCCORPUSD.ProductCondition = "('MC Platinum Credit Standard')";
                    statHtmlBDKMCCORPUSD.setFrm = this;
                    checkErrRslt = statHtmlBDKMCCORPUSD.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlBDKMCCORPUSD = null;
                    break;

                case 274:   // [50] ICBG Access Bank Ghana Limited >> Credit Emails 1/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlICBG statHtmlicbgCR = new clsStatHtmlICBG();
                    statHtmlicbgCR.emailFromName = "Access Bank Alert Service";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlicbgCR.emailFrom = "noreplyinfo@accessbankplc.com";//mmohammed@emp-group.com
                    //statHtmlicbgCR.emailFrom = "noreplyghana@accessbankplc.com";//mmohammed@emp-group.com
                    //statHtmlicbgCR.emailFrom = "Access.BankGH@emp-group.com";//mmohammed@emp-group.com
                    statHtmlicbgCR.bankWebLink = "subs.accessbankplc.com/gh/";
                    statHtmlicbgCR.bankLogo = @"D:\pC#\ProjData\Statement\ICBG\logo.jpg";
                    statHtmlicbgCR.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    statHtmlicbgCR.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.gif";
                    statHtmlicbgCR.bottomBanner = @"D:\pC#\ProjData\Statement\ICBG\bottom.jpg";
                    statHtmlicbgCR.waitPeriod = 3000;
                    statHtmlicbgCR.HasAttachement = true;
                    statHtmlicbgCR.setFrm = this;
                    checkErrRslt = statHtmlicbgCR.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlicbgCR = null;
                    break;

                case 106: //50] ICBG Intercontinental Bank Ghana Limited Corporate >> PDF 1/m
                    //clsMaintainData maintainDataicbgcorp = new clsMaintainData();
                    //maintainDataicbgcorp.fixAddress(bankCode);
                    //maintainDataicbgcorp = null;
                    clsStatement_CommonCorpExp stmntCorporateICBG = new clsStatement_CommonCorpExp();
                    stmntCorporateICBG.mantainBank(bankCode);
                    stmntCorporateICBG.reportCompany = Application.StartupPath + @"\Reports\Statement_ICBG_Corporate_Company.rpt";
                    //stmntCorporateICBG.reportIndividual = Application.StartupPath + @"\Reports\Statement_Common_Corporate_Individual.rpt";
                    stmntCorporateICBG.reportIndividual = Application.StartupPath + @"\Reports\Statement_ICBG_Credit.rpt";
                    //stmntCorporateICBG.export(txtFileName.Text, strStatementType, bankCode, strFileName, Application.StartupPath + @"\Reports\Statement_Corp_Com_SSB.rpt", ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                    stmntCorporateICBG.export(txtFileName.Text, strStatementType, bankCode, strFileName, Application.StartupPath + @"\Reports\Statement_Common_Corporate_Company.rpt", ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                    stmntCorporateICBG.CreateZip();
                    stmntCorporateICBG = null;
                    break;

                case 110:   //69] First Bank of Nigeria  >> Credit Emails 1/m
                case 124:   //69] First Bank of Nigeria  >> Credit Emails 16/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrlFBN statHtmlFBN = new clsStatHtmlGnrlFBN();
                    statHtmlFBN.emailFromName = "First Bank of Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlFBN.emailFrom = "cardservices@firstbanknigeria.com";//mmohammed@emp-group.com
                    statHtmlFBN.bankWebLink = "www.firstbanknigeria.com";
                    statHtmlFBN.bankLogo = @"D:\pC#\ProjData\Statement\FBN\logo.jpg";
                    //statHtmlFBN.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    //if (stmntType.Contains("MasterCard"))
                    //    statHtmlFBN.visaLogo = @"D:\pC#\ProjData\Statement\FBN\MClogo.jpg";
                    //else
                    statHtmlFBN.visaLogo = @"D:\pC#\ProjData\Statement\FBN\visalogo.jpg";
                    //statHtmlFBN.WaterMark = @"D:\pC#\ProjData\Statement\FBN\watermark.jpg";
                    statHtmlFBN.BottomBanner = @"D:\pC#\ProjData\Statement\FBN\banner.jpg";
                    // The new Image Jira FBN-8345
                    statHtmlFBN.AdImage = @"D:\pC#\ProjData\Statement\FBN\Adimage.jpg";
                    //statHtmlFBN.waitPeriod = 500;
                    statHtmlFBN.waitPeriod = 500;
                    statHtmlFBN.ProductCondition = "('VISA Infinite','VISA Gold - Standard','VISA Gold - Staff','VISA Classic','VISA Classic PLATINUM - NGN')";
                    statHtmlFBN.setFrm = this;
                    checkErrRslt = statHtmlFBN.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlFBN = null;
                    break;

                case 321:   //69] First Bank of Nigeria  >> Emails Classic 16/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrlFBN statHtmlFBNClassic = new clsStatHtmlGnrlFBN();
                    statHtmlFBNClassic.emailFromName = "First Bank of Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlFBNClassic.emailFrom = "cardservices@firstbanknigeria.com";//mmohammed@emp-group.com
                    statHtmlFBNClassic.bankWebLink = "www.firstbanknigeria.com";
                    statHtmlFBNClassic.bankLogo = @"D:\pC#\ProjData\Statement\FBN\logo.jpg";
                    statHtmlFBNClassic.visaLogo = @"D:\pC#\ProjData\Statement\FBN\visalogo.jpg";
                    statHtmlFBNClassic.BottomBanner = @"D:\pC#\ProjData\Statement\FBN\banner.jpg";
                    // The new Image Jira FBN-8345
                    statHtmlFBNClassic.AdImage = @"D:\pC#\ProjData\Statement\FBN\Adimage.jpg";
                    statHtmlFBNClassic.waitPeriod = 500;
                    statHtmlFBNClassic.ProductCondition = "('VISA Classic')";
                    statHtmlFBNClassic.setFrm = this;
                    checkErrRslt = statHtmlFBNClassic.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlFBNClassic = null;
                    break;

                case 322:   //69] First Bank of Nigeria  >> Emails Gold 16/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrlFBN statHtmlFBNGold = new clsStatHtmlGnrlFBN();
                    statHtmlFBNGold.emailFromName = "First Bank of Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlFBNGold.emailFrom = "cardservices@firstbanknigeria.com";//mmohammed@emp-group.com
                    statHtmlFBNGold.bankWebLink = "www.firstbanknigeria.com";
                    statHtmlFBNGold.bankLogo = @"D:\pC#\ProjData\Statement\FBN\logo.jpg";
                    statHtmlFBNGold.visaLogo = @"D:\pC#\ProjData\Statement\FBN\visalogo.jpg";
                    statHtmlFBNGold.BottomBanner = @"D:\pC#\ProjData\Statement\FBN\banner.jpg";
                    // The new Image Jira FBN-8345
                    statHtmlFBNGold.AdImage = @"D:\pC#\ProjData\Statement\FBN\Adimage.jpg";
                    statHtmlFBNGold.waitPeriod = 500;
                    statHtmlFBNGold.ProductCondition = "('VISA Gold - Standard','VISA Gold - Staff')";
                    statHtmlFBNGold.setFrm = this;
                    checkErrRslt = statHtmlFBNGold.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlFBNGold = null;
                    break;

                case 323:   //69] First Bank of Nigeria  >> Emails Infinite 16/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrlFBN statHtmlFBNInfinite = new clsStatHtmlGnrlFBN();
                    statHtmlFBNInfinite.emailFromName = "First Bank of Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlFBNInfinite.emailFrom = "cardservices@firstbanknigeria.com";//mmohammed@emp-group.com
                    statHtmlFBNInfinite.bankWebLink = "www.firstbanknigeria.com";
                    statHtmlFBNInfinite.bankLogo = @"D:\pC#\ProjData\Statement\FBN\logo.jpg";
                    statHtmlFBNInfinite.visaLogo = @"D:\pC#\ProjData\Statement\FBN\visalogo.jpg";
                    statHtmlFBNInfinite.BottomBanner = @"D:\pC#\ProjData\Statement\FBN\banner.jpg";
                    // The new Image Jira FBN-8345
                    statHtmlFBNInfinite.AdImage = @"D:\pC#\ProjData\Statement\FBN\Adimage.jpg";
                    statHtmlFBNInfinite.waitPeriod = 500;
                    statHtmlFBNInfinite.ProductCondition = "('VISA Infinite')";
                    statHtmlFBNInfinite.setFrm = this;
                    checkErrRslt = statHtmlFBNInfinite.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlFBNInfinite = null;
                    break;

                case 324:   //69] First Bank of Nigeria  >> Emails Classic Platinum 16/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrlFBN statHtmlFBNClassicPlatinum = new clsStatHtmlGnrlFBN();
                    statHtmlFBNClassicPlatinum.emailFromName = "First Bank of Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlFBNClassicPlatinum.emailFrom = "cardservices@firstbanknigeria.com";//mmohammed@emp-group.com
                    statHtmlFBNClassicPlatinum.bankWebLink = "www.firstbanknigeria.com";
                    statHtmlFBNClassicPlatinum.bankLogo = @"D:\pC#\ProjData\Statement\FBN\logo.jpg";
                    statHtmlFBNClassicPlatinum.visaLogo = @"D:\pC#\ProjData\Statement\FBN\visalogo.jpg";
                    statHtmlFBNClassicPlatinum.BottomBanner = @"D:\pC#\ProjData\Statement\FBN\banner.jpg";
                    // The new Image Jira FBN-8345
                    statHtmlFBNClassicPlatinum.AdImage = @"D:\pC#\ProjData\Statement\FBN\Adimage.jpg";
                    statHtmlFBNClassicPlatinum.waitPeriod = 500;
                    statHtmlFBNClassicPlatinum.ProductCondition = "('VISA Classic PLATINUM - NGN')";
                    statHtmlFBNClassicPlatinum.setFrm = this;
                    checkErrRslt = statHtmlFBNClassicPlatinum.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlFBNClassicPlatinum = null;
                    break;

                case 122:   //69] First Bank of Nigeria  >> Corporate Cardholder Emails 1/m
                    clsStatHtmlGnrlFBN statHtmlFBNCorp = new clsStatHtmlGnrlFBN();
                    statHtmlFBNCorp.emailFromName = "First Bank of Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlFBNCorp.emailFrom = "cardservices@firstbanknigeria.com";//mmohammed@emp-group.com
                    statHtmlFBNCorp.bankWebLink = "www.firstbanknigeria.com";
                    statHtmlFBNCorp.bankLogo = @"D:\pC#\ProjData\Statement\FBN\logo.jpg";
                    //statHtmlFBNCorp.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    //if (stmntType.Contains("MasterCard"))
                    //    statHtmlFBNCorp.visaLogo = @"D:\pC#\ProjData\Statement\FBN\MClogo.jpg";
                    //else
                    statHtmlFBNCorp.visaLogo = @"D:\pC#\ProjData\Statement\FBN\visalogo.jpg";
                    //statHtmlFBNCorp.WaterMark = @"D:\pC#\ProjData\Statement\FBN\watermark.jpg";
                    statHtmlFBNCorp.BottomBanner = @"D:\pC#\ProjData\Statement\FBN\banner.jpg";
                    // The new Image Jira FBN-8345
                    statHtmlFBNCorp.AdImage = @"D:\pC#\ProjData\Statement\FBN\Adimage.jpg";
                    //statHtmlFBNCorp.waitPeriod = 500;
                    statHtmlFBNCorp.waitPeriod = 500;
                    statHtmlFBNCorp.setFrm = this;
                    statHtmlFBNCorp.CreateCorporate = true;
                    checkErrRslt = statHtmlFBNCorp.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlFBNCorp = null;
                    break;
                //case 124:   //69] First Bank of Nigeria  >> Credit Emails 16/m
                case 147:   //69) FBN First Bank of Nigeria  >> MasterCard Credit Standard Default Emails 1/m
                ////chkDontPrompt.Checked = true;
                //clsStatHtmlGnrlFBN statHtmlFBN = new clsStatHtmlGnrlFBN();
                //statHtmlFBN.emailFromName = "First Bank of Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                //statHtmlFBN.emailFrom = "cardservices@firstbanknigeria.com";//mmohammed@emp-group.com
                //statHtmlFBN.bankWebLink = "www.firstbanknigeria.com";
                //statHtmlFBN.bankLogo = @"D:\pC#\ProjData\Statement\FBN\logo.jpg";
                //statHtmlFBN.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                //if (stmntType.Contains("MasterCard"))
                //    statHtmlFBN.visaLogo = @"D:\pC#\ProjData\Statement\FBN\MClogo.jpg";
                //else
                //statHtmlFBN.visaLogo = @"D:\pC#\ProjData\Statement\FBN\visalogo.jpg";
                //statHtmlFBN.WaterMark = @"D:\pC#\ProjData\Statement\FBN\watermark.jpg";
                //statHtmlFBN.BottomBanner = @"D:\pC#\ProjData\Statement\FBN\banner.jpg";
                ////statHtmlFBN.waitPeriod = 500;
                //statHtmlFBN.waitPeriod = 500;
                //statHtmlFBN.setFrm = this;
                //checkErrRslt = statHtmlFBN.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                //statHtmlFBN = null;
                //break;

                case 152:   //[69] FBN First Bank of Nigeria  >> Supplementary Credit Emails 16/m
                case 153:   //[69] FBN First Bank of Nigeria  >> Supplementary Credit Emails 1/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrlFBN_Sup statHtmlFBNsup = new clsStatHtmlGnrlFBN_Sup();
                    statHtmlFBNsup.emailFromName = "First Bank of Nigeria - Transaction Details";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlFBNsup.emailFrom = "cardservices@firstbanknigeria.com";//mmohammed@emp-group.com
                    statHtmlFBNsup.bankWebLink = "www.firstbanknigeria.com";
                    statHtmlFBNsup.bankLogo = @"D:\pC#\ProjData\Statement\FBN\logo.jpg";
                    statHtmlFBNsup.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    if (stmntType.Contains("MasterCard"))
                        statHtmlFBNsup.visaLogo = @"D:\pC#\ProjData\Statement\FBN\MClogo.jpg";
                    else
                        statHtmlFBNsup.visaLogo = @"D:\pC#\ProjData\Statement\FBN\visalogo.jpg";
                    //statHtmlFBNsup.WaterMark = @"D:\pC#\ProjData\Statement\FBN\watermark.jpg";
                    statHtmlFBNsup.BottomBanner = @"D:\pC#\ProjData\Statement\FBN\banner.jpg";
                    statHtmlFBNsup.ProductCondition = "('VISA Infinite','VISA Gold - Standard','VISA Gold - Staff','VISA Classic','VISA Classic PLATINUM - NGN')";
                    statHtmlFBNsup.setFrm = this;
                    //statHtmlFBNsup.waitPeriod = 500;
                    statHtmlFBNsup.waitPeriod = 500;
                    checkErrRslt = statHtmlFBNsup.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlFBNsup = null;
                    break;

                case 117:   //69] First Bank of Nigeria  >> Prepaid Emails 16/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrlFBNDebit statHtmlFBNdebit = new clsStatHtmlGnrlFBNDebit();
                    statHtmlFBNdebit.emailFromName = "First Bank of Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlFBNdebit.emailFrom = "cardservices@firstbanknigeria.com";//mmohammed@emp-group.com
                    statHtmlFBNdebit.bankWebLink = "www.firstbanknigeria.com";
                    statHtmlFBNdebit.bankLogo = @"D:\pC#\ProjData\Statement\FBN\logo.jpg";
                    //statHtmlFBN.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    statHtmlFBNdebit.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.jpg";
                    //statHtmlFBNdebit.WaterMark = @"D:\pC#\ProjData\Statement\FBN\watermark.jpg";
                    statHtmlFBNdebit.BottomBanner = @"D:\pC#\ProjData\Statement\FBN\banner.jpg";
                    //statHtmlFBNdebit.waitPeriod = 500;
                    statHtmlFBNdebit.waitPeriod = 500;
                    statHtmlFBNdebit.isPrepaid = true;
                    statHtmlFBNdebit.PrepaidCondition = "('VISA PREPAID')";
                    statHtmlFBNdebit.setFrm = this;
                    checkErrRslt = statHtmlFBNdebit.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlFBNdebit = null;
                    break;

                case 203:   //69] First Bank of Nigeria  >> MasterCard Prepaid Emails 16/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrlFBNDebit statHtmlFBNdebitmc = new clsStatHtmlGnrlFBNDebit();
                    statHtmlFBNdebitmc.emailFromName = "First Bank of Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlFBNdebitmc.emailFrom = "cardservices@firstbanknigeria.com";//mmohammed@emp-group.com
                    statHtmlFBNdebitmc.bankWebLink = "www.firstbanknigeria.com";
                    statHtmlFBNdebitmc.bankLogo = @"D:\pC#\ProjData\Statement\FBN\MClogo.jpg";
                    //statHtmlFBNdebitmc.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    statHtmlFBNdebitmc.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.jpg";
                    //statHtmlFBNdebitmc.WaterMark = @"D:\pC#\ProjData\Statement\FBN\watermark.jpg";
                    statHtmlFBNdebitmc.BottomBanner = @"D:\pC#\ProjData\Statement\FBN\banner.jpg";
                    //statHtmlFBNdebitmc.waitPeriod = 500;
                    statHtmlFBNdebitmc.waitPeriod = 500;
                    statHtmlFBNdebitmc.isPrepaid = true;
                    statHtmlFBNdebitmc.PrepaidCondition = "('PREPAID MC Unembossed')";
                    statHtmlFBNdebitmc.setFrm = this;
                    checkErrRslt = statHtmlFBNdebitmc.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlFBNdebitmc = null;
                    break;

                case 190:   // [6] AAIB Arab African International Bank >> Prepaid Email 1/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlAAIB statHtmlaaib = new clsStatHtmlAAIB();
                    statHtmlaaib.emailFromName = "Arab African International - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlaaib.emailFrom = "statement@emp-group.com";//mmohammed@emp-group.com
                    statHtmlaaib.bankWebLink = "www.aaib.com";
                    //statHtmlaaib.bankLogo = @"D:\pC#\ProjData\Statement\AAIB\logo.jpg";
                    //statHtmlaaib.visaLogo = @"D:\pC#\ProjData\Statement\AAIB\visalogo.jpg";
                    //statHtmlaaib.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\SIBN\EmailMessageMonthly.txt";
                    statHtmlaaib.HasAttachement = true;
                    statHtmlaaib.waitPeriod = 500;
                    statHtmlaaib.setFrm = this;
                    checkErrRslt = statHtmlaaib.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlaaib = null;
                    break;

                case 118:   //69] First Bank of Nigeria  >> Corporate Company Emails 1/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGnrlFBNCompany statHtmlFBNcorporate = new clsStatHtmlGnrlFBNCompany();
                    statHtmlFBNcorporate.emailFromName = "First Bank of Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlFBNcorporate.emailFrom = "cardservices@firstbanknigeria.com";//mmohammed@emp-group.com
                    statHtmlFBNcorporate.bankWebLink = "www.firstbanknigeria.com";
                    statHtmlFBNcorporate.bankLogo = @"D:\pC#\ProjData\Statement\FBN\logo.jpg";
                    //statHtmlFBN.backGround = @"D:\Web\Email\_Background\Background01.jpg";//Background06.jpg
                    statHtmlFBNcorporate.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.jpg";
                    //statHtmlFBNcorporate.WaterMark = @"D:\pC#\ProjData\Statement\FBN\watermark.jpg";
                    statHtmlFBNcorporate.BottomBanner = @"D:\pC#\ProjData\Statement\FBN\banner.jpg";
                    //statHtmlFBNcorporate.waitPeriod = 500;
                    statHtmlFBNcorporate.waitPeriod = 500;
                    statHtmlFBNcorporate.setFrm = this;
                    checkErrRslt = statHtmlFBNcorporate.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlFBNcorporate = null;
                    break;

                case 162:  //[38] GTBN GUARANTY TRUST BANK PLC NIGERIA  >> Credit Emails 16/m
                case 327:  //[38] GTBN GUARANTY TRUST BANK PLC NIGERIA  >> Corporate Cardholder Emails 16/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlGTBN statHtmlGTBN = new clsStatHtmlGTBN();
                    statHtmlGTBN.emailFromName = "GUARANTY TRUST BANK PLC Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlGTBN.emailFrom = "statement@gtbank.com";//mmohammed@emp-group.com 
                    statHtmlGTBN.bankWebLink = "www.gtbank.com";
                    statHtmlGTBN.bankLogo = @"D:\pC#\ProjData\Statement\GTBN\logo.gif";
                    //statHtmlGTBN.attachedFiles = @"D:\pC#\ProjData\Statement\GTBN\E-News";
                    //statHtmlGTBN.HasAttachement = true;
                    //statHtmlGTBN.waitPeriod = 500;
                    statHtmlGTBN.waitPeriod = 500;
                    if (pCmbProducts == 327)
                        statHtmlGTBN.CreateCorporate = true;
                    statHtmlGTBN.setFrm = this;
                    checkErrRslt = statHtmlGTBN.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlGTBN = null;
                    break;

                case 157:   //[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit  Emails 1/m [not used]
                case 159:   //[81] SIBN STANBIC IBTC BANK NIGERIA  >> Corporate Cardholder Emails 1/m
                case 208:   //[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Platinum Emails 1/m
                case 209:   //[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Gold Emails 1/m
                case 210:   //[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Infinite Emails 1/m
                case 259:   //[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Silver Emails 1/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlSIBN statHtmlSIBN = new clsStatHtmlSIBN();
                    statHtmlSIBN.emailFromName = "Stanbic IBTC Bank Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlSIBN.emailFrom = "customercarenigeria@stanbicibtc.com";//mmohammed@emp-group.com
                    statHtmlSIBN.bankWebLink = "www.stanbicibtcbank.com";
                    statHtmlSIBN.bankLogo = @"D:\pC#\ProjData\Statement\SIBN\logo.jpg";
                    statHtmlSIBN.visaLogo = @"D:\pC#\ProjData\Statement\SIBN\visalogo.jpg";
                    statHtmlSIBN.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\SIBN\EmailMessageMonthly.txt";
                    statHtmlSIBN.HasAttachement = true; //SIBN-2604
                    statHtmlSIBN.waitPeriod = 1000;
                    if (pCmbProducts == 159)
                    {
                        statHtmlSIBN.IsSplitted = false;
                        statHtmlSIBN.CreateCorporate = true;
                        statHtmlSIBN.bottomBanner = @"D:\pC#\ProjData\Statement\SIBN\corporate.jpg";

                    }
                    else if (pCmbProducts == 208)//[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Platinum Emails 1/m
                    {
                        statHtmlSIBN.bottomBanner = @"D:\pC#\ProjData\Statement\SIBN\Platinum.jpg";
                        statHtmlSIBN.IsSplitted = true;
                        statHtmlSIBN.productCond = "('Visa Platinum Credit','Visa Platinum Credit Executive','Visa Platinum Credit - USD')";
                    }
                    else if (pCmbProducts == 209)//[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Gold Emails 1/m
                    {
                        statHtmlSIBN.bottomBanner = @"D:\pC#\ProjData\Statement\SIBN\Gold.jpg";
                        statHtmlSIBN.IsSplitted = true;
                        statHtmlSIBN.productCond = "('Visa Gold Credit','Visa Gold Credit - USD')";
                    }
                    else if (pCmbProducts == 210)//[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Infinite Emails 1/m
                    {
                        statHtmlSIBN.bottomBanner = @"D:\pC#\ProjData\Statement\SIBN\Infinite.jpg";
                        statHtmlSIBN.IsSplitted = true;
                        statHtmlSIBN.productCond = "('Visa Infinite Credit','Visa Infinite Credit - USD')";
                    }
                    else if (pCmbProducts == 259)//[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Silver Emails 1/m
                    {
                        statHtmlSIBN.bottomBanner = @"D:\pC#\ProjData\Statement\SIBN\Gold.jpg";
                        statHtmlSIBN.IsSplitted = true;
                        statHtmlSIBN.productCond = "('Visa Silver Credit')";
                    }
                    statHtmlSIBN.setFrm = this;
                    checkErrRslt = statHtmlSIBN.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlSIBN = null;
                    break;

                //case 208:   //[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Platinum Emails 1/m
                //    clsStatHtmlSIBN statHtmlSIBNPlatinum = new clsStatHtmlSIBN();
                //    statHtmlSIBNPlatinum.emailFromName = "Stanbic IBTC Bank Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                //    statHtmlSIBNPlatinum.emailFrom = "customercarenigeria@stanbicibtc.com";//mmohammed@emp-group.com
                //    statHtmlSIBNPlatinum.bankWebLink = "www.stanbicibtcbank.com";
                //    statHtmlSIBNPlatinum.bankLogo = @"D:\pC#\ProjData\Statement\SIBN\logo.jpg";
                //    statHtmlSIBNPlatinum.visaLogo = @"D:\pC#\ProjData\Statement\SIBN\visalogo.jpg";
                //    statHtmlSIBNPlatinum.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\SIBN\EmailMessageMonthly.txt";
                //    statHtmlSIBNPlatinum.IsSplitted = true;
                //    statHtmlSIBNPlatinum.productCond = "('Visa Platinum Credit','Visa Platinum Credit Executive','Visa Platinum Credit - USD')";
                //    statHtmlSIBNPlatinum.bottomBanner = @"D:\pC#\ProjData\Statement\SIBN\Platinum.jpg";
                //    statHtmlSIBNPlatinum.waitPeriod = 1000;
                //    statHtmlSIBNPlatinum.HasAttachement = true; //SIBN-2604
                //    statHtmlSIBNPlatinum.setFrm = this;
                //    checkErrRslt = statHtmlSIBNPlatinum.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                //    statHtmlSIBNPlatinum = null;
                //    break;

                //case 209:   //[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Gold Emails 1/m
                //    clsStatHtmlSIBN statHtmlSIBNGold = new clsStatHtmlSIBN();
                //    statHtmlSIBNGold.emailFromName = "Stanbic IBTC Bank Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                //    statHtmlSIBNGold.emailFrom = "customercarenigeria@stanbicibtc.com";//mmohammed@emp-group.com
                //    statHtmlSIBNGold.bankWebLink = "www.stanbicibtcbank.com";
                //    statHtmlSIBNGold.bankLogo = @"D:\pC#\ProjData\Statement\SIBN\logo.jpg";
                //    statHtmlSIBNGold.visaLogo = @"D:\pC#\ProjData\Statement\SIBN\visalogo.jpg";
                //    statHtmlSIBNGold.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\SIBN\EmailMessageMonthly.txt";
                //    statHtmlSIBNGold.IsSplitted = true;
                //    statHtmlSIBNGold.productCond = "('Visa Gold Credit','Visa Gold Credit - USD')";
                //    statHtmlSIBNGold.bottomBanner = @"D:\pC#\ProjData\Statement\SIBN\Gold.jpg";
                //    statHtmlSIBNGold.waitPeriod = 1000;
                //    statHtmlSIBNGold.HasAttachement = true; //SIBN-2604
                //    statHtmlSIBNGold.setFrm = this;
                //    checkErrRslt = statHtmlSIBNGold.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                //    statHtmlSIBNGold = null;
                //    break;

                //case 210:   //[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Infinite Emails 1/m
                //    clsStatHtmlSIBN statHtmlSIBNInfinite = new clsStatHtmlSIBN();
                //    statHtmlSIBNInfinite.emailFromName = "Stanbic IBTC Bank Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                //    statHtmlSIBNInfinite.emailFrom = "customercarenigeria@stanbicibtc.com";//mmohammed@emp-group.com
                //    statHtmlSIBNInfinite.bankWebLink = "www.stanbicibtcbank.com";
                //    statHtmlSIBNInfinite.bankLogo = @"D:\pC#\ProjData\Statement\SIBN\logo.jpg";
                //    statHtmlSIBNInfinite.visaLogo = @"D:\pC#\ProjData\Statement\SIBN\visalogo.jpg";
                //    statHtmlSIBNInfinite.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\SIBN\EmailMessageMonthly.txt";
                //    statHtmlSIBNInfinite.IsSplitted = true;
                //    statHtmlSIBNInfinite.productCond = "('Visa Infinite Credit','Visa Infinite Credit - USD')";
                //    statHtmlSIBNInfinite.bottomBanner = @"D:\pC#\ProjData\Statement\SIBN\Infinite.jpg";
                //    statHtmlSIBNInfinite.waitPeriod = 1000;
                //    statHtmlSIBNInfinite.HasAttachement = true; //SIBN-2604
                //    statHtmlSIBNInfinite.setFrm = this;
                //    checkErrRslt = statHtmlSIBNInfinite.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                //    statHtmlSIBNInfinite = null;
                //    break;

                //case 259:   //[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Silver Emails 1/m
                //    clsStatHtmlSIBN statHtmlSIBNSilver = new clsStatHtmlSIBN();
                //    statHtmlSIBNSilver.emailFromName = "Stanbic IBTC Bank Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                //    statHtmlSIBNSilver.emailFrom = "customercarenigeria@stanbicibtc.com";//mmohammed@emp-group.com
                //    statHtmlSIBNSilver.bankWebLink = "www.stanbicibtcbank.com";
                //    statHtmlSIBNSilver.bankLogo = @"D:\pC#\ProjData\Statement\SIBN\logo.jpg";
                //    statHtmlSIBNSilver.visaLogo = @"D:\pC#\ProjData\Statement\SIBN\visalogo.jpg";
                //    statHtmlSIBNSilver.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\SIBN\EmailMessageMonthly.txt";
                //    statHtmlSIBNSilver.IsSplitted = true;
                //    statHtmlSIBNSilver.productCond = "('Visa Silver Credit')";
                //    statHtmlSIBNSilver.bottomBanner = @"D:\pC#\ProjData\Statement\SIBN\Gold.jpg";
                //    statHtmlSIBNSilver.waitPeriod = 1000;
                //    statHtmlSIBNSilver.HasAttachement = true; //SIBN-2604
                //    statHtmlSIBNSilver.setFrm = this;
                //    checkErrRslt = statHtmlSIBNSilver.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                //    statHtmlSIBNSilver = null;
                //    break;

                case 87:   // 50) ICBG Intercontinental Bank Ghana Limited Debit >> PDF 1/m
                    //clsStatTxtLblDb statTxtLblDb = new clsStatTxtLblDb();
                    //statTxtLblDb.setFrm = this;
                    ////if (pCmbProducts == 100)
                    ////    statTxtLblDb.emailService = true;
                    //checkErrRslt = statTxtLblDb.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    //statTxtLblDb = null;
                    //break;
                    //clsMaintainData maintainDataicbgdb = new clsMaintainData();
                    //maintainDataicbgdb.fixAddress(bankCode);
                    //maintainDataicbgdb = null;
                    clsStatTxtLblDb_ICBG statTxtLblDb_Debit = new clsStatTxtLblDb_ICBG();
                    statTxtLblDb_Debit.setFrm = this;
                    //statTxtLblDb_Debit.PrepaidCondition = "('Visa Classic Prepaid','Visa Classic Prepaid Gift')";
                    statTxtLblDb_Debit.PrepaidCondition = "('Visa Electron (Debit)','Visa Classic (Debit)')";
                    checkErrRslt = statTxtLblDb_Debit.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblDb_Debit = null;
                    break;

                case 52: //case 52:   // 22) BPC BANCO DE POUPANCA E CREDITO, SARL Corporate >> Default PDF 1/m
                    //clsMaintainData maintainDatabpc = new clsMaintainData();
                    //maintainDatabpc.fixAddress(bankCode);
                    //maintainDatabpc = null;
                    clsStatement_CommonCorpExp stmntCorporateBPC = new clsStatement_CommonCorpExp();
                    stmntCorporateBPC.mantainBank(bankCode);
                    stmntCorporateBPC.CreateCorporate = true;
                    stmntCorporateBPC.reportCompany = Application.StartupPath + @"\Reports\Statement_BPC_Corporate_Company.rpt";
                    stmntCorporateBPC.reportIndividual = Application.StartupPath + @"\Reports\Statement_BPC_Portuguese.rpt";
                    stmntCorporateBPC.export(txtFileName.Text, strStatementType, bankCode, strFileName, Application.StartupPath + @"\Reports\Statement_Common_Corporate_Company.rpt", ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                    stmntCorporateBPC.CreateZip();
                    stmntCorporateBPC = null;
                    break;

                case 120: //47) SBN Sterling Bank Nigeria >> Default Debit Text 1/m
                    //clsMaintainData maintainDatasbndb = new clsMaintainData();
                    //maintainDatasbndb.fixAddress(bankCode);
                    //maintainDatasbndb = null;
                    clsStatTxtLblDb_ICBG statTxtLblDb_Prepaid = new clsStatTxtLblDb_ICBG();
                    statTxtLblDb_Prepaid.setFrm = this;
                    statTxtLblDb_Prepaid.PrepaidCondition = "('Visa Classic Debit')";
                    //statTxtLblDb_Prepaid.PrepaidCondition = "('Visa Gold Debit')";
                    checkErrRslt = statTxtLblDb_Prepaid.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblDb_Prepaid = null;
                    break;

                case 248: //44) BOAK Bank of Africa Kenya >> Default Debit Text 1/m
                    //clsMaintainData maintainDatasbndb = new clsMaintainData();
                    //maintainDatasbndb.fixAddress(bankCode);
                    //maintainDatasbndb = null;
                    clsStatTxtLblDb_ICBG statTxtLblDb_Prepaid_BOAK = new clsStatTxtLblDb_ICBG();
                    statTxtLblDb_Prepaid_BOAK.setFrm = this;
                    statTxtLblDb_Prepaid_BOAK.PrepaidCondition = "('Visa Gold Differed Debit KES','Visa Gold Differed Debit USD')";
                    checkErrRslt = statTxtLblDb_Prepaid_BOAK.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblDb_Prepaid_BOAK = null;
                    break;

                case 254: //113) GTBU Guaranty trust bank Uganda  >> Debit Text 1/m
                    //clsMaintainData maintainDatasbndb = new clsMaintainData();
                    //maintainDatasbndb.fixAddress(bankCode);
                    //maintainDatasbndb = null;
                    clsStatTxtLblDb_ICBG statTxtLblDb_Prepaid_GTBU = new clsStatTxtLblDb_ICBG();
                    statTxtLblDb_Prepaid_GTBU.setFrm = this;
                    statTxtLblDb_Prepaid_GTBU.PrepaidCondition = "('MasterCard Prepaid General Spend Card')";
                    checkErrRslt = statTxtLblDb_Prepaid_GTBU.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblDb_Prepaid_GTBU = null;
                    break;

                case 294: //129) SBL SAHARA BANK BNP PARIBAS  >> Prepaid Text 1/m
                    //clsMaintainData maintainDatasbndb = new clsMaintainData();
                    //maintainDatasbndb.fixAddress(bankCode);
                    //maintainDatasbndb = null;
                    clsStatTxtLblDb_ICBG statTxtLblDb_Prepaid_SBL = new clsStatTxtLblDb_ICBG();
                    statTxtLblDb_Prepaid_SBL.setFrm = this;
                    statTxtLblDb_Prepaid_SBL.PrepaidCondition = "('VISA Prepaid RAFEEQ','VISA Prepaid SHER')";
                    checkErrRslt = statTxtLblDb_Prepaid_SBL.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblDb_Prepaid_SBL = null;
                    break;

                case 132: // 72) ABPR ACCESS BANK RWANDA SA >> Default Debit Text 1/m
                    //clsMaintainData maintainDatasbndb = new clsMaintainData();
                    //maintainDatasbndb.fixAddress(bankCode);
                    //maintainDatasbndb = null;
                    clsStatTxtLblDb_ICBG statTxtLblDb_Prepaid_ABPR = new clsStatTxtLblDb_ICBG();
                    statTxtLblDb_Prepaid_ABPR.setFrm = this;
                    statTxtLblDb_Prepaid_ABPR.PrepaidCondition = "('Visa Classic Debit')";
                    checkErrRslt = statTxtLblDb_Prepaid_ABPR.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblDb_Prepaid_ABPR = null;
                    break;

                case 239: //109) CGBK COMPAGNIE GENERALE DE BANQUE LTD >> MasterCard Prepaid Text 1/m
                    clsStatTxtLblDb_ICBG statTxtLblDb_MCPrepaid = new clsStatTxtLblDb_ICBG();
                    statTxtLblDb_MCPrepaid.setFrm = this;
                    statTxtLblDb_MCPrepaid.PrepaidCondition = "('MasterCard Prepaid')";
                    checkErrRslt = statTxtLblDb_MCPrepaid.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblDb_MCPrepaid = null;
                    break;

                case 275: //33) ZENG Zenith Bank (Ghana) Limited >> MasterCard Prepaid text 16/m
                    clsStatTxtLblDb_ICBG statTxtLblDb_ZENGMCPrepaid = new clsStatTxtLblDb_ICBG();
                    statTxtLblDb_ZENGMCPrepaid.setFrm = this;
                    statTxtLblDb_ZENGMCPrepaid.PrepaidCondition = "('MasterCard Prepaid')";
                    checkErrRslt = statTxtLblDb_ZENGMCPrepaid.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblDb_ZENGMCPrepaid = null;
                    break;

                case 175: //58) SBP SKYE BANK PLC >> Default Prepaid Text 1/m
                    //clsMaintainData maintainDatasbpdb = new clsMaintainData();
                    //maintainDatasbpdb.fixAddress(bankCode);
                    //maintainDatasbpdb = null;
                    clsStatTxtLblDb_ICBG statTxtLblSBP_Prepaid = new clsStatTxtLblDb_ICBG();
                    statTxtLblSBP_Prepaid.setFrm = this;
                    statTxtLblSBP_Prepaid.PrepaidCondition = "('Visa Prepaid Travel Allowance Card','MasterCard Prepaid','Visa Prepaid-NTDC')";
                    checkErrRslt = statTxtLblSBP_Prepaid.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLblSBP_Prepaid = null;
                    break;

                case 143:  //74) BLME Blom Bank Egypt VISA>>  Credit Text 1/m
                case 141:  //74) BLME Blom Bank Egypt MasterCard >> Credit Text 1/m
                    //clsMaintainData maintainDatablme = new clsMaintainData();
                    //maintainDatablme.fixAddress(bankCode);
                    //maintainDatablme = null;
                    clsStatementBLMEcredit stmntBLMEcredit = new clsStatementBLMEcredit();
                    stmntBLMEcredit.setFrm = this;
                    stmntBLMEcredit.savedDataset = true;
                    stmntBLMEcredit.whereCond = whereCond;
                    if (pCmbProducts == 143)
                        //stmntBLMEcredit.productCond = "('Visa Classic','Visa Gold','Visa Platinum','Visa Classic with 1.5% Interest','Visa Gold with 1.5% Interest','Visa Platinum with 1.5 interest')";
                        stmntBLMEcredit.productCond = "'Visa%'";
                    if (pCmbProducts == 141)
                        //stmntBLMEcredit.productCond = "('MasterCard Classic','MasterCard Platinum','MasterCard Titanium','MasterCard Gold','MasterCard Classic with 1.5 interest','MasterCard Platinum with 1.5 interest','MasterCard Titanium with 1.5 interest','MasterCard Gold with 1.5 interest')";
                        stmntBLMEcredit.productCond = "'MasterCard%'";
                    checkErrRslt = stmntBLMEcredit.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);
                    stmntBLMEcredit = null;
                    break;

                //case 143:  //74) BLME Blom Bank Egypt VISA>>  Credit Text 1/m
                //    clsStatementinsBLME stmntinsblme = new clsStatementinsBLME();
                //    stmntinsblme.setFrm = this;
                //    stmntinsblme.savedDataset = true;
                //    stmntinsblme.statMessageBox = @"";

                //    stmntinsblme.productCond = "('Visa Classic')";
                //    stmntinsblme.InstallmentCondition = "('Visa Classic')";
                //    stmntinsblme.isInstallmentVal = true;

                //    checkErrRslt = stmntinsblme.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);
                //    stmntinsblme = null;
                //    break;

                case 144:   // [22] BPC BANCO DE POUPANCA E CREDITO, SARL Credit >> Email 16/m
                case 186:   // [22] BPC BANCO DE POUPANCA E CREDITO, SARL Credit >> Cardholder Email 1/m
                    clsStatHtmlBPC statHtmlBPC = new clsStatHtmlBPC();
                    statHtmlBPC.emailFromName = "Extracto Cartão BPC";//"BAI - Statement""mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBPC.emailFrom = "CartoesdeCredito@bpc.ao";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlBPC.bankWebLink = "www.bpc.ao";//"www.emp-group.com""www.socgen.com"
                    statHtmlBPC.bankWebLinkService = "www.bpc.ao.";//"www.socgen.com"
                    statHtmlBPC.bankLogo = @"D:\pC#\ProjData\Statement\BPC\Logo.jpg";
                    statHtmlBPC.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\BPC\EmailMessageMonthly.txt";
                    //statHtmlBPC.waitPeriod = 500;
                    statHtmlBPC.waitPeriod = 500;
                    statHtmlBPC.setFrm = this;
                    if (pCmbProducts == 186)
                        statHtmlBPC.CreateCorporate = true;
                    checkErrRslt = statHtmlBPC.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    statHtmlBPC = null;
                    break;

                case 145:   // 31) NBS NBS Bank Limited >> Email 16/m
                    clsStatHtmlNBS statHtmlNBS = new clsStatHtmlNBS();
                    statHtmlNBS.emailFromName = "NBS BANK Limited - Statement";//"BAI - Statement""mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    //statHtmlNBS.emailFromName = "NBS BANK PLC - Statement";//"BAI - Statement""mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlNBS.emailFrom = "visacardstmt@nbsmw.com";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    //statHtmlNBS.emailFrom = "mabouleila@emp-group.com";
                    //statHtmlNBS.bankWebLink = "www.nbs.mw";//"www.emp-group.com""www.socgen.com"
                    statHtmlNBS.bankWebLink = "www.nbsmw.com";//"www.emp-group.com""www.socgen.com"
                    statHtmlNBS.bankLogo = @"D:\pC#\ProjData\Statement\NBS\logo.gif";
                    statHtmlNBS.visaLogo = @"D:\pC#\ProjData\Statement\_Background\VisaLogo.gif";
                    statHtmlNBS.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";//Background06.jpg
                    statHtmlNBS.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NBS\EmailMessageMonthly.txt";
                    statHtmlNBS.logoAlignment = "left";
                    //statHtmlNBS.waitPeriod = 500;
                    statHtmlNBS.setFrm = this;
                    statHtmlNBS.waitPeriod = 3000;
                    statHtmlNBS.HasAttachement = true;
                    statHtmlNBS.IsSplitted = false;
                    checkErrRslt = statHtmlNBS.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlNBS = null;
                    break;

                case 146: // 76) EDBE EXPORT DEVELOPMENT BANK OF EGYPT >> Default Credit Text 1/m
                    //clsMaintainData maintainDataedbe = new clsMaintainData();
                    //maintainDataedbe.fixAddress(bankCode);
                    //maintainDataedbe = null;
                    clsStatTxt_EDBE statTxtLbledbe = new clsStatTxt_EDBE();// + "ABP_Statement_File.txt"
                    statTxtLbledbe.setFrm = this;

                    statTxtLbledbe.isRewardVal = true;
                    statTxtLbledbe.rewardCondition = "('EBE Loyalty Program')";
                    //statTxtLbledbe.ExcludeCond = "";

                    checkErrRslt = statTxtLbledbe.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    statTxtLbledbe = null;
                    break;

                case 42:   //33) ZENG Zenith Bank (Ghana) Limited >> Visa Classic Credit PDF 16/m
                case 281:   //33) ZENG Zenith Bank (Ghana) Limited >> Visa Classic Credit (Staff) PDF 16/m
                case 282:   //33) ZENG Zenith Bank (Ghana) Limited >> MasterCard Credit PDF 16/m
                case 283:   //33) ZENG Zenith Bank (Ghana) Limited >> MasterCard Credit (Staff) PDF 16/m
                case 442:   //33) ZENG Zenith Bank (Ghana) Limited >> MasterCard Prepaid PDF 16/m
                    //clsMaintainData maintainDatazeng = new clsMaintainData();
                    //maintainDatazeng.fixAddress(bankCode);
                    //maintainDatazeng = null;
                    clsStatement_ExportRpt stmntZENG = new clsStatement_ExportRpt();
                    stmntZENG.setFrm = this;
                    stmntZENG.mantainBank(bankCode);
                    stmntZENG.whereCond = whereCond;
                    if (pCmbProducts == 42)
                        stmntZENG.productCond = "'Visa Classic Credit'";
                    if (pCmbProducts == 281)
                        stmntZENG.productCond = "'Visa Classic Credit (Staff)'";
                    if (pCmbProducts == 282)
                        stmntZENG.productCond = "'MasterCard Credit'";
                    if (pCmbProducts == 283)
                        stmntZENG.productCond = "'MasterCard Credit (Staff)'";
                    if (pCmbProducts == 442) // List Prepaid Product , brackets and parentheses in query, [] for crystal formula () for sql summary query
                        stmntZENG.productCond = "'MasterCard Prepaid','MasterCard Platinum Prepaid'," +
                            "'Prepaid MasterCard Business Card','Business Debit MasterCard'," +
                            "'Visa Classic Prepaid Travel Card'";

                    stmntZENG.SplitByBranchByProfile(txtFileName.Text, strStatementType, bankCode, strFileName, reportFleName, ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                    stmntZENG.CreateZip();
                    stmntZENG = null;
                    break;

                case 56:   // 1) NSGB Credit >> Business Individual Email 1/m
                    clsStatementNSGBcreditHtml emailStmntNSGBIndividual = new clsStatementNSGBcreditHtml();
                    emailStmntNSGBIndividual.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    //emailStmntNSGBIndividual.emailFrom = "Card.Support@QNBALAHLI.COM";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBIndividual.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBIndividual.bankWebLink = "qnbalahli.com";//"www.emp-group.com""www.socgen.com"
                    //emailStmntNSGBIndividual.bankWebLinkService = " Card.Support@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBIndividual.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBIndividual.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBIndividual.logoAlignment = "left";
                    //emailStmntNSGBIndividual.statMessageFile = @"D:\pC#\ProjData\Statement\NSGB\Individual\EmailMessage.txt";
                    emailStmntNSGBIndividual.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\Individual\Individual.jpg";
                    emailStmntNSGBIndividual.IsSplitted = true;
                    emailStmntNSGBIndividual.productCond = "Visa Business - Individuals";
                    emailStmntNSGBIndividual.HasAttachement = true;
                    emailStmntNSGBIndividual.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\Individual\E-News";
                    emailStmntNSGBIndividual.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\Individual\Individual.txt";
                    //NSGB-3615
                    //emailStmntNSGBIndividual.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.gif";
                    emailStmntNSGBIndividual.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.jpg";
                    emailStmntNSGBIndividual.waitPeriod = 3000;
                    emailStmntNSGBIndividual.setFrm = this;
                    //emailStmntNSGBIndividual.isReward = true;
                    //emailStmntNSGBIndividual.rewardCond = "'Reward Program (T)'";
                    checkErrRslt = emailStmntNSGBIndividual.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntNSGBIndividual = null;
                    break;

                case 169:   // 1) NSGB Credit >> Classic Email 1/m
                    clsStatementNSGBcreditHtml emailStmntNSGBClassic = new clsStatementNSGBcreditHtml();
                    emailStmntNSGBClassic.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    //to be changed to any fake domain in case of testing and sampling
                    //emailStmntNSGBClassic.emailFrom = "Card.Support@QNBALAHLI.COM";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBClassic.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBClassic.bankWebLink = "qnbalahli.com";//"www.emp-group.com""www.socgen.com"
                    //emailStmntNSGBClassic.bankWebLinkService = " Card.Support@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBClassic.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBClassic.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBClassic.logoAlignment = "left";
                    //emailStmntNSGBClassic.statMessageFile = @"D:\pC#\ProjData\Statement\NSGB\Classic\EmailMessage.txt";
                    emailStmntNSGBClassic.statMessageBox = @"D:\pC#\ProjData\Statement\NSGB\EmailMessageBox.txt";
                    emailStmntNSGBClassic.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\Classic\Classic.jpg";
                    emailStmntNSGBClassic.IsSplitted = true;
                    emailStmntNSGBClassic.productCond = "Visa Classic";
                    emailStmntNSGBClassic.HasAttachement = true;
                    emailStmntNSGBClassic.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\Classic\E-News";
                    emailStmntNSGBClassic.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\Classic\Classic.txt";
                    //NSGB-3615
                    //emailStmntNSGBClassic.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.gif";
                    //emailStmntNSGBClassic.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\MCLogo.jpg";
                    //NSGB-3945
                    emailStmntNSGBClassic.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.jpg";
                    emailStmntNSGBClassic.waitPeriod = 3000;
                    emailStmntNSGBClassic.setFrm = this;
                    emailStmntNSGBClassic.isReward = true;
                    emailStmntNSGBClassic.rewardCond = "('Reward Program - VISA Classic')";
                    checkErrRslt = emailStmntNSGBClassic.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntNSGBClassic = null;
                    break;

                case 170:   // 1) NSGB Credit >> Gold Email 1/m
                    clsStatementNSGBcreditHtml emailStmntNSGBGold = new clsStatementNSGBcreditHtml();
                    emailStmntNSGBGold.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    //emailStmntNSGBGold.emailFrom = "Card.Support@QNBALAHLI.COM";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBGold.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBGold.bankWebLink = "qnbalahli.com";//"www.emp-group.com""www.socgen.com"
                    //emailStmntNSGBGold.bankWebLinkService = " Card.Support@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBGold.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBGold.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBGold.logoAlignment = "left";
                    //emailStmntNSGBGold.statMessageFile = @"D:\pC#\ProjData\Statement\NSGB\Gold\EmailMessage.txt";
                    emailStmntNSGBGold.statMessageBox = @"D:\pC#\ProjData\Statement\NSGB\EmailMessageBox.txt";
                    emailStmntNSGBGold.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\Gold\Gold.jpg";
                    emailStmntNSGBGold.IsSplitted = true;
                    emailStmntNSGBGold.productCond = "Visa Gold";
                    emailStmntNSGBGold.HasAttachement = true;
                    emailStmntNSGBGold.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\Gold\E-News";
                    emailStmntNSGBGold.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\Gold\Gold.txt";
                    //NSGB-3615
                    //emailStmntNSGBGold.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.gif";
                    emailStmntNSGBGold.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.jpg";
                    emailStmntNSGBGold.waitPeriod = 3000;
                    emailStmntNSGBGold.setFrm = this;
                    emailStmntNSGBGold.isReward = true;
                    emailStmntNSGBGold.rewardCond = "('Reward Program - Visa Gold')";
                    checkErrRslt = emailStmntNSGBGold.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntNSGBGold = null;
                    break;

                case 171:   // 1) NSGB Credit >> Platinum Email 1/m
                    clsStatementNSGBcreditHtml emailStmntNSGBPlatinum = new clsStatementNSGBcreditHtml();
                    emailStmntNSGBPlatinum.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    //emailStmntNSGBPlatinum.emailFrom = "Card.Support@QNBALAHLI.COM";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBPlatinum.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBPlatinum.bankWebLink = "qnbalahli.com";//"www.emp-group.com""www.socgen.com"
                    //emailStmntNSGBPlatinum.bankWebLinkService = " Card.Support@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBPlatinum.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBPlatinum.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBPlatinum.logoAlignment = "left";
                    //emailStmntNSGBPlatinum.statMessageFile = @"D:\pC#\ProjData\Statement\NSGB\Platinum\EmailMessage.txt";
                    emailStmntNSGBPlatinum.statMessageBox = @"D:\pC#\ProjData\Statement\NSGB\EmailMessageBox.txt";
                    emailStmntNSGBPlatinum.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\Platinum\Platinum.jpg";
                    emailStmntNSGBPlatinum.IsSplitted = true;
                    emailStmntNSGBPlatinum.productCond = "Visa Platinum";
                    emailStmntNSGBPlatinum.HasAttachement = true;
                    emailStmntNSGBPlatinum.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\Platinum\E-News";
                    emailStmntNSGBPlatinum.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\Platinum\Platinum.txt";
                    //NSGB-3615
                    //emailStmntNSGBPlatinum.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.gif";
                    emailStmntNSGBPlatinum.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.jpg";
                    emailStmntNSGBPlatinum.waitPeriod = 3000;
                    emailStmntNSGBPlatinum.setFrm = this;
                    emailStmntNSGBPlatinum.isReward = true;
                    emailStmntNSGBPlatinum.rewardCond = "('Reward Program - Visa Platinum')";
                    checkErrRslt = emailStmntNSGBPlatinum.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntNSGBPlatinum = null;
                    break;

                case 404:   // 1) NSGB Credit >> Platinum PDF Email 1/m
                    clsStatHtmlQNBPDF emailStmntNSGBPlatinumPDF = new clsStatHtmlQNBPDF();
                    emailStmntNSGBPlatinumPDF.isSplitted = true;
                    //statHtmlunbn.PaymentSystem = "Visa";
                    emailStmntNSGBPlatinumPDF.BillingCycle = "EOM";
                    emailStmntNSGBPlatinumPDF.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBPlatinumPDF.emailFrom = "noreply.e-statement@qnbalahli.com";//mmohammed@emp-group.com
                    emailStmntNSGBPlatinumPDF.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBPlatinumPDF.bankWebLink = "qnbalahli.com";
                    emailStmntNSGBPlatinumPDF.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    //emailStmntNSGBPlatinumPDF.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    //emailStmntNSGBPlatinumPDF.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.gif";
                    //emailStmntNSGBPlatinumPDF.bottomBanner = @"D:\pC#\ProjData\Statement\UNBN\bottom.jpg";
                    emailStmntNSGBPlatinumPDF.waitPeriod = 3000;
                    emailStmntNSGBPlatinumPDF.HasAttachement = true;
                    emailStmntNSGBPlatinumPDF.isReward = true;
                    emailStmntNSGBPlatinumPDF.rewardCond = "('Reward Program - Visa Platinum')";

                    emailStmntNSGBPlatinumPDF.IsCardProduct = true;
                    emailStmntNSGBPlatinumPDF.productCond = "('Visa Platinum')";
                    emailStmntNSGBPlatinumPDF.IsContractType = false;
                    emailStmntNSGBPlatinumPDF.ContractTypeCond = "('VISA Platinum - Staff')";

                    emailStmntNSGBPlatinumPDF.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\Platinum\Platinum.jpg";
                    emailStmntNSGBPlatinumPDF.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\Platinum\E-News";
                    emailStmntNSGBPlatinumPDF.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\Platinum\Platinum.txt";


                    emailStmntNSGBPlatinumPDF.setFrm = this;
                    checkErrRslt = emailStmntNSGBPlatinumPDF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntNSGBPlatinumPDF = null;
                    break;

                case 420:   // 1) NSGB Credit >> Classic PDF Email 1/m
                    clsStatHtmlQNBPDF emailStmntNSGBClassicPDF = new clsStatHtmlQNBPDF();
                    emailStmntNSGBClassicPDF.isSplitted = true;
                    //statHtmlunbn.PaymentSystem = "Visa";
                    emailStmntNSGBClassicPDF.BillingCycle = "EOM";
                    emailStmntNSGBClassicPDF.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBClassicPDF.emailFrom = "noreply.e-statement@qnbalahli.com";//mmohammed@emp-group.com
                    emailStmntNSGBClassicPDF.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBClassicPDF.bankWebLink = "qnbalahli.com";
                    emailStmntNSGBClassicPDF.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    //emailStmntNSGBClassicPDF.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    //emailStmntNSGBClassicPDF.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.gif";
                    //emailStmntNSGBClassicPDF.bottomBanner = @"D:\pC#\ProjData\Statement\UNBN\bottom.jpg";
                    emailStmntNSGBClassicPDF.waitPeriod = 3000;
                    emailStmntNSGBClassicPDF.HasAttachement = true;
                    emailStmntNSGBClassicPDF.isReward = true;
                    emailStmntNSGBClassicPDF.rewardCond = "('Reward Program - VISA Classic')";

                    emailStmntNSGBClassicPDF.IsCardProduct = true;
                    emailStmntNSGBClassicPDF.productCond = "('Visa Classic')";
                    emailStmntNSGBClassicPDF.IsContractType = false;
                    emailStmntNSGBClassicPDF.ContractTypeCond = "('Visa Classic - Staff')";

                    emailStmntNSGBClassicPDF.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\Classic\Classic.jpg";
                    emailStmntNSGBClassicPDF.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\Classic\E-News";
                    emailStmntNSGBClassicPDF.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\Classic\Classic.txt";


                    emailStmntNSGBClassicPDF.setFrm = this;
                    checkErrRslt = emailStmntNSGBClassicPDF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntNSGBClassicPDF = null;
                    break;

                case 421:   // 1) NSGB Credit >> Gold PDF Email 1/m
                    clsStatHtmlQNBPDF emailStmntNSGBGoldPDF = new clsStatHtmlQNBPDF();
                    emailStmntNSGBGoldPDF.isSplitted = true;
                    //statHtmlunbn.PaymentSystem = "Visa";
                    emailStmntNSGBGoldPDF.BillingCycle = "EOM";
                    emailStmntNSGBGoldPDF.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBGoldPDF.emailFrom = "noreply.e-statement@qnbalahli.com";//mmohammed@emp-group.com
                    emailStmntNSGBGoldPDF.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBGoldPDF.bankWebLink = "qnbalahli.com";
                    emailStmntNSGBGoldPDF.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    //emailStmntNSGBGoldPDF.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    //emailStmntNSGBGoldPDF.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.gif";
                    //emailStmntNSGBGoldPDF.bottomBanner = @"D:\pC#\ProjData\Statement\UNBN\bottom.jpg";
                    emailStmntNSGBGoldPDF.waitPeriod = 3000;
                    emailStmntNSGBGoldPDF.HasAttachement = true;
                    emailStmntNSGBGoldPDF.isReward = true;
                    emailStmntNSGBGoldPDF.rewardCond = "('Reward Program - Visa Gold')";

                    emailStmntNSGBGoldPDF.IsCardProduct = true;
                    emailStmntNSGBGoldPDF.productCond = "('Visa Gold')";
                    emailStmntNSGBGoldPDF.IsContractType = false;
                    emailStmntNSGBGoldPDF.ContractTypeCond = "('Visa Gold - Staff')";

                    emailStmntNSGBGoldPDF.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\Gold\Gold.jpg";
                    emailStmntNSGBGoldPDF.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\Gold\E-News";
                    emailStmntNSGBGoldPDF.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\Gold\Gold.txt";


                    emailStmntNSGBGoldPDF.setFrm = this;
                    checkErrRslt = emailStmntNSGBGoldPDF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntNSGBGoldPDF = null;
                    break;

                case 422:   // 1) NSGB Credit >> Infinite PDF Email 1/m
                    clsStatHtmlQNBPDF emailStmntNSGBInfinitePDF = new clsStatHtmlQNBPDF();
                    emailStmntNSGBInfinitePDF.isSplitted = true;
                    //statHtmlunbn.PaymentSystem = "Visa";
                    emailStmntNSGBInfinitePDF.BillingCycle = "EOM";
                    emailStmntNSGBInfinitePDF.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBInfinitePDF.emailFrom = "noreply.e-statement@qnbalahli.com";//mmohammed@emp-group.com
                    emailStmntNSGBInfinitePDF.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBInfinitePDF.bankWebLink = "qnbalahli.com";
                    emailStmntNSGBInfinitePDF.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    //emailStmntNSGBInfinitePDF.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    //emailStmntNSGBInfinitePDF.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.gif";
                    //emailStmntNSGBInfinitePDF.bottomBanner = @"D:\pC#\ProjData\Statement\UNBN\bottom.jpg";
                    emailStmntNSGBInfinitePDF.waitPeriod = 3000;
                    emailStmntNSGBInfinitePDF.HasAttachement = true;
                    emailStmntNSGBInfinitePDF.isReward = true;
                    emailStmntNSGBInfinitePDF.rewardCond = "('Reward Program - Visa infinite')";

                    emailStmntNSGBInfinitePDF.IsCardProduct = true;
                    emailStmntNSGBInfinitePDF.productCond = "('Visa infinite')";
                    emailStmntNSGBInfinitePDF.IsContractType = false;
                    emailStmntNSGBInfinitePDF.ContractTypeCond = "('Visa infinite - Staff')";

                    emailStmntNSGBInfinitePDF.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\Infinite\Infinite.jpg";
                    emailStmntNSGBInfinitePDF.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\Infinite\E-News";
                    emailStmntNSGBInfinitePDF.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\Infinite\Infinite.txt";


                    emailStmntNSGBInfinitePDF.setFrm = this;
                    checkErrRslt = emailStmntNSGBInfinitePDF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntNSGBInfinitePDF = null;
                    break;

                case 423:   // 1) NSGB Credit >> Individual PDF Email 1/m
                    clsStatHtmlQNBPDF emailStmntNSGBIndividualPDF = new clsStatHtmlQNBPDF();
                    emailStmntNSGBIndividualPDF.isSplitted = true;
                    //statHtmlunbn.PaymentSystem = "Visa";
                    emailStmntNSGBIndividualPDF.BillingCycle = "EOM";
                    emailStmntNSGBIndividualPDF.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBIndividualPDF.emailFrom = "noreply.e-statement@qnbalahli.com";//mmohammed@emp-group.com
                    emailStmntNSGBIndividualPDF.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBIndividualPDF.bankWebLink = "qnbalahli.com";
                    emailStmntNSGBIndividualPDF.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    //emailStmntNSGBIndividualPDF.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    //emailStmntNSGBIndividualPDF.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.gif";
                    //emailStmntNSGBIndividualPDF.bottomBanner = @"D:\pC#\ProjData\Statement\UNBN\bottom.jpg";
                    emailStmntNSGBIndividualPDF.waitPeriod = 3000;
                    emailStmntNSGBIndividualPDF.HasAttachement = true;
                    emailStmntNSGBIndividualPDF.isReward = false;
                    emailStmntNSGBIndividualPDF.rewardCond = ""; //"('Reward Program - ')";

                    emailStmntNSGBIndividualPDF.IsCardProduct = true;
                    emailStmntNSGBIndividualPDF.productCond = "('Visa Business - Individuals')";
                    emailStmntNSGBIndividualPDF.IsContractType = false;
                    emailStmntNSGBIndividualPDF.ContractTypeCond = "('Business card for individuals')";

                    emailStmntNSGBIndividualPDF.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\Individual\Individual.jpg";
                    emailStmntNSGBIndividualPDF.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\Individual\E-News";
                    emailStmntNSGBIndividualPDF.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\Individual\Individual.txt";


                    emailStmntNSGBIndividualPDF.setFrm = this;
                    checkErrRslt = emailStmntNSGBIndividualPDF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntNSGBIndividualPDF = null;
                    break;

                case 424:   // 1) NSGB Credit >> MasterCard Standard PDF Email 1/m
                    clsStatHtmlQNBPDF emailStmntNSGBmcStandardPDF = new clsStatHtmlQNBPDF();
                    emailStmntNSGBmcStandardPDF.isSplitted = true;
                    //statHtmlunbn.PaymentSystem = "Visa";
                    emailStmntNSGBmcStandardPDF.BillingCycle = "EOM";
                    emailStmntNSGBmcStandardPDF.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBmcStandardPDF.emailFrom = "noreply.e-statement@qnbalahli.com";//mmohammed@emp-group.com
                    emailStmntNSGBmcStandardPDF.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBmcStandardPDF.bankWebLink = "qnbalahli.com";
                    emailStmntNSGBmcStandardPDF.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    //emailStmntNSGBmcStandardPDF.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    //emailStmntNSGBmcStandardPDF.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.gif";
                    //emailStmntNSGBmcStandardPDF.bottomBanner = @"D:\pC#\ProjData\Statement\UNBN\bottom.jpg";
                    emailStmntNSGBmcStandardPDF.waitPeriod = 3000;
                    emailStmntNSGBmcStandardPDF.HasAttachement = true;
                    emailStmntNSGBmcStandardPDF.isReward = true;
                    emailStmntNSGBmcStandardPDF.rewardCond = "('Reward Program - MasterCard Standard')";

                    emailStmntNSGBmcStandardPDF.IsCardProduct = true;
                    emailStmntNSGBmcStandardPDF.productCond = "('MasterCard Standard')";
                    emailStmntNSGBmcStandardPDF.IsContractType = false;
                    emailStmntNSGBmcStandardPDF.ContractTypeCond = "('MasterCard Standard - Staff')";

                    emailStmntNSGBmcStandardPDF.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\MasterCardStandard\MasterCardStandard.jpg";
                    emailStmntNSGBmcStandardPDF.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\MasterCardStandard\E-News";
                    emailStmntNSGBmcStandardPDF.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\MasterCardStandard\MasterCardStandard.txt";


                    emailStmntNSGBmcStandardPDF.setFrm = this;
                    checkErrRslt = emailStmntNSGBmcStandardPDF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntNSGBmcStandardPDF = null;
                    break;

                case 425:   // 1) NSGB Credit >> Titanium PDF Email 1/m
                    clsStatHtmlQNBPDF emailStmntNSGBmcTitaniumPDF = new clsStatHtmlQNBPDF();
                    emailStmntNSGBmcTitaniumPDF.isSplitted = true;
                    //statHtmlunbn.PaymentSystem = "Visa";
                    emailStmntNSGBmcTitaniumPDF.BillingCycle = "EOM";
                    emailStmntNSGBmcTitaniumPDF.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBmcTitaniumPDF.emailFrom = "noreply.e-statement@qnbalahli.com";//mmohammed@emp-group.com
                    emailStmntNSGBmcTitaniumPDF.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBmcTitaniumPDF.bankWebLink = "qnbalahli.com";
                    emailStmntNSGBmcTitaniumPDF.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    //emailStmntNSGBmcTitaniumPDF.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    //emailStmntNSGBmcTitaniumPDF.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.gif";
                    //emailStmntNSGBmcTitaniumPDF.bottomBanner = @"D:\pC#\ProjData\Statement\UNBN\bottom.jpg";
                    emailStmntNSGBmcTitaniumPDF.waitPeriod = 3000;
                    emailStmntNSGBmcTitaniumPDF.HasAttachement = true;
                    emailStmntNSGBmcTitaniumPDF.isReward = true;
                    emailStmntNSGBmcTitaniumPDF.rewardCond = "('Reward Program - MasterCard TITANIUM')";

                    emailStmntNSGBmcTitaniumPDF.IsCardProduct = true;
                    emailStmntNSGBmcTitaniumPDF.productCond = "('MasterCard Titanium CR')";
                    emailStmntNSGBmcTitaniumPDF.IsContractType = false;
                    emailStmntNSGBmcTitaniumPDF.ContractTypeCond = "('MasterCard Titanium CR - Staff')";

                    emailStmntNSGBmcTitaniumPDF.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\MasterCardTitanium\MasterCardTitanium.jpg";
                    emailStmntNSGBmcTitaniumPDF.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\MasterCardTitanium\E-News";
                    emailStmntNSGBmcTitaniumPDF.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\MasterCardTitanium\MasterCardTitanium.txt";


                    emailStmntNSGBmcTitaniumPDF.setFrm = this;
                    checkErrRslt = emailStmntNSGBmcTitaniumPDF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntNSGBmcTitaniumPDF = null;
                    break;

                case 446:   // 1) NSGB Credit >> MasterCard World Elite PDF Email 1/m
                    clsStatHtmlQNBPDF emailStmntNSGBMC_World_ElitePDF = new clsStatHtmlQNBPDF();
                    emailStmntNSGBMC_World_ElitePDF.isSplitted = true;
                    emailStmntNSGBMC_World_ElitePDF.BillingCycle = "EOM";
                    emailStmntNSGBMC_World_ElitePDF.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBMC_World_ElitePDF.emailFrom = "noreply.e-statement@qnbalahli.com";//mmohammed@emp-group.com
                    emailStmntNSGBMC_World_ElitePDF.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBMC_World_ElitePDF.bankWebLink = "qnbalahli.com";
                    emailStmntNSGBMC_World_ElitePDF.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBMC_World_ElitePDF.waitPeriod = 3000;
                    emailStmntNSGBMC_World_ElitePDF.HasAttachement = true;
                    emailStmntNSGBMC_World_ElitePDF.isReward = true;
                    emailStmntNSGBMC_World_ElitePDF.rewardCond = "('Reward Program - MasterCard World Elite')";
                    emailStmntNSGBMC_World_ElitePDF.IsCardProduct = true;
                    emailStmntNSGBMC_World_ElitePDF.productCond = "('MasterCard World Elite CR')";
                    emailStmntNSGBMC_World_ElitePDF.IsContractType = false;
                    emailStmntNSGBMC_World_ElitePDF.ContractTypeCond = "('MasterCard World Elite - Staff')";
                    emailStmntNSGBMC_World_ElitePDF.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\WorldElite\WorldElite.jpg";
                    emailStmntNSGBMC_World_ElitePDF.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\WorldElite\E-News";
                    emailStmntNSGBMC_World_ElitePDF.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\WorldElite\WorldElite.txt";

                    emailStmntNSGBMC_World_ElitePDF.setFrm = this;
                    checkErrRslt = emailStmntNSGBMC_World_ElitePDF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntNSGBMC_World_ElitePDF = null;
                    break;

                case 472: // 1) QNB ALAHLI Credit >> VISA Signature PDF Email 1/m 1/m
                    clsStatHtmlQNBPDF emailStmntNSGBMC_Visa_SignaturePDF = new clsStatHtmlQNBPDF();
                    emailStmntNSGBMC_Visa_SignaturePDF.isSplitted = true;
                    emailStmntNSGBMC_Visa_SignaturePDF.BillingCycle = "EOM";
                    emailStmntNSGBMC_Visa_SignaturePDF.emailFromName = "QNB ALAHLI - Statement";
                    emailStmntNSGBMC_Visa_SignaturePDF.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBMC_Visa_SignaturePDF.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";
                    emailStmntNSGBMC_Visa_SignaturePDF.bankWebLink = "qnbalahli.com";
                    emailStmntNSGBMC_Visa_SignaturePDF.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBMC_Visa_SignaturePDF.waitPeriod = 3000;
                    emailStmntNSGBMC_Visa_SignaturePDF.HasAttachement = true;
                    emailStmntNSGBMC_Visa_SignaturePDF.isReward = true;
                    emailStmntNSGBMC_Visa_SignaturePDF.rewardCond = "('Reward Program - Visa Signature')";
                    emailStmntNSGBMC_Visa_SignaturePDF.IsCardProduct = true;
                    emailStmntNSGBMC_Visa_SignaturePDF.productCond = "('Visa Signature')";
                    emailStmntNSGBMC_Visa_SignaturePDF.IsContractType = false;
                    emailStmntNSGBMC_Visa_SignaturePDF.ContractTypeCond = "('VISA Signature - Standard')";
                    emailStmntNSGBMC_Visa_SignaturePDF.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\VisaSignature\VisaSignature.jpg";
                    emailStmntNSGBMC_Visa_SignaturePDF.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\VisaSignature\E-News";
                    emailStmntNSGBMC_Visa_SignaturePDF.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\VisaSignature\VisaSignature.txt";
                    emailStmntNSGBMC_Visa_SignaturePDF.setFrm = this;
                    checkErrRslt = emailStmntNSGBMC_Visa_SignaturePDF.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntNSGBMC_World_ElitePDF = null;
                    break;
                case 473: // 1) QNB ALAHLI Credit >> bebasata Gold Credit PDF Email 1/m 1/m
                    clsStatHtmlQNBPDF emailStmntNSGBMC_Gold_En = new clsStatHtmlQNBPDF();
                    emailStmntNSGBMC_Gold_En.isSplitted = true;
                    emailStmntNSGBMC_Gold_En.BillingCycle = "EOM";
                    emailStmntNSGBMC_Gold_En.emailFromName = "QNB ALAHLI - Statement";
                    emailStmntNSGBMC_Gold_En.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBMC_Gold_En.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";
                    emailStmntNSGBMC_Gold_En.bankWebLink = "qnbalahli.com";
                    emailStmntNSGBMC_Gold_En.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBMC_Gold_En.waitPeriod = 3000;
                    emailStmntNSGBMC_Gold_En.HasAttachement = true;
                    emailStmntNSGBMC_Gold_En.isReward = true;
                    emailStmntNSGBMC_Gold_En.rewardCond = "('Reward Program - MasterCard Gold En')";//????
                    emailStmntNSGBMC_Gold_En.IsCardProduct = true;
                    emailStmntNSGBMC_Gold_En.productCond = "('bebasata Gold Credit Card')";
                    emailStmntNSGBMC_Gold_En.IsContractType = false;
                    emailStmntNSGBMC_Gold_En.ContractTypeCond = "('MasterCard Gold En - Staff','MasterCard Gold En - Standard','MasterCard Gold En - Standard (XML ONLY)','MasterCard Gold En - Staff(XML ONLY)','MasterCard Gold En - Club')";//MasterCard Gold En - Staff  -- rewardCond - Standard
                    emailStmntNSGBMC_Gold_En.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\BebasataGoldCredit\BebasataGoldCredit.jpg";
                    emailStmntNSGBMC_Gold_En.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\BebasataGoldCredit\E-News";
                    emailStmntNSGBMC_Gold_En.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\BebasataGoldCredit\BebasataGoldCredit.txt";
                    emailStmntNSGBMC_Gold_En.setFrm = this;
                    checkErrRslt = emailStmntNSGBMC_Gold_En.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    emailStmntNSGBMC_World_ElitePDF = null;
                    break;

                case 172:   // 1) NSGB Credit >> Infinite Email 1/m
                    clsStatementNSGBcreditHtml emailStmntNSGBInfinite = new clsStatementNSGBcreditHtml();
                    emailStmntNSGBInfinite.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    //emailStmntNSGBInfinite.emailFrom = "Card.Support@QNBALAHLI.COM";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBInfinite.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBInfinite.bankWebLink = "qnbalahli.com";//"www.emp-group.com""www.socgen.com"
                    //emailStmntNSGBInfinite.bankWebLinkService = " Card.Support@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBInfinite.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBInfinite.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBInfinite.logoAlignment = "left";
                    //emailStmntNSGBInfinite.statMessageFile = @"D:\pC#\ProjData\Statement\NSGB\Infinite\EmailMessage.txt";
                    emailStmntNSGBInfinite.statMessageBox = @"D:\pC#\ProjData\Statement\NSGB\EmailMessageBox.txt";
                    emailStmntNSGBInfinite.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\Infinite\Infinite.jpg";
                    emailStmntNSGBInfinite.IsSplitted = true;
                    emailStmntNSGBInfinite.productCond = "Visa infinite";
                    emailStmntNSGBInfinite.HasAttachement = true;
                    emailStmntNSGBInfinite.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\Infinite\E-News";
                    emailStmntNSGBInfinite.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\Infinite\Infinite.txt";
                    //NSGB-3615
                    //emailStmntNSGBInfinite.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.gif";
                    emailStmntNSGBInfinite.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.jpg";
                    emailStmntNSGBInfinite.waitPeriod = 3000;
                    emailStmntNSGBInfinite.setFrm = this;
                    emailStmntNSGBInfinite.isReward = true;
                    emailStmntNSGBInfinite.rewardCond = "('Reward Program - Visa infinite')";
                    checkErrRslt = emailStmntNSGBInfinite.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntNSGBInfinite = null;
                    break;

                case 174:   // 1) NSGB Credit >> MasterCard Standard Email 1/m
                    clsStatementNSGBcreditHtml emailStmntNSGBMasterCardStandard = new clsStatementNSGBcreditHtml();
                    emailStmntNSGBMasterCardStandard.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    //emailStmntNSGBMasterCardStandard.emailFrom = "Card.Support@QNBALAHLI.COM";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBMasterCardStandard.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBMasterCardStandard.bankWebLink = "qnbalahli.com";//"www.emp-group.com""www.socgen.com"
                    //emailStmntNSGBMasterCardStandard.bankWebLinkService = " Card.Support@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBMasterCardStandard.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBMasterCardStandard.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBMasterCardStandard.logoAlignment = "left";
                    //emailStmntNSGBMasterCardStandard.statMessageFile = @"D:\pC#\ProjData\Statement\NSGB\MasterCardStandard\EmailMessage.txt";
                    emailStmntNSGBMasterCardStandard.statMessageBox = @"D:\pC#\ProjData\Statement\NSGB\EmailMessageBox.txt";
                    emailStmntNSGBMasterCardStandard.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\MasterCardStandard\MasterCardStandard.jpg";
                    emailStmntNSGBMasterCardStandard.IsSplitted = true;
                    emailStmntNSGBMasterCardStandard.productCond = "MasterCard Standard";
                    emailStmntNSGBMasterCardStandard.HasAttachement = true;
                    emailStmntNSGBMasterCardStandard.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\MasterCardStandard\E-News";
                    emailStmntNSGBMasterCardStandard.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\MasterCardStandard\MasterCardStandard.txt";
                    emailStmntNSGBMasterCardStandard.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\MCLogo.jpg";
                    emailStmntNSGBMasterCardStandard.waitPeriod = 3000;
                    emailStmntNSGBMasterCardStandard.setFrm = this;
                    emailStmntNSGBMasterCardStandard.isReward = true;
                    emailStmntNSGBMasterCardStandard.rewardCond = "('Reward Program - MasterCard Standard')";
                    checkErrRslt = emailStmntNSGBMasterCardStandard.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntNSGBMasterCardStandard = null;
                    break;

                case 403:   // 1) NSGB Credit >> MasterCard Titanium Email 1/m
                    clsStatementNSGBcreditHtml emailStmntNSGBMasterCardTitanium = new clsStatementNSGBcreditHtml();
                    emailStmntNSGBMasterCardTitanium.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    //emailStmntNSGBMasterCardTitanium.emailFrom = "Card.Support@QNBALAHLI.COM";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBMasterCardTitanium.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBMasterCardTitanium.bankWebLink = "qnbalahli.com";//"www.emp-group.com""www.socgen.com"
                    //emailStmntNSGBMasterCardTitanium.bankWebLinkService = " Card.Support@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBMasterCardTitanium.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBMasterCardTitanium.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBMasterCardTitanium.logoAlignment = "left";
                    //emailStmntNSGBMasterCardTitanium.statMessageFile = @"D:\pC#\ProjData\Statement\NSGB\MasterCardTitanium\EmailMessage.txt";
                    emailStmntNSGBMasterCardTitanium.statMessageBox = @"D:\pC#\ProjData\Statement\NSGB\EmailMessageBox.txt";
                    emailStmntNSGBMasterCardTitanium.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\MasterCardTitanium\MasterCardTitanium.jpg";
                    emailStmntNSGBMasterCardTitanium.IsSplitted = true;
                    emailStmntNSGBMasterCardTitanium.productCond = "MasterCard Titanium CR";
                    emailStmntNSGBMasterCardTitanium.HasAttachement = true;
                    emailStmntNSGBMasterCardTitanium.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\MasterCardTitanium\E-News";
                    emailStmntNSGBMasterCardTitanium.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\MasterCardTitanium\MasterCardTitanium.txt";
                    emailStmntNSGBMasterCardTitanium.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\MCLogo.jpg";
                    emailStmntNSGBMasterCardTitanium.waitPeriod = 3000;
                    emailStmntNSGBMasterCardTitanium.setFrm = this;
                    emailStmntNSGBMasterCardTitanium.isReward = true;
                    emailStmntNSGBMasterCardTitanium.rewardCond = "('Reward Program - MasterCard TITANIUM')";
                    checkErrRslt = emailStmntNSGBMasterCardTitanium.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntNSGBMasterCardTitanium = null;
                    break;

                case 236:   // 1) NSGB MC Business SME >> Email 1/m
                    clsStatementNSGBcreditHtml emailStmntNSGBMCSME = new clsStatementNSGBcreditHtml();
                    emailStmntNSGBMCSME.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBMCSME.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBMCSME.bankWebLink = "qnbalahli.com";//"www.emp-group.com""www.socgen.com"
                    emailStmntNSGBMCSME.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBMCSME.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBMCSME.logoAlignment = "left";
                    //emailStmntNSGBMCSME.statMessageFile = @"D:\pC#\ProjData\Statement\NSGB\MasterCardSME\EmailMessage.txt";
                    //emailStmntNSGBMCSME.statMessageBox = @"D:\pC#\ProjData\Statement\NSGB\EmailMessageBox.txt";
                    emailStmntNSGBMCSME.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\MasterCardSME\MasterCardSME.jpg";
                    emailStmntNSGBMCSME.IsSplitted = true;
                    emailStmntNSGBMCSME.productCond = "MasterCard Business - SME";
                    emailStmntNSGBMCSME.HasAttachement = true;
                    emailStmntNSGBMCSME.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\MasterCardSME\E-News";
                    emailStmntNSGBMCSME.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\MasterCardSME\MasterCardSME.txt";
                    emailStmntNSGBMCSME.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\MCLogo.jpg";
                    emailStmntNSGBMCSME.waitPeriod = 3000;
                    emailStmntNSGBMCSME.setFrm = this;
                    emailStmntNSGBMCSME.CreateCorporate = true;
                    checkErrRslt = emailStmntNSGBMCSME.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntNSGBMCSME = null;
                    break;

                case 443:   // 1) NSGB Visa Business B2B >> Email 1/m
                    clsStatementNSGBcreditHtml emailStmntNSGBB2B = new clsStatementNSGBcreditHtml();
                    emailStmntNSGBB2B.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBB2B.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBB2B.bankWebLink = "qnbalahli.com";//"www.emp-group.com""www.socgen.com"
                    emailStmntNSGBB2B.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBB2B.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBB2B.logoAlignment = "left";
                    emailStmntNSGBB2B.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\B2B\B2B.jpg";
                    emailStmntNSGBB2B.IsSplitted = true;
                    emailStmntNSGBB2B.productCond = "VISA Business - B2B";
                    emailStmntNSGBB2B.HasAttachement = true;
                    emailStmntNSGBB2B.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\B2B\E-News";
                    emailStmntNSGBB2B.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\B2B\B2B.txt";
                    emailStmntNSGBB2B.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.jpg";
                    emailStmntNSGBB2B.waitPeriod = 3000;
                    emailStmntNSGBB2B.setFrm = this;
                    emailStmntNSGBB2B.CreateCorporate = true;
                    checkErrRslt = emailStmntNSGBB2B.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntNSGBB2B = null;
                    break;

                case 444:   // 1) NSGB Visa Business FEDCOC >> Email 1/m
                    clsStatementNSGBcreditHtml emailStmntNSGBFEDCOC = new clsStatementNSGBcreditHtml();
                    emailStmntNSGBFEDCOC.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBFEDCOC.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBFEDCOC.bankWebLink = "qnbalahli.com";//"www.emp-group.com""www.socgen.com"
                    emailStmntNSGBFEDCOC.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBFEDCOC.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBFEDCOC.logoAlignment = "left";
                    emailStmntNSGBFEDCOC.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\FEDCOC\FEDCOC.jpg";
                    emailStmntNSGBFEDCOC.IsSplitted = true;
                    emailStmntNSGBFEDCOC.productCond = "Visa Business - FEDCOC";
                    emailStmntNSGBFEDCOC.HasAttachement = true;
                    emailStmntNSGBFEDCOC.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\FEDCOC\E-News";
                    emailStmntNSGBFEDCOC.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\FEDCOC\FEDCOC.txt";
                    emailStmntNSGBFEDCOC.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.jpg";
                    emailStmntNSGBFEDCOC.waitPeriod = 3000;
                    emailStmntNSGBFEDCOC.setFrm = this;
                    emailStmntNSGBFEDCOC.CreateCorporate = true;
                    checkErrRslt = emailStmntNSGBFEDCOC.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntNSGBFEDCOC = null;
                    break;

                case 317:   // 1) NSGB VISA Business Corporate >> Email 1/m
                    clsStatementNSGBcreditHtml emailStmntNSGBVISACORP = new clsStatementNSGBcreditHtml();
                    emailStmntNSGBVISACORP.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBVISACORP.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBVISACORP.bankWebLink = "qnbalahli.com";//"www.emp-group.com""www.socgen.com"
                    emailStmntNSGBVISACORP.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBVISACORP.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBVISACORP.logoAlignment = "left";
                    //emailStmntNSGBVISACORP.statMessageFile = @"D:\pC#\ProjData\Statement\NSGB\MasterCardSME\EmailMessage.txt";
                    //emailStmntNSGBVISACORP.statMessageBox = @"D:\pC#\ProjData\Statement\NSGB\EmailMessageBox.txt";
                    emailStmntNSGBVISACORP.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\Business\Business.jpg";
                    emailStmntNSGBVISACORP.IsSplitted = true;
                    emailStmntNSGBVISACORP.productCond = "Visa Business Corporate";
                    emailStmntNSGBVISACORP.HasAttachement = true;
                    emailStmntNSGBVISACORP.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\Business\E-News";
                    emailStmntNSGBVISACORP.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\Business\Business.txt";
                    emailStmntNSGBVISACORP.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.jpg";
                    emailStmntNSGBVISACORP.waitPeriod = 3000;
                    emailStmntNSGBVISACORP.setFrm = this;
                    emailStmntNSGBVISACORP.CreateCorporate = true;
                    checkErrRslt = emailStmntNSGBVISACORP.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntNSGBVISACORP = null;
                    break;

                case 3001:   // 1) NSGB VISA Business Corporate >> Email 1/m
                    clsStatementNSGBcreditHtml emailStmntNSGBVISAb2bCORP = new clsStatementNSGBcreditHtml();
                    emailStmntNSGBVISAb2bCORP.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBVISAb2bCORP.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBVISAb2bCORP.bankWebLink = "qnbalahli.com";//"www.emp-group.com""www.socgen.com"
                    emailStmntNSGBVISAb2bCORP.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBVISAb2bCORP.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBVISAb2bCORP.logoAlignment = "left";
                    //emailStmntNSGBVISACORP.statMessageFile = @"D:\pC#\ProjData\Statement\NSGB\MasterCardSME\EmailMessage.txt";
                    //emailStmntNSGBVISACORP.statMessageBox = @"D:\pC#\ProjData\Statement\NSGB\EmailMessageBox.txt";
                    emailStmntNSGBVISAb2bCORP.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\Business\Business.jpg";
                    emailStmntNSGBVISAb2bCORP.IsSplitted = true;
                    emailStmntNSGBVISAb2bCORP.productCond = "VISA Platinum B2B";
                    emailStmntNSGBVISAb2bCORP.HasAttachement = true;
                    emailStmntNSGBVISAb2bCORP.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\Business\E-News";
                    emailStmntNSGBVISAb2bCORP.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\Business\Business.txt";
                    emailStmntNSGBVISAb2bCORP.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.jpg";
                    emailStmntNSGBVISAb2bCORP.waitPeriod = 3000;
                    emailStmntNSGBVISAb2bCORP.setFrm = this;
                    emailStmntNSGBVISAb2bCORP.CreateCorporate = true;
                    checkErrRslt = emailStmntNSGBVISAb2bCORP.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntNSGBVISAb2bCORP = null;
                    break;


                case 3000:   // 1) NSGB VISA Business Corporate >> Email 1/m
                    clsStatementNSGBcreditHtml emailStmntNSGBVISAMVSECORP = new clsStatementNSGBcreditHtml();
                    emailStmntNSGBVISAMVSECORP.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBVISAMVSECORP.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBVISAMVSECORP.bankWebLink = "qnbalahli.com";//"www.emp-group.com""www.socgen.com"
                    emailStmntNSGBVISAMVSECORP.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBVISAMVSECORP.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBVISAMVSECORP.logoAlignment = "left";
                    //emailStmntNSGBVISACORP.statMessageFile = @"D:\pC#\ProjData\Statement\NSGB\MasterCardSME\EmailMessage.txt";
                    //emailStmntNSGBVISACORP.statMessageBox = @"D:\pC#\ProjData\Statement\NSGB\EmailMessageBox.txt";
                    emailStmntNSGBVISAMVSECORP.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\Business\Business.jpg";
                    emailStmntNSGBVISAMVSECORP.IsSplitted = true;
                    emailStmntNSGBVISAMVSECORP.productCond = "Visa Platinum MVSE";
                    emailStmntNSGBVISAMVSECORP.HasAttachement = true;
                    emailStmntNSGBVISAMVSECORP.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\Business\E-News";
                    emailStmntNSGBVISAMVSECORP.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\Business\Business.txt";
                    emailStmntNSGBVISAMVSECORP.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.jpg";
                    emailStmntNSGBVISAMVSECORP.waitPeriod = 3000;
                    emailStmntNSGBVISAMVSECORP.setFrm = this;
                    emailStmntNSGBVISAMVSECORP.CreateCorporate = true;
                    checkErrRslt = emailStmntNSGBVISAMVSECORP.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntNSGBVISAMVSECORP = null;
                    break;

                case 3002:   // 1) NSGB VISA Business Corporate >> Email 1/m
                    clsStatementNSGBcreditHtml emailStmntNSGBVISACORCORP = new clsStatementNSGBcreditHtml();
                    emailStmntNSGBVISACORCORP.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBVISACORCORP.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBVISACORCORP.bankWebLink = "qnbalahli.com";//"www.emp-group.com""www.socgen.com"
                    emailStmntNSGBVISACORCORP.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBVISACORCORP.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBVISACORCORP.logoAlignment = "left";
                    //emailStmntNSGBVISACORP.statMessageFile = @"D:\pC#\ProjData\Statement\NSGB\MasterCardSME\EmailMessage.txt";
                    //emailStmntNSGBVISACORP.statMessageBox = @"D:\pC#\ProjData\Statement\NSGB\EmailMessageBox.txt";
                    emailStmntNSGBVISACORCORP.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\Business\Business.jpg";
                    emailStmntNSGBVISACORCORP.IsSplitted = true;
                    emailStmntNSGBVISACORCORP.productCond = "Visa Platinum Corporate";
                    emailStmntNSGBVISACORCORP.HasAttachement = true;
                    emailStmntNSGBVISACORCORP.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\Business\E-News";
                    emailStmntNSGBVISACORCORP.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\Business\Business.txt";
                    emailStmntNSGBVISACORCORP.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.jpg";
                    emailStmntNSGBVISACORCORP.waitPeriod = 3000;
                    emailStmntNSGBVISACORCORP.setFrm = this;
                    emailStmntNSGBVISACORCORP.CreateCorporate = true;
                    checkErrRslt = emailStmntNSGBVISACORCORP.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntNSGBVISACORCORP = null;
                    break;

                case 3003:   // 1) NSGB VISA Business Corporate >> Email 1/m
                    clsStatementNSGBcreditHtml emailStmntNSGBVISASMECORP = new clsStatementNSGBcreditHtml();
                    emailStmntNSGBVISASMECORP.emailFromName = "QNB ALAHLI - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
                    emailStmntNSGBVISASMECORP.emailFrom = "noreply.e-statement@qnbalahli.com";
                    emailStmntNSGBVISASMECORP.bankWebLink = "qnbalahli.com";//"www.emp-group.com""www.socgen.com"
                    emailStmntNSGBVISASMECORP.bankWebLinkService = " noreply.e-statement@QNBALAHLI.COM";//"www.socgen.com"
                    emailStmntNSGBVISASMECORP.bankLogo = @"D:\pC#\ProjData\Statement\NSGB\logo.gif";
                    emailStmntNSGBVISASMECORP.logoAlignment = "left";
                    //emailStmntNSGBVISACORP.statMessageFile = @"D:\pC#\ProjData\Statement\NSGB\MasterCardSME\EmailMessage.txt";
                    //emailStmntNSGBVISACORP.statMessageBox = @"D:\pC#\ProjData\Statement\NSGB\EmailMessageBox.txt";
                    emailStmntNSGBVISASMECORP.bottomBanner = @"D:\pC#\ProjData\Statement\NSGB\Business\Business.jpg";
                    emailStmntNSGBVISASMECORP.IsSplitted = true;
                    emailStmntNSGBVISASMECORP.productCond = "Visa Platinum SME";
                    emailStmntNSGBVISASMECORP.HasAttachement = true;
                    emailStmntNSGBVISASMECORP.attachedFiles = @"D:\pC#\ProjData\Statement\NSGB\Business\E-News";
                    emailStmntNSGBVISASMECORP.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\NSGB\Business\Business.txt";
                    emailStmntNSGBVISASMECORP.visaLogo = @"D:\pC#\ProjData\Statement\NSGB\VisaLogo.jpg";
                    emailStmntNSGBVISASMECORP.waitPeriod = 3000;
                    emailStmntNSGBVISASMECORP.setFrm = this;
                    emailStmntNSGBVISASMECORP.CreateCorporate = true;
                    checkErrRslt = emailStmntNSGBVISASMECORP.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData);
                    emailStmntNSGBVISASMECORP = null;
                    break;

                case 405:   //[41] BDCA Banque Du Caire Classic  >> Credit Emails 5/m
                case 406:   //[41] BDCA Banque Du Caire Gold  >> Credit Emails 5/m
                case 407:   //[41] BDCA Banque Du Caire MasterCard Standard  >> Credit Emails 5/m
                case 408:   //[41] BDCA Banque Du Caire MasterCard Gold  >> Credit Emails 5/m
                case 409:   //[41] BDCA Banque Du Caire MasterCard Installment  >> Credit Emails 5/m
                case 410:   //[41] BDCA Banque Du Caire MasterCard Titanium  >> Credit Emails 5/m
                case 411:   //[41] BDCA Banque Du Caire MasterCard World Elite  >> Credit Emails 5/m
                case 511:   //[41] BDCA Banque Du Caire MasterCard Platinum  >> Credit email 5/m
                case 512:   //[41] BDCA Banque Du Caire MasterCard Corporate  >> Credit email 1/m
                    clsStatHtmlBDCA statHtmlBDCA = new clsStatHtmlBDCA();
                    statHtmlBDCA.emailFromName = "BDC E-Statement";
                    statHtmlBDCA.emailFrom = "bdc.estatment@bdc.com.eg";
                    statHtmlBDCA.bankWebLink = "https://bit.ly/33peXqd";
                    statHtmlBDCA.bankLogo = @"D:\pC#\ProjData\Statement\BDCA\BDCALogo.png";
                    statHtmlBDCA.BackGround = @"D:\pC#\ProjData\Statement\BDCA\EmailBody.jpg";
                    //statHtmlBDCA.statMessageFileMonthly = @"D:\pC#\ProjData\Statement\BDCA\EmailMessageBox.txt";
                    statHtmlBDCA.waitPeriod = 3000;
                    statHtmlBDCA.HasAttachement = true;
                    statHtmlBDCA.setFrm = this;
                    if (pCmbProducts == 405)
                    {
                        statHtmlBDCA.productCond = "('Visa Classic','Visa Classic - STAFF')";
                        //statHtmlBDCA.productCond = "('Visa Classic - STAFF')";
                        //statHtmlBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        //statHtmlBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 406)
                    {
                        statHtmlBDCA.productCond = "('Visa Gold','Visa Gold - STAFF')";
                        //statHtmlBDCA.productCond = "('Visa Gold - STAFF')";
                        //statHtmlBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        //statHtmlBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 407)
                    {
                        statHtmlBDCA.productCond = "('MasterCard Standard','MasterCard Standard 1','MasterCard Standard - STAFF','MasterCard Standard 1 - STAFF')";
                        //statHtmlBDCA.productCond = "('MasterCard Standard - STAFF','MasterCard Standard 1 - STAFF')";
                        //statHtmlBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        //statHtmlBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 408)
                    {
                        statHtmlBDCA.productCond = "('MasterCard Gold', 'MasterCard Gold 1' ,'MasterCard Gold old','MasterCard Gold - STAFF', 'MasterCard Gold 1 - STAFF' ,'MasterCard Gold old - STAFF')";
                        //statHtmlBDCA.productCond = "('MasterCard Gold - STAFF', 'MasterCard Gold 1 - STAFF' ,'MasterCard Gold old - STAFF')";
                        //statHtmlBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        //statHtmlBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 409)
                    {
                        statHtmlBDCA.productCond = "('MasterCard Installment', 'MasterCard Installment 1' ,'MasterCard Installment old','MasterCard Installment - STAFF', 'MasterCard Installment 1 - STAFF' ,'MasterCard Installment old - STAFF')";
                        //statHtmlBDCA.productCond = "('MasterCard Installment - STAFF', 'MasterCard Installment 1 - STAFF' ,'MasterCard Installment old - STAFF')";
                        //statHtmlBDCA.InstallmentCondition = "('BDC Installment Card Program')";
                        //statHtmlBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 410)
                    {
                        statHtmlBDCA.productCond = "('MasterCard Titanium', 'MasterCard Titanium 1' ,'MasterCard Co-Brand Titanium Credit','MasterCard Titanium - STAFF', 'MasterCard Titanium 1 - STAFF' ,'MasterCard Co-Brand Titanium Credit - STAFF')";
                        //statHtmlBDCA.productCond = "('MasterCard Titanium - STAFF', 'MasterCard Titanium 1 - STAFF' ,'MasterCard Co-Brand Titanium Credit - STAFF')";
                        //statHtmlBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        //statHtmlBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 411)
                    {
                        statHtmlBDCA.productCond = "('MasterCard World Elite','MasterCard World Elite - STAFF')";
                        //statHtmlBDCA.productCond = "('MasterCard World Elite - STAFF')";
                        //statHtmlBDCA.InstallmentCondition = "('BDC Easy Payment Plan')";
                        //statHtmlBDCA.isInstallmentVal = true;
                    }
                    if (pCmbProducts == 511)
                    {
                        statHtmlBDCA.productCond = "('MC Platinum Credit Card','MC Platinum Credit Card - STAFF')";
                    }
                    if (pCmbProducts == 512)
                    {
                        statHtmlBDCA.productCond = "('MC Corporate Credit Card','MC Corporate Credit Card - STAFF')";
                    }

                    checkErrRslt = statHtmlBDCA.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlBDCA = null;
                    break;


                //AIBK iatta
                case 302:   //[127] AIBK Arab Investment Bank of Egypt  >> Credit Emails 1/m
                    //clsMaintainData maintaingtbkestmt = new clsMaintainData();
                    //maintaingtbkestmt.fixAddress(bankCode);
                    //maintaingtbkestmt = null;
                    clsStatHtmlAIBK statHtmlAIBK = new clsStatHtmlAIBK();
                    //statHtmlAIBK.emailFromName = "Arab Investment Bank - E-statement";
                    statHtmlAIBK.emailFromName = "aiBANK - e-statement";
                    //statHtmlAIBK.emailFrom = "AIBKBank@emp-group.com"; estatement@aibegypt.com
                    statHtmlAIBK.emailFrom = "estatement@aibegypt.com";
                    statHtmlAIBK.bankWebLink = "www.AIBK.com";
                    //statHtmlAIBK.bankLogo = @"D:\pC#\ProjData\Statement\AIBK\AIBK.jpg";
                    statHtmlAIBK.bankLogo = @"D:\pC#\ProjData\Statement\AIBK\aiBANK.png";
                    statHtmlAIBK.waitPeriod = 3000;
                    statHtmlAIBK.HasAttachement = true;
                    statHtmlAIBK.setFrm = this;
                    statHtmlAIBK.IsSplitted = true;
                    statHtmlAIBK.productCond = "VISA VALU Reward";
                    checkErrRslt = statHtmlAIBK.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlAIBK = null;
                    break;

                //AIBK iatta
                case 3097:   //[127] AIBK Arab Investment Bank of Egypt  >> Credit Emails 1/m
                    //clsMaintainData maintaingtbkestmt = new clsMaintainData();
                    //maintaingtbkestmt.fixAddress(bankCode);
                    //maintaingtbkestmt = null;
                    var statHtmlAIBK_valu = new clsStatHtmlAIBKValu();
                    //statHtmlAIBK.emailFromName = "Arab Investment Bank - E-statement";
                    statHtmlAIBK_valu.emailFromName = "aiBANK - e-statement";
                    //statHtmlAIBK.emailFrom = "AIBKBank@emp-group.com"; estatement@aibegypt.com
                    statHtmlAIBK_valu.emailFrom = "estatement@aibegypt.com";
                    statHtmlAIBK_valu.bankWebLink = "www.AIBK.com";
                    //statHtmlAIBK_valu.bankLogo = @"D:\pC#\ProjData\Statement\AIBK\AIBK.jpg";
                    statHtmlAIBK_valu.bankLogo = @"D:\pC#\ProjData\Statement\AIBK\aiBANK.png";
                    statHtmlAIBK_valu.waitPeriod = 3000;
                    statHtmlAIBK_valu.HasAttachement = true;
                    statHtmlAIBK_valu.setFrm = this;
                    statHtmlAIBK_valu.IsSplitted = true;
                    statHtmlAIBK_valu.productCond = "VISA VALU Reward";
                    checkErrRslt = statHtmlAIBK_valu.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlAIBK_valu = null;
                    break;

                case 312:   //[127] AIBK Arab Investment Bank of Egypt  >> Installment Emails 1/m
                    //clsMaintainData maintaingtbkestmt = new clsMaintainData();
                    //maintaingtbkestmt.fixAddress(bankCode);
                    //maintaingtbkestmt = null;
                    clsStatHtmlAIBK statHtmlInstallmentAIBK = new clsStatHtmlAIBK();
                    //statHtmlInstallmentAIBK.emailFromName = "Arab Investment Bank - Installment E-statement";
                    statHtmlInstallmentAIBK.emailFromName = "aiBANK - Installment e-statement";
                    statHtmlInstallmentAIBK.emailFrom = "AIBKBank@emp-group.com";
                    statHtmlInstallmentAIBK.bankWebLink = "www.AIBK.com";
                    //statHtmlInstallmentAIBK.bankLogo = @"D:\pC#\ProjData\Statement\AIBK\AIBK.jpg";
                    statHtmlInstallmentAIBK.bankLogo = @"D:\pC#\ProjData\Statement\AIBK\aiBANK.png";
                    statHtmlInstallmentAIBK.waitPeriod = 3000;
                    statHtmlInstallmentAIBK.HasAttachement = true;
                    statHtmlInstallmentAIBK.setFrm = this;
                    statHtmlInstallmentAIBK.IsSplitted = true;
                    statHtmlInstallmentAIBK.productCond = "VISA VALU Reward";
                    checkErrRslt = statHtmlInstallmentAIBK.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlAIBK = null;
                    break;

                //iatta ALXB 

                case 301:   //[122] ALXB ALEXBANK   >> Credit Emails 30/m
                case 342:   //[122] ALXB ALEXBANK  >> Credit MF email 30/m
                    clsStatHtmlALXB statHtmlALXB = new clsStatHtmlALXB();
                    statHtmlALXB.emailFromName = "ALEXBANK E-statement";
                    statHtmlALXB.emailFrom = "ALEXBANK@emp-group.com";
                    statHtmlALXB.bankWebLink = "https://www.alexbank.com";
                    statHtmlALXB.bankfacebookLink = "https://www.facebook.com/ALEXBANKOFFICIAL";
                    statHtmlALXB.bankLinkedInLink = "https://www.linkedin.com/company/bank-of-alexandria";
                    statHtmlALXB.bankyoutubeLink = "https://www.youtube.com/user/AlexBankOfficial";
                    statHtmlALXB.bankoffersLink = "https://www.alexbank.com/products/retail/Cards/Offers";
                    statHtmlALXB.bankbackimg2 = @"D:\pC#\ProjData\Statement\ALXB\ALX2.png";
                    statHtmlALXB.bankbackimg3 = @"D:\pC#\ProjData\Statement\ALXB\ALX3.png";
                    statHtmlALXB.bankFooterLogoimg = @"D:\pC#\ProjData\Statement\ALXB\FooterLogoImg.png";
                    statHtmlALXB.bankfooterFacebookimg = @"D:\pC#\ProjData\Statement\ALXB\FooterFacebookImg.png";
                    statHtmlALXB.bankfooterLinkedInimg = @"D:\pC#\ProjData\Statement\ALXB\FooterLinkedinImg.png";
                    statHtmlALXB.bankfooterYoutubeimg = @"D:\pC#\ProjData\Statement\ALXB\FooterYoutubeImg.png";
                    statHtmlALXB.bankfooterDetailsimg = @"D:\pC#\ProjData\Statement\ALXB\FooterDetailImg.png";
                    statHtmlALXB.bankfooterOffersimg = @"D:\pC#\ProjData\Statement\ALXB\FooterOffersImg.png";
                    statHtmlALXB.bankfooterimg = @"D:\pC#\ProjData\Statement\ALXB\FooterImg.png";
                    statHtmlALXB.waitPeriod = 3000;
                    statHtmlALXB.HasAttachement = true;
                    //filter by cardProduct
                    statHtmlALXB.IsSplitted = true;
                    if (pCmbProducts == 301)
                    {
                        statHtmlALXB.bankbackimg1 = @"D:\pC#\ProjData\Statement\ALXB\ALX1.png";
                        statHtmlALXB.productCond = "('MC Platinum Credit','MC World Credit','Visa Credit Classic','Visa Credit Gold','MC Gold Credit','MC Titanium Credit')";

                    }
                    else if (pCmbProducts == 342)//MF
                    {
                        statHtmlALXB.bankbackimg1 = @"D:\pC#\ProjData\Statement\ALXB\ALXMF.png";
                        statHtmlALXB.productCond = "('MC MF Individual Credit','MC MF Corporate Credit')";
                    }
                    statHtmlALXB.setFrm = this;
                    checkErrRslt = statHtmlALXB.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlALXB = null;
                    break;

                case 438:   //[122] ALXB ALEXBANK   >> Corporate Emails 30/m
                    clsStatHtmlALXB_CP statHtmlALXBcp = new clsStatHtmlALXB_CP();
                    statHtmlALXBcp.emailFromName = "ALEXBANK E-statement";
                    statHtmlALXBcp.emailFrom = "ALEXBANK@emp-group.com";
                    statHtmlALXBcp.bankWebLink = "https://www.alexbank.com";
                    statHtmlALXBcp.bankfacebookLink = "https://www.facebook.com/ALEXBANKOFFICIAL";
                    statHtmlALXBcp.bankLinkedInLink = "https://www.linkedin.com/company/bank-of-alexandria";
                    statHtmlALXBcp.bankyoutubeLink = "https://www.youtube.com/user/AlexBankOfficial";
                    statHtmlALXBcp.bankoffersLink = "https://www.alexbank.com/products/retail/Cards/Offers";
                    statHtmlALXBcp.bankbackimg1 = @"D:\pC#\ProjData\Statement\ALXBCorp\ALX1.png";
                    statHtmlALXBcp.bankbackimg2 = @"D:\pC#\ProjData\Statement\ALXBCorp\ALX2.png";
                    statHtmlALXBcp.bankbackimg3 = @"D:\pC#\ProjData\Statement\ALXBCorp\ALX3.png";
                    statHtmlALXBcp.bankFooterLogoimg = @"D:\pC#\ProjData\Statement\ALXBCorp\FooterLogoImg.png";
                    statHtmlALXBcp.bankfooterFacebookimg = @"D:\pC#\ProjData\Statement\ALXBCorp\FooterFacebookImg.png";
                    statHtmlALXBcp.bankfooterLinkedInimg = @"D:\pC#\ProjData\Statement\ALXBCorp\FooterLinkedinImg.png";
                    statHtmlALXBcp.bankfooterYoutubeimg = @"D:\pC#\ProjData\Statement\ALXBCorp\FooterYoutubeImg.png";
                    statHtmlALXBcp.bankfooterDetailsimg = @"D:\pC#\ProjData\Statement\ALXBCorp\FooterDetailImg.png";
                    statHtmlALXBcp.bankfooterOffersimg = @"D:\pC#\ProjData\Statement\ALXBCorp\FooterOffersImg.png";
                    statHtmlALXBcp.bankfooterimg = @"D:\pC#\ProjData\Statement\ALXBCorp\FooterImg.png";
                    statHtmlALXBcp.waitPeriod = 3000;
                    statHtmlALXBcp.HasAttachement = true;
                    statHtmlALXBcp.productCond = "('VISA Corporate Cardholder', 'VISA Corporate', 'MC Corporate Executive', 'MC Corporate Executive Cardholder')";
                    statHtmlALXBcp.setFrm = this;
                    statHtmlALXBcp.IsSplitted = true;
                    statHtmlALXBcp.CreateCorporate = true;
                    checkErrRslt = statHtmlALXBcp.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlALXBcp = null;
                    break;

                case 402: //[139] CMB Coronation Merchant Bank  >> Credit Emails 30/15/m
                    //chkDontPrompt.Checked = true;
                    clsStatHtmlCMB statHtmlCMB = new clsStatHtmlCMB();

                    statHtmlCMB.isSplitted = false;
                    //statHtmlCMB.PaymentSystem = "Visa";
                    statHtmlCMB.BillingCycle = "SD 15";

                    statHtmlCMB.emailFromName = "Coronation Merchant Bank Client Care"; //"mamr@network.com.eg""Nermine.Bahaa-Eldin@socgen.com"
                    statHtmlCMB.emailFrom = "Statement@emp-group.com"; //"clientcare@coronationmb.com";//mamr@network.com.eg
                    statHtmlCMB.bankWebLink = "www.coronationmb.com";
                    statHtmlCMB.bankLogo = @"D:\pC#\ProjData\Statement\CMB\CoronationMB_Boiler_Plate.jpg";
                    //statHtmlCMB.backGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
                    statHtmlCMB.visaLogo = @"D:\pC#\ProjData\Statement\CMB\CMBlogo.jpg";
                    //statHtmlCMB.bottomBanner = @"D:\pC#\ProjData\Statement\UNBN\bottom.jpg";
                    statHtmlCMB.waitPeriod = 3000;
                    statHtmlCMB.HasAttachement = true;
                    //statHtmlCMB.isReward = true;
                    //statHtmlCMB.rewardCond = "('Credit Card Rewards Program')";
                    statHtmlCMB.setFrm = this;
                    checkErrRslt = statHtmlCMB.Statement(txtFileName.Text + "_WaitReply\\", bankName, bankCode, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlCMB = null;
                    break;


                //FBPG
                case 500:   //[146] Fidelity Bank Ghana Limited  >> PrePaid email 1/m
                case 501:   //[146] Fidelity Bank Ghana Limited  >> Credit email 1/m
                    clsStatHtmlFBPG statHtmlFBPG = new clsStatHtmlFBPG();
                    statHtmlFBPG.emailFromName = "Fidelity Bank - Estatement";
                    statHtmlFBPG.emailFrom = "cards@myfidelitybank.net";
                    statHtmlFBPG.bankWebLink = "www.fidelitybank.com.gh";
                    statHtmlFBPG.bankLogo = @"D:\pC#\ProjData\Statement\FBPG\FBPGLogo.jpg";
                    statHtmlFBPG.strbottomBanner = @"D:\pC#\ProjData\Statement\FBPG\FBPGFooter.jpg";
                    statHtmlFBPG.waitPeriod = 500;
                    statHtmlFBPG.setFrm = this;
                    statHtmlFBPG.HasAttachement = true;
                    statHtmlFBPG.IsSplitted = true;
                    EventLog.WriteEntry("STMT", "frm statement file extn", EventLogEntryType.Error);

                    if (pCmbProducts == 500)
                    {
                        //statHtmlFBPG.curFileName = "Fidelity Bank Prepaid Statement ";
                        statHtmlFBPG.productCond = "('Visa Classic Prepaid Card','Visa GoG Staff and Travel Prepaid Card','MC Business Prepaid','MC Personal Prepaid')";
                    }
                    if (pCmbProducts == 501)
                    {
                        //statHtmlFBPG.curFileName = "Fidelity Bank Credit Card Statement ";
                        //statHtmlFBPG.productCond = "('Mastercard Credit')";
                        statHtmlFBPG.productCond = "('Visa Platinum Credit GHS', 'Visa Signature Credit GHS', 'Visa Signature Credit USD')";
                    }

                    checkErrRslt = statHtmlFBPG.Statement(txtFileName.Text + "_WaitReply\\", bankName, 146, strFileName, stmntDate, stmntClientEmail, appendData, reportFleName);
                    statHtmlFBPG = null;
                    break;


                    //case 113:   // 50) ICBG Intercontinental Bank Ghana Limited Prepaid >> PDF 1/m
                    //    clsStatTxtLblDb_ICBG statTxtLblDb_Prepaid = new clsStatTxtLblDb_ICBG();
                    //    statTxtLblDb_Prepaid.setFrm = this;
                    //    //if (pCmbProducts == 100)
                    //    //    statTxtLblDb.emailService = true;
                    //    statTxtLblDb_Prepaid.PrepaidCondition = "('Visa Classic Prepaid','Visa Classic Prepaid Gift')";
                    //    checkErrRslt = statTxtLblDb_Prepaid.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "ABP_Statement_File.txt"
                    //    statTxtLblDb_Prepaid = null;
                    //    break;

                    //case 114: // 1] NSGB MasterCard Salary Prepaid >> PDF 1/m
                    //    //clsMaintainData maintainDataNSGB = new clsMaintainData();
                    //    //maintainDataNSGB.makeMainCardNum(bankCode);
                    //    //maintainDataNSGB.makeBranchAsMainCard(bankCode);
                    //    //maintainDataNSGB = null;
                    //    clsStatement_ExportRpt stmntNSGB = new clsStatement_ExportRpt();
                    //    stmntNSGB.setFrm = this;
                    //    stmntNSGB.mantainBank(bankCode);
                    //    stmntNSGB.SplitByCompany(txtFileName.Text, strStatementType, bankCode, strFileName, reportFleName, ExportFormatType.PortableDocFormat, stmntDate, stmntType, appendData);
                    //    stmntNSGB.CreateZip();
                    //    stmntNSGB = null;
                    //    break;

            } // end switch
            //lblStatus.ForeColor = Color.Blue;
            //lblStatus.Text = "Statement File Creation Done for " + strStatementType + " \r\n" + checkErrRslt;
            //txtRunResult.Text = lblStatus.Text + "\r\n\r\n" + basText.Left(txtRunResult.Text, 2000);
            //btnSaveStatement.Enabled = true;
            //btnCancelGeneration.Enabled = false;

            BeginInvoke(setStatusDelegate, new object[] { bankCode, "Statement File Creation Done for " + strStatementType });//strStatementType

            checkErrRslt = string.Empty;
        }
        catch (Exception ex)  //(Exception ex)  //
        {
            clsBasErrors.catchError(ex);
        }
        //finally
        //    {
        //    }
    }


    private void setCurrentBank()
    {
        clsSessionValues.mainDbSchema = txtDbSchema.Text;
        clsSessionValues.mainTable = txtTblMaster.Text;
        clsSessionValues.detailTable = txtTblDetail.Text;
        clsSessionValues.Prompt4Msg = !chkDontPrompt.Checked;
        clsSessionValues.basPath = txtFileName.Text;
        //clsSessionValues.isUpdateDatble = clsStatementSummary.isUpdateDatble = chkUpdateSummary.Checked;
        clsStatementSummary.isUpdatedatble = chkUpdateSummary.Checked;
        clsStatementSummary.isUpdatable = chkUpdateSummaryPart.Checked;//false
        clsSessionValues.statGenAfterMonth = chkStatGenAfterMonth.Checked == true ? 1 : 0;
        bankCode = 0; bankName = ""; strStatementType = ""; stmntType = "Credit"; // Corporate Reward Debit

        strFileName = "_Statement_File_";
        stmntDate = datStmntData.Value; // = DateTime.Now.AddMonths(-1); bug1
        appendData = chkAppendData.Checked;
        stmntClientEmail = "";//ClientsEmails
        whereCond = string.Empty;
        statPeriod = "01";
        reportFleName = Application.StartupPath + @"\Reports\";
        //curIndx = chkLstProducts.SelectedIndex;
        switch (curIndx)//  chkLstProducts.SelectedIndex
        {
            //case 0:   // Common
            //    reportFleName += "Statement_Common_English.rpt";
            //    goto case 72;
            //case 5:   // Statement_Common_English_Nice1 
            //    reportFleName += "Statement_Common_English_Nice1.rpt";
            //    goto case 72;
            //case 28:   // Common Statement English Arabic
            //    reportFleName += "Statement_Common_English_Arabic.rpt";
            //    goto case 72;
            //case 29:   // Common Statement Portuguese
            //    reportFleName += "Statement_Common_Portuguese.rpt";
            //    goto case 72;
            //case 30:   // Statement Common English Nice2 
            //    reportFleName += "Statement_Common_English_Nice2.rpt";
            //    goto case 72;
            //case 32:   // 
            //    reportFleName += "Statement_Common_English_Debit_EMP.rpt";
            //    goto case 72;
            //case 33:   // 
            //case 34:   // 
            //case 38:   // 
            //case 39:   // 
            //case 59:   // Test Statement
            //case 67:   // 
            //case 72:   // Common Corporate Detail English Text - No Charge
            //bankCode = Convert.ToInt32(txtBankCode.Text); bankName = txtBank.Text;
            //strStatementType = txtBank.Text;
            //stmntMail = txtBank.Text;
            //break;

            case 1:   // clsStatementNSGBcredit Splitted by product
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit";
                strStatementType = "QNB_ALAHLI_Credit";
                stmntMail = "QNB_ALAHLI";
                break;

            case 2:   // clsStatementNSGBbusiness
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Corporate";
                strStatementType = "QNB_ALAHLI_Business";
                stmntMail = "QNB_ALAHLI";
                break;

            case 4000: // NSGB Mastercard Business - SME 
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Corporate";
                strStatementType = "Visa_Platinum_MVSE";
                //strFileName = "_Statement_File_";//_Business
                stmntMail = "QNB_ALAHLI";
                break;

            case 4001: // NSGB Mastercard Business - SME 
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Corporate";
                strStatementType = "Visa_Platinum_Corporate";
                //strFileName = "_Statement_File_";//_Business
                stmntMail = "QNB_ALAHLI";
                break;


            case 4002: // NSGB Mastercard Business - SME 
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Corporate";
                strStatementType = "Visa_Platinum_SME";
                //strFileName = "_Statement_File_";//_Business
                stmntMail = "QNB_ALAHLI";
                break;

            case 4003: // NSGB Mastercard Business - SME 
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Corporate";
                strStatementType = "VISA_Platinum_B2B";
                //strFileName = "_Statement_File_";//_Business
                stmntMail = "QNB_ALAHLI";
                break;

            case 3:   // 3) NBK VISA Classic >> Text 1/m
                bankCode = 3; bankName = "NBK"; stmntType = "Credit_Classic";
                strStatementType = "NBK_Credit_Classic";
                stmntMail = "NBK";
                break;

            case 4:   // 5) AIB Raw Data VISA USD>> Text 1/m
                bankCode = 5; bankName = "AIB"; stmntType = "Credit_VISA_USD";
                strStatementType = "AIB_VISA_USD";// Raw Data
                stmntMail = "AIB";
                break;

            case 64:   // 5) AIB Raw Data VISA EUR>> Text 1/m
                bankCode = 5; bankName = "AIB"; stmntType = "Credit_VISA_EUR";
                strStatementType = "AIB_VISA_EUR";// Raw Data
                stmntMail = "AIB";
                break;

            case 187:   // 5) AIB Raw Data VISA EGP>> Text 1/m
                bankCode = 5; bankName = "AIB"; stmntType = "Credit_VISA_EGP";
                strStatementType = "AIB_VISA_EGP";// Raw Data
                stmntMail = "AIB";
                break;

            case 237:   // 5) AIB Raw Data MasterCard USD>> Text 1/m
                bankCode = 5; bankName = "AIB"; stmntType = "Credit_MasterCard_USD";
                strStatementType = "AIB_MasterCard_USD";// Raw Data
                stmntMail = "AIB";
                break;

            case 296:   // 5) AIB Raw Data Corporate>> Text 1/m
                bankCode = 5; bankName = "AIB"; stmntType = "Corporate";
                strStatementType = "AIB_Corporate";// Raw Data
                stmntMail = "AIB";
                break;

            case 238: //109) CGBK COMPAGNIE GENERALE DE BANQUE LTD >> MasterCard Credit Text 1/m
                bankCode = 109; bankName = "CGBK"; stmntType = "MasterCard_Credit";
                strStatementType = "CGBK_MasterCard_Credit";
                stmntMail = "CGBK";
                break;

            case 239: //109) CGBK COMPAGNIE GENERALE DE BANQUE LTD >> MasterCard Prepaid Text 1/m
                bankCode = 109; bankName = "CGBK"; stmntType = "MasterCard_Prepaid";
                strStatementType = "CGBK_MasterCard_Prepaid";
                stmntMail = "CGBK";
                break;

            case 275: //33) ZENG Zenith Bank (Ghana) Limited >> MasterCard Prepaid text 16/m
                bankCode = 33; bankName = "ZENG"; stmntType = "MasterCard_Prepaid";
                strStatementType = "ZENG_MasterCard_Prepaid";
                stmntMail = "ZENG";
                break;

            case 240: //110) HBLN Heritage Banking Company Nigeria >> MasterCard Credit Text 1/m
                bankCode = 110; bankName = "HBLN"; stmntType = "MasterCard_Credit";
                strStatementType = "HBLN_MasterCard_Credit";
                stmntMail = "HBLN";
                break;

            case 246: //110) HBLN Heritage Banking Company Nigeria >> MasterCard Prepaid Text 1/m
                bankCode = 110; bankName = "HBLN"; stmntType = "MasterCard_Prepaid";
                strStatementType = "HBLN_MasterCard_Prepaid";
                stmntMail = "HBLN";
                break;

            case 241: //[110] HBLN Heritage Banking Company Nigeria >> MasterCard Credit Emails 1/m
                bankCode = 110; bankName = "HBLN"; stmntType = "MasterCard_Credit_Emails";
                strStatementType = "HBLN_MasterCard_Credit_ClientsEmails";
                stmntMail = "HBLN";
                stmntClientEmail = "MasterCard_Credit_ClientsEmails";
                strFileName = "_";
                break;

            case 247: //[110] HBLN Heritage Banking Company Nigeria >> MasterCard Prepaid Emails 1/m
                bankCode = 110; bankName = "HBLN"; stmntType = "MasterCard_Prepaid_Emails";
                strStatementType = "HBLN_MasterCard_Prepaid_ClientsEmails";
                stmntMail = "HBLN";
                stmntClientEmail = "MasterCard_Prepaid_ClientsEmails";
                strFileName = "_";
                break;

            case 255: //[113] GTBU Guaranty trust bank Uganada >> MasterCard Debit Emails 1/m
                bankCode = 113; bankName = "GTBU"; stmntType = "MasterCard_Debit_Emails";
                strStatementType = "GTBU_MasterCard_Debit_ClientsEmails";
                stmntMail = "GTBU";
                stmntClientEmail = "MasterCard_Debit_ClientsEmails";
                strFileName = "_";
                break;

            case 6:   // clsStatementBNP
                bankCode = 10; bankName = "BNP"; stmntType = "Credit";
                strStatementType = "BNP_Credit";
                stmntMail = "BNP";
                break;

            case 7:   // 7) BAI Products >> PDF 1/m
                bankCode = 7; bankName = "BAI"; stmntType = "Credit";
                strStatementType = "BAI_Credit";
                stmntMail = "BAI";
                reportFleName += "Statement_BAI_Portuguese_English.rpt";
                break;

            case 464:   // 7) BAI Visa Classic >> PDF 1/m
                bankCode = 7; bankName = "BAI"; stmntType = "Credit";
                strStatementType = "BAI_Credit_Classic";
                stmntMail = "BAI";
                reportFleName += "Statement_BAI_Portuguese_Classic.rpt";
                break;

            case 465:   // 7) BAI Visa Gold >> PDF 1/m
                bankCode = 7; bankName = "BAI"; stmntType = "Credit";
                strStatementType = "BAI_Credit_Gold";
                stmntMail = "BAI";
                reportFleName += "Statement_BAI_Portuguese_Gold.rpt";
                break;

            case 555:   //[41] BDCA Banque Du Caire Classic  >> Credit Emails 5/m
                bankCode = 41; bankName = "BDCA"; stmntType = "Classic_Credit_PDF";
                strStatementType = "BDCA_Credit_Classic_PDF";
                stmntMail = "BDCA";
                stmntClientEmail = "Credit_Classic_PDF";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "5";// DateTime.Now;
                reportFleName += "BDCAStatementEmailingClassic.rpt";
                break;

            case 466:   // 7) BAI Visa Platinum >> PDF 1/m
                bankCode = 7; bankName = "BAI"; stmntType = "Credit";
                strStatementType = "BAI_Credit_Platinum";
                stmntMail = "BAI";
                reportFleName += "Statement_BAI_Portuguese_Platinum.rpt";
                break;

            case 127: // 7) BAI Prepaid >> PDF 1/m
                bankCode = 7; bankName = "BAI"; stmntType = "Prepaid";
                strStatementType = "BAI_Prepaid";
                stmntMail = "BAI";
                reportFleName += "Statement_BAI_Portuguese_Debit.rpt";
                break;

            case 8:   //29) UBA United Bank for Africa Plc Nigeria PDF 
                bankCode = 29; bankName = "UBA"; stmntType = "Credit";
                strStatementType = "UBA_Credit";
                stmntMail = "UBA";
                //reportFleName += "Statement_Common_English_Nice1.rpt";
                reportFleName += "Statement_UBA_Credit.rpt";
                break;

            case 436:   //29) UBA United Bank for Africa Plc Nigeria PDF 
                bankCode = 29; bankName = "UBA"; stmntType = "Credit_15";
                strStatementType = "UBA_Credit_15";
                stmntMail = "UBA";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                //reportFleName += "Statement_Common_English_Nice1.rpt";
                reportFleName += "Statement_UBA_Credit.rpt";
                break;

            case 9:   // 
                bankCode = 14; bankName = "BIC"; stmntType = "Credit";
                strStatementType = "BIC_Credit";
                stmntMail = "BIC";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                reportFleName += "Statement_BIC_Portuguese_English.rpt";
                break;

            case 10:   // 16) ABP Access Bank Plc with Label Credit >> Text 1/m
                bankCode = 16; bankName = "ABP"; stmntType = "Credit";//
                strStatementType = "ABP_Credit";
                stmntMail = "ABP";
                break;

            case 211:   //[16] ABP Access Bank Plc  >> Credit Emails 1/m
                bankCode = 16; bankName = "ABP"; stmntType = "Credit_Emails";
                strStatementType = "ABP_Credit_ClientsEmails";
                stmntMail = "ABP";
                stmntClientEmail = "Credit_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_ABP_Credit.rpt";
                break;

            case 230:   //[98] GTBK Guaranty trust bank Kenya  >> Credit Emails 15/m
                bankCode = 98; bankName = "GTBK"; stmntType = "Credit_Emails";
                strStatementType = "GTBK_Credit_ClientsEmails";
                stmntMail = "GTBK";
                stmntClientEmail = "Credit_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                reportFleName += "Statement_GTBK_Credit.rpt";
                break;

            case 257:   //[98] GTBK Guaranty trust bank Kenya  >> Prepaid Emails 15/m
                bankCode = 98; bankName = "GTBK"; stmntType = "Prepaid_Emails";
                strStatementType = "GTBK_Prepaid_ClientsEmails";
                stmntMail = "GTBK";
                stmntClientEmail = "Prepaid_ClientsEmails";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                reportFleName += "Statement_GTBK_Prepaid.rpt";
                break;

            case 225:   //[16] ABP Access Bank Plc  >> Corporate Cardholder Emails 1/m
                bankCode = 16; bankName = "ABP"; stmntType = "Corporate_Cardholder_Emails";
                strStatementType = "ABP_Corporate_Cardholder_ClientsEmails";
                stmntMail = "ABP";
                stmntClientEmail = "Corporate_Cardholder_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_ABP_Corporate_Cardholder.rpt";
                break;

            case 11:   // clsStatementAAIB
                bankCode = 6; bankName = "AAIB"; stmntType = "Credit";
                strStatementType = "AAIB_Credit";
                stmntMail = "AAIB";
                break;

            case 189: //6) AAIB Arab African International Bank >> Prepaid Text 1/m
                bankCode = 6; bankName = "AAIB"; stmntType = "Prepaid";
                strStatementType = "AAIB_Prepaid";
                stmntMail = "AAIB";
                break;

            case 12:   // clsStatementBMSRsave
                bankCode = 4; bankName = "BMSR"; stmntType = "Debit";
                strStatementType = "BMSR_Debit";
                //strFileName = "_Statement_File_";//_Save
                stmntMail = "BMSR_Debit";
                break;

            case 13:   // clsStatementBAI
                bankCode = 25; bankName = "AUB"; stmntType = "Credit";
                strStatementType = "AUB_Credit";
                stmntMail = "AUB";
                //>stmntDate = datStmntData.Value; // DateTime.Now;
                break;

            case 14:   // 22) BPC BANCO DE POUPANCA E CREDITO, SARL Credit >> PDF 1/m
                bankCode = 22; bankName = "BPC"; stmntType = "Credit";
                strStatementType = "BPC_Credit";
                stmntMail = "BPC";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                reportFleName += "Statement_BPC_Portuguese.rpt";
                break;

            case 191:   // 22) BPC BANCO DE POUPANCA E CREDITO, SARL Debit >> Debit PDF 1/m
                bankCode = 22; bankName = "BPC"; stmntType = "Debit";
                strStatementType = "BPC_Debit";
                stmntMail = "BPC";
                reportFleName += "Statement_BPC_Portuguese_Debit.rpt";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 15:   // clsStatementBMSRcredit
                bankCode = 4; bankName = "BMSR"; stmntType = "Credit";
                strStatementType = "BMSR_Credit_Moga";
                stmntMail = "BMSR_Credit";
                //strFileName = "_Statement_File_";//_Moga
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                //whereCond = " contracttype = 'Visa Credit Moga (441194)'";
                whereCond = " contracttype in ('Visa Credit Moga - Secured','Visa Credit Moga - UnSecured')";
                break;

            case 16:   // BMSR Bank Misr Classic >> Text 1/m
                bankCode = 4; bankName = "BMSR"; stmntType = "Credit";
                strStatementType = "BMSR_Credit_Classic";
                stmntMail = "BMSR_Credit";
                //strFileName = "_Statement_Classic_File_";
                whereCond = " contracttype in ('Visa Credit Classic - Secured','Visa Credit Classic - UnSecured')";
                break;

            case 17:   // 21) ZEN Zenith Bank PLC >> MS Access MDB 16/m
                bankCode = 21; bankName = "ZEN"; stmntType = "Credit";
                strStatementType = "ZEN_Credit";
                stmntMail = "ZEN";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                //stmntDate = DateTime.Now;
                break;

            case 430:   // 11) GTB  >> MS Access MDB EOM
                bankCode = 11; bankName = "GTB_MDB"; stmntType = "Credit";
                strStatementType = "GTB_MDB_Credit";
                stmntMail = "GTB_MDB";
                stmntDate = datStmntData.Value; //statPeriod = "15"; // DateTime.Now;
                //stmntDate = DateTime.Now;
                break;

            case 66:   // 21) ZEN Zenith Bank PLC >> Default Text for Email 16/m
                bankCode = 21; bankName = "ZEN"; stmntType = "Credit";
                strStatementType = "ZEN_Credit_TextEmail";
                stmntMail = "ZEN";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                //stmntDate = DateTime.Now;
                break;

            case 41:   // 21] ZEN Zenith Bank PLC >> Email 16/m
                bankCode = 21; bankName = "ZEN"; stmntType = "Credit_Emails";
                strStatementType = "ZEN_ClientsEmails";//ZEN_Credit_ClientsEmails
                stmntMail = "ZEN";
                stmntClientEmail = "ClientsEmails";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                //stmntDate = DateTime.Now;
                strFileName = "_";
                break;

            case 18:   // 15) SSB SG-SSB Ltd Corporate >> PDF 1/m
                bankCode = 15; bankName = "SSB"; stmntType = "Corporate";
                strStatementType = "SSB_Corporate";
                stmntMail = "SSB_Corporate";
                strFileName = "_Statement_";
                break;

            //case 19:   // clsStatementNSGBcredit
            //    bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Individual";
            //    strStatementType = "QNB_ALAHLI_Individual";
            //    //strFileName = "_Statement_File_";//_Credit
            //    stmntMail = "QNB_ALAHLI";
            //    break;

            case 20:   // BCA
                bankCode = 27; bankName = "BCA"; stmntType = "Credit";
                strStatementType = "BCA_Credit";
                stmntMail = "BCA";
                reportFleName += "Statement_BCA_Credit_Portuguese_English.rpt";
                //reportFleName += "Statement_BCA.rpt";
                break;

            //case 21:   // BCNS Credit
            //    bankCode = 19; bankName = "BCNS"; stmntType = "Credit_Business";
            //    strStatementType = "BCNS_Business";
            //    stmntMail = "BCNS";
            //    reportFleName += "Statement_BCNS.rpt";
            //    break;

            //case 22:   // BCNS Debit
            //    bankCode = 19; bankName = "BCNS"; stmntType = "Debit_PreTravel";
            //    strStatementType = "BCNS_PreTravel";
            //    stmntMail = "BCNS";
            //    reportFleName += "Statement_Debit_BCNS.rpt";
            //    break;

            case 23:   // BICV
                bankCode = 20; bankName = "BICV"; stmntType = "Credit";
                strStatementType = "BICV_Credit_PDF";
                stmntMail = "BICV";
                reportFleName += "Statement_BICV.rpt";
                break;

            case 165:   // BICV Credit  Text
                bankCode = 20; bankName = "BICV"; stmntType = "Credit";
                strStatementType = "BICV_Credit_Text";
                stmntMail = "BICV";
                break;

            case 24:   // 28) CECV CAIXA ECONOMICA DE CABO VERDE >> Credit Classic >> PDF 1/m
                bankCode = 28; bankName = "CECV"; stmntType = "Credit";
                strStatementType = "CECV_Credit_PDF";
                stmntMail = "CECV";
                reportFleName += "Statement_CECV.rpt";
                break;

            case 163:   // 28) CECV CAIXA ECONOMICA DE CABO VERDE >> Credit Classic >> Text 1/m
                bankCode = 28; bankName = "CECV"; stmntType = "Credit";
                strStatementType = "CECV_Credit_Text";
                stmntMail = "CECV";
                break;

            case 91:   // 28) CECV CAIXA ECONOMICA DE CABO VERDE >> Debit - Prepaid >> PDF 1/m
                bankCode = 28; bankName = "CECV"; stmntType = "Debit";
                strStatementType = "CECV_Debit_PDF";
                stmntMail = "CECV";
                reportFleName += "Statement_CECV_Debit.rpt";
                break;

            case 164:   // 28) CECV CAIXA ECONOMICA DE CABO VERDE >> Debit - Prepaid >> Text 1/m
                bankCode = 28; bankName = "CECV"; stmntType = "Debit";
                strStatementType = "CECV_Debit_Text";
                stmntMail = "CECV";
                break;

            case 25:   // BICV Debit   pre-paid 
                bankCode = 20; bankName = "BICV"; stmntType = "Debit";
                strStatementType = "BICV_Debit_PDF";
                stmntMail = "BICV";
                reportFleName += "Statement_Debit_BICV.rpt";
                break;

            case 166:   // BICV Debit   pre-paid  Text
                bankCode = 20; bankName = "BICV"; stmntType = "Debit";
                strStatementType = "BICV_Debit_Text";
                stmntMail = "BICV";
                break;

            case 26:   // 23) Suez Canal Bank(SCB) Gold Standard>> PDF 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "Credit_Gold_Standard";
                strStatementType = "Suez_Credit_Gold_Standard";
                stmntMail = "Suez";
                whereCond = " and contracttype = 'Visa Gold - Standard'";
                reportFleName += "Statement_Suez_Canal_Bank_SCB.rpt";
                break;

            case 27:   // 23) Suez Canal Bank(SCB) Classic Standard>> PDF 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "Credit_Classic_Standard";
                strStatementType = "Suez_Credit_Classic_Standard";
                stmntMail = "Suez";
                whereCond = " and contracttype = 'Visa Classic - Standard'";
                reportFleName += "Statement_Suez_Canal_Bank_SCB.rpt";
                break;

            case 196:   // 23) Suez Canal Bank(SCB) MC Gold Standard>> PDF 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "MC_Credit_Titanium_Standard";
                strStatementType = "Suez_MC_Credit_Titanium_Standard";
                stmntMail = "Suez";
                whereCond = " and contracttype = 'MC Titanium - Standard'";
                reportFleName += "Statement_Suez_Canal_Bank_SCB_MC.rpt";
                break;

            case 197:   // 23) Suez Canal Bank(SCB) MC Classic Standard>> PDF 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "MC_Credit_Classic_Standard";
                strStatementType = "Suez_MC_Credit_Classic_Standard";
                stmntMail = "Suez";
                whereCond = " and contracttype = 'MC Classic - Standard'";
                reportFleName += "Statement_Suez_Canal_Bank_SCB_MC.rpt";
                break;

            case 242:   // 23) Suez Canal Bank(SCB) Gold Staff>> PDF 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "Credit_Gold_Staff";
                strStatementType = "Suez_Credit_Gold_Staff";
                stmntMail = "Suez";
                whereCond = " and contracttype = 'Visa Gold - Staff'";
                reportFleName += "Statement_Suez_Canal_Bank_SCB.rpt";
                break;

            case 243:   // 23) Suez Canal Bank(SCB) Classic Staff>> PDF 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "Credit_Classic_Staff";
                strStatementType = "Suez_Credit_Classic_Staff";
                stmntMail = "Suez";
                whereCond = " and contracttype = 'Visa Classic - Staff'";
                reportFleName += "Statement_Suez_Canal_Bank_SCB.rpt";
                break;

            case 244:   // 23) Suez Canal Bank(SCB) MC Titanium Staff>> PDF 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "MC_Credit_Titanium_Staff";
                strStatementType = "Suez_MC_Credit_Titanium_Staff";
                stmntMail = "Suez";
                whereCond = " and contracttype = 'MC Titanium - Staff'";
                reportFleName += "Statement_Suez_Canal_Bank_SCB_MC.rpt";
                break;

            case 245:   // 23) Suez Canal Bank(SCB) MC Classic Staff>> PDF 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "MC_Credit_Classic_Staff";
                strStatementType = "Suez_MC_Credit_Classic_Staff";
                stmntMail = "Suez";
                whereCond = " and contracttype = 'MC Classic - Staff'";
                reportFleName += "Statement_Suez_Canal_Bank_SCB_MC.rpt";
                break;

            case 426:   // 23) Suez Canal Bank(SCB) MC Platnium Standard>> PDF 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "MC_Credit_Platnium_Standard";
                strStatementType = "Suez_MC_Credit_Platnium_Standard";
                stmntMail = "Suez";
                whereCond = " and contracttype = 'MC Platnium - Standard'";
                reportFleName += "Statement_Suez_Canal_Bank_SCB_MC.rpt";
                break;

            case 428:   // 23) Suez Canal Bank(SCB) MC Platnium Staff>> PDF 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "MC_Credit_Platnium_Staff";
                strStatementType = "Suez_MC_Credit_PlatniumStaff";
                stmntMail = "Suez";
                whereCond = " and contracttype = 'MC Platnium - Staff'";
                reportFleName += "Statement_Suez_Canal_Bank_SCB_MC.rpt";
                break;

            case 249:   // 23) Suez Canal Bank(SCB) Gold Standard>> Text 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "Credit_Gold_Standard_Text";
                strStatementType = "Suez_Credit_Gold_Standard_Text";
                stmntMail = "Suez";
                break;

            case 260:   // 23) Suez Canal Bank(SCB) Classic Standard>> Text 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "Credit_Classic_Standard_Text";
                strStatementType = "Suez_Credit_Classic_Standard_Text";
                stmntMail = "Suez";
                break;

            case 261:   // 23) Suez Canal Bank(SCB) MC Gold Standard>> PDF 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "MC_Credit_Titanium_Standard_Text";
                strStatementType = "Suez_MC_Credit_Titanium_Standard_Text";
                stmntMail = "Suez";
                break;

            case 262:   // 23) Suez Canal Bank(SCB) MC Classic Standard>> Text 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "MC_Credit_Classic_Standard_Text";
                strStatementType = "Suez_MC_Credit_Classic_Standard_Text";
                stmntMail = "Suez";
                break;

            case 263:   // 23) Suez Canal Bank(SCB) Gold Staff>> Text 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "Credit_Gold_Staff_Text";
                strStatementType = "Suez_Credit_Gold_Staff_Text";
                stmntMail = "Suez";
                break;

            case 264:   // 23) Suez Canal Bank(SCB) Classic Staff>> Text 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "Credit_Classic_Staff_Text";
                strStatementType = "Suez_Credit_Classic_Staff_Text";
                stmntMail = "Suez";
                break;

            case 265:   // 23) Suez Canal Bank(SCB) MC Titanium Staff>> Text 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "MC_Credit_Titanium_Staff_Text";
                strStatementType = "Suez_MC_Credit_Titanium_Staff_Text";
                stmntMail = "Suez";
                break;

            case 266:   // 23) Suez Canal Bank(SCB) MC Classic Staff>> Text 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "MC_Credit_Classic_Staff_Text";
                strStatementType = "Suez_MC_Credit_Classic_Staff_Text";
                stmntMail = "Suez";
                break;

            case 427:   // 23) Suez Canal Bank(SCB) MC Platnium Standard>> Text 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "MC_Credit_Platnium_Standard_Text";
                strStatementType = "Suez_MC_Credit_Platnium_Standard_Text";
                stmntMail = "Suez";
                break;

            case 429:   // 23) Suez Canal Bank(SCB) MC Platnium Staff>> Text 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "MC_Credit_Platnium_Staff_Text";
                strStatementType = "Suez_MC_Credit_Platnium_Staff_Text";
                stmntMail = "Suez";
                break;

            case 7100:   // 23) Suez Canal Bank(SCB) MC Titanium Staff>> PDF 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "Suez_Corperate_PDF";
                strStatementType = "Suez_Corperate_PDF";
                stmntMail = "Suez";
                whereCond = " and contracttype = 'MC Titanium - Staff'";
                reportFleName += "Statement_Suez_Canal_Bank_SCB_CP.rpt";
                break;

            case 31:   //32)UBAG UNITED BANK FOR AFRICA GHANA LIMITED
                bankCode = 32; bankName = "UBAG"; stmntType = "Credit";
                strStatementType = "UBAG_Credit";
                stmntMail = "UBAG";
                reportFleName += "Statement_UBAG.rpt";
                break;

            case 193:   //32)UBAG UNITED BANK FOR AFRICA GHANA LIMITED >> Debit PDF 1/m
                bankCode = 32; bankName = "UBAG"; stmntType = "Debit";
                strStatementType = "UBAG_Debit";
                stmntMail = "UBAG";
                reportFleName += "Statement_Debit_UBAG.rpt";
                break;

            case 35:   // clsStatementABPlbl Dual Currency
                bankCode = 16; bankName = "ABP"; stmntType = "Credit_Dual_Currency";//
                strStatementType = "ABP_Credit_Dual_Currency";
                stmntMail = "ABP";//ABP_Dual_Currency
                break;

            case 36:   // BNP Corporate 10   1
                bankCode = 10; bankName = "BNP"; stmntType = "Corporate";//C
                strStatementType = "BNP_Corporate";
                stmntMail = "BNP";
                break;

            case 37:   // 10) BNP Credit Reward >> Text 1/m
                bankCode = 10; bankName = "BNP"; stmntType = "Credit_Reward";//R
                strStatementType = "BNP_Credit_Reward";
                stmntMail = "BNP";
                break;

            case 442:   //33) ZENG Zenith Bank (Ghana) Limited >> MasterCard Prepaid PDF 16/m
                bankCode = 33; bankName = "ZENG"; stmntType = "Prepaid";
                strStatementType = "ZENG_Prepaid";
                stmntMail = "ZENG";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                whereCond = "";
                reportFleName += "Statement_DB_ZENG.rpt";
                break;

            case 40:   //31) NBS NBS Bank Limited >> PDF 16/m
                bankCode = 31; bankName = "NBS"; stmntType = "Credit";
                strStatementType = "NBS_Credit";
                stmntMail = "NBS";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                reportFleName += "Statement_NBS_Credit.rpt";
                break;

            case 42:   //33) ZENG Zenith Bank (Ghana) Limited >> Visa Classic Credit PDF 16/m
                bankCode = 33; bankName = "ZENG"; stmntType = "Visa_Credit_Classic";
                strStatementType = "ZENG_Visa_Credit_Classic";
                stmntMail = "ZENG";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                whereCond = " and contracttype = 'Visa Classic Credit'";
                reportFleName += "Statement_Credit_ZENG.rpt";
                break;

            case 281:   //33) ZENG Zenith Bank (Ghana) Limited >> Visa Classic Credit (Staff) PDF 16/m
                bankCode = 33; bankName = "ZENG"; stmntType = "Visa_Credit_Classic_Staff";
                strStatementType = "ZENG_Visa_Credit_Classic_Staff";
                stmntMail = "ZENG";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                whereCond = " and contracttype = 'Visa Classic Credit (Staff)'";
                reportFleName += "Statement_Credit_ZENG.rpt";
                break;

            case 282:   //33) ZENG Zenith Bank (Ghana) Limited >> MasterCard Credit PDF 16/m
                bankCode = 33; bankName = "ZENG"; stmntType = "MasterCard_Credit";
                strStatementType = "ZENG_MasterCard_Credit";
                stmntMail = "ZENG";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                whereCond = " and contracttype = 'MasterCard Credit'";
                reportFleName += "Statement_Credit_MC_ZENG.rpt";
                break;

            case 283:   //33) ZENG Zenith Bank (Ghana) Limited >> MasterCard Credit (Staff) PDF 16/m
                bankCode = 33; bankName = "ZENG"; stmntType = "MasterCard_Credit_Staff";
                strStatementType = "ZENG_MasterCard_Credit_Staff";
                stmntMail = "ZENG";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                whereCond = " and contracttype = 'MasterCard Credit (Staff)'";
                reportFleName += "Statement_Credit_MC_ZENG.rpt";
                break;

            case 43:   // clsStatementABPlbl Corporate 16
                bankCode = 16; bankName = "ABP"; stmntType = "Corporate";
                strStatementType = "ABP_Corporate";
                stmntMail = "ABP";
                break;

            case 44:   // ICB INTERCONTINENTAL BANK PLC >> Default Text 1/m
                bankCode = 40; bankName = "ICB"; stmntType = "Credit";//
                strStatementType = "ICB_Credit";
                stmntMail = "ICB";
                break;

            case 45:   //29) UBA United Bank for Africa Plc Nigeria >> Emails
                bankCode = 29; bankName = "UBA"; stmntType = "Credit_Emails";
                strStatementType = "UBA_ClientsEmails";
                stmntClientEmail = "ClientsEmails";
                reportFleName += "Statement_UBA_Credit.rpt";
                stmntMail = "UBA";
                strFileName = "_";
                break;

            case 435:   //29) UBA United Bank for Africa Plc Nigeria >> Emails
                bankCode = 29; bankName = "UBA"; stmntType = "Credit_Emails_15";
                strStatementType = "UBA_ClientsEmails_15";
                stmntClientEmail = "ClientsEmails_15";
                reportFleName += "Statement_UBA_Credit.rpt";
                stmntMail = "UBA";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                break;

            case 46:   //36) DBN Diamond Bank Nigeria >> Default Text 15/m
                bankCode = 36; bankName = "DBN"; stmntType = "Credit";
                strStatementType = "DBN_Credit";
                stmntMail = "DBN";// DBN
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 111:   //36) DBN Diamond Bank Nigeria >> VISA Platinum - ParkNShop Co-Brand Text 15/m
                bankCode = 36; bankName = "DBN"; stmntType = "Credit_ParkNShop";
                strStatementType = "DBN_Credit_ParkNShop";
                stmntMail = "DBN";// DBN
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 149:   //36) 36) DBN Diamond Bank Nigeria >> VISA Platinum - EXCO_VIP-Brand Text 16/m
                bankCode = 36; bankName = "DBN"; stmntType = "Credit_EXCO_VIP";
                strStatementType = "DBN_Credit_EXCO_VIP";
                stmntMail = "DBN";// DBN
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 220:   //36) 36) DBN Diamond Bank Nigeria >> MasterCard Credit Text 16/m
                bankCode = 36; bankName = "DBN"; stmntType = "MasterCard_Credit";
                strStatementType = "DBN_MasterCard_Credit";
                stmntMail = "DBN";// DBN
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 83:   //36) DBN Diamond Bank Nigeria - VIP 1,2,5 >> Reward Default Text 1/m
                bankCode = 36; bankName = "DBN"; stmntType = "Credit_VIP1-2-5";
                strStatementType = "DBN_Credit_VIP1-2-5";
                stmntMail = "DBN";
                break;

            case 97:   //36) DBN Diamond Bank Nigeria - VIP >> Reward Default Text 16/m
                bankCode = 36; bankName = "DBN"; stmntType = "Credit_VIP";
                strStatementType = "DBN_Credit_VIP";
                stmntMail = "DBN";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 84:   //36] DBN Diamond Bank Nigeria - VIP 1,2,5 >> Email 1/m
                bankCode = 36; bankName = "DBN"; stmntType = "Credit_VIP1-2-5_Emails"; //_VIP
                strStatementType = "DBN_VIP1-2-5_ClientsEmails";//_Credit_VIP
                stmntClientEmail = "VIP1-2-5_ClientsEmails";
                stmntMail = "DBN";//_Credit_VIP
                strFileName = "_";
                reportFleName += "Statement_ABP_DBN_VIP1_2_5.rpt";
                break;




            //case 48:   //36] DBN Diamond Bank Nigeria >> Email 16/m
            //    bankCode = 36; bankName = "DBN"; stmntType = "Credit_Emails";
            //    strStatementType = "DBN_Credit_ClientsEmails";//DBN_Credit_ClientsEmails
            //    stmntClientEmail = "Credit_ClientsEmails";
            //    stmntMail = "DBN";
            //    stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
            //    strFileName = "_";
            //    break;

            case 284:   //36] DBN Diamond Bank Nigeria >> Email Classic 16/m
                bankCode = 36; bankName = "DBN"; stmntType = "Credit_Classic_Emails";
                strStatementType = "DBN_Credit_Classic_ClientsEmails";//DBN_Credit_ClientsEmails
                stmntClientEmail = "Credit_Classic_ClientsEmails";
                stmntMail = "DBN";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                strFileName = "_";
                reportFleName += "Statement_ABP_DBN_Classic.rpt";
                break;

            case 285:   //36] DBN Diamond Bank Nigeria >> Email Gold 16/m
                bankCode = 36; bankName = "DBN"; stmntType = "Credit_Gold_Emails";
                strStatementType = "DBN_Credit_Gold_ClientsEmails";//DBN_Credit_ClientsEmails
                stmntClientEmail = "Credit_Gold_ClientsEmails";
                stmntMail = "DBN";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                strFileName = "_";
                reportFleName += "Statement_ABP_DBN_Gold.rpt";
                break;

            case 286:   //36] DBN Diamond Bank Nigeria >> Email Platinum 16/m
                bankCode = 36; bankName = "DBN"; stmntType = "Credit_Platinum_Emails";
                strStatementType = "DBN_Credit_Platinum_ClientsEmails";//DBN_Credit_ClientsEmails
                stmntClientEmail = "Credit_Platinum_ClientsEmails";
                stmntMail = "DBN";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                strFileName = "_";
                reportFleName += "Statement_ABP_DBN_Platinum.rpt";
                break;

            case 340:   //36] DBN Diamond Bank Nigeria >> Email Platinum-USD 16/m
                bankCode = 36; bankName = "DBN"; stmntType = "Credit_Platinum_USD_Emails";
                strStatementType = "DBN_Credit_Platinum_USD_ClientsEmails";//DBN_Credit_ClientsEmails
                stmntClientEmail = "Credit_Platinum_USD_ClientsEmails";
                stmntMail = "DBN";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                strFileName = "_";
                reportFleName += "Statement_ABP_DBN_Platinum_USD.rpt";
                break;

            case 98:   //36] DBN Diamond Bank Nigeria - VIP >> Email 16/m
                bankCode = 36; bankName = "DBN"; stmntType = "Credit_VIP_Emails";
                strStatementType = "DBN_VIP_ClientsEmails";//DBN_Credit_ClientsEmails
                stmntClientEmail = "VIP_ClientsEmails";
                stmntMail = "DBN";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                strFileName = "_";
                reportFleName += "Statement_ABP_DBN_VIP.rpt";
                break;

            case 112:   //36] DBN Diamond Bank Nigeria >> VISA Platinum - ParkNShop Co-Brand Email 16/m
                bankCode = 36; bankName = "DBN"; stmntType = "Credit_ParkNShop_Emails";
                strStatementType = "DBN_ParkNShop_ClientsEmails";//DBN_Credit_ClientsEmails
                stmntClientEmail = "ParkNShop_ClientsEmails";
                stmntMail = "DBN";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                strFileName = "_";
                reportFleName += "Statement_ABP_DBN_ParkNShop.rpt";
                break;

            case 150:   //[36] DBN Diamond Bank Nigeria >> VISA SPECIAL EXCO VIP - Privilege Email 16/m
                bankCode = 36; bankName = "DBN"; stmntType = "Credit_Privilege_Emails";
                strStatementType = "DBN_Privilege_ClientsEmails";//DBN_Credit_ClientsEmails
                stmntClientEmail = "Privilege_ClientsEmails";
                stmntMail = "DBN";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                strFileName = "_";
                reportFleName += "Statement_ABP_DBN_Privilege.rpt";
                break;

            case 221:   // 36) DBN Diamond Bank Nigeria >> MasterCard Credit Email 16/m
                bankCode = 36; bankName = "DBN"; stmntType = "MasterCard_Credit_Emails";
                strStatementType = "DBN_MasterCard_Credit_ClientsEmails";//DBN_Credit_ClientsEmails
                stmntClientEmail = "MasterCard_Credit_ClientsEmails";
                stmntMail = "DBN";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                strFileName = "_";
                reportFleName += "Statement_ABP_DBN_MasterCard.rpt";
                break;

            //case 47:   //36) DBN Diamond Bank Nigeria >> Default Text 16/m
            //  bankCode = 36; bankName = "DBN"; stmntType = "Credit";
            //  strStatementType = "DBN_Credit";
            //  stmntMail = "DBN";
            //  stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
            //  break;

            case 49:   // 41) BDCA Banque Du Caire Classic >> Text 1/m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_Classic";
                strStatementType = "BDCA_Credit_Classic";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 138: // 41) BDCA Banque Du Caire MasterCard Standard >> Text 1/m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_Standard";
                strStatementType = "BDCA_MasterCard_Credit_Standard";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;


            case 50:   // 12) BSIC Alwaha Bank Libya Inc (Oasis Bank) >> PDF 1/m
                bankCode = 12; bankName = "BSIC"; stmntType = "Debit";
                strStatementType = "BSIC_Debit";
                stmntMail = "BSIC";
                reportFleName += "Statement_BSIC.rpt";
                break;

            case 51:   // 25) AUB Ahli United Bank >> Text 1/m
                bankCode = 25; bankName = "AUB"; stmntType = "Credit";
                strStatementType = "AUB_Credit";
                stmntMail = "AUB";
                break;

            case 52:   // 22) BPC BANCO DE POUPANCA E CREDITO, SARL Corporate >> Default PDF 1/m
                bankCode = 22; bankName = "BPC"; stmntType = "Corporate";
                strStatementType = "BPC_Corporate";
                stmntMail = "BPC";
                //stmntDate = DateTime.Now;
                break;

            case 53:   // 39) MBL Mainstreet Bank Limited >> Default Text 1/m
                bankCode = 39; bankName = "MBL"; stmntType = "Credit";
                strStatementType = "MBL_Credit";
                stmntMail = "MBL";
                //stmntDate = DateTime.Now;
                break;

            case 54:   //1) NSGB Credit >> Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit_Emails";
                strStatementType = "QNB_ALAHLI_Credit_ClientsEmails";
                stmntClientEmail = "Credit_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                //stmntDate = datStmntData.Value; // DateTime.Now;
                strFileName = "_";
                break;

            case 55:   //1) NSGB Business >> Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Business_Emails";
                strStatementType = "QNB_ALAHLI_Business_ClientsEmails";
                stmntClientEmail = "Business_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                //stmntDate = datStmntData.Value; // DateTime.Now;
                strFileName = "_";
                break;

            case 236:   // 1) NSGB MC Business SME >> Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "MC_Business_SME_Emails";
                strStatementType = "QNB_ALAHLI_MC_Business_SME_ClientsEmails";
                stmntClientEmail = "MC_Business_SME_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                //stmntDate = datStmntData.Value; // DateTime.Now;
                strFileName = "_";
                break;

            case 443:   // 1) NSGB Visa Business B2B >> Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Visa_Business_B2B_Emails";
                strStatementType = "QNB_ALAHLI_Visa_Business_B2B_ClientsEmails";
                stmntClientEmail = "Visa_Business_B2B_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                //stmntDate = datStmntData.Value; // DateTime.Now;
                strFileName = "_";
                break;

            case 444:   // 1) NSGB Visa Business FEDCOC >> Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Visa_FEDCOC_Emails";
                strStatementType = "QNB_ALAHLI_Visa_FEDCOC_ClientsEmails";
                stmntClientEmail = "Visa_FEDCOC_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                //stmntDate = datStmntData.Value; // DateTime.Now;
                strFileName = "_";
                break;

            case 317:   // 1) NSGB VISA Business Corporate >> Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "VISA_Business_Emails";
                strStatementType = "QNB_ALAHLI_VISA_Business_ClientsEmails";
                stmntClientEmail = "VISA_Business_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                //stmntDate = datStmntData.Value; // DateTime.Now;
                strFileName = "_";
                break;

            case 3000:   // 1) NSGB VISA Business Corporate >> Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Visa_Platinum_MVSE_Emails";
                strStatementType = "QNB_ALAHLI_Visa_Platinum_MVSE_ClientsEmails";
                stmntClientEmail = "Visa_Platinum_MVSE_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                //stmntDate = datStmntData.Value; // DateTime.Now;
                strFileName = "_";
                break;

            case 3001:   // 1) NSGB VISA Business Corporate >> Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Visa_Platinum_B2B_Emails";
                strStatementType = "QNB_Visa_Platinum_B2B_ClientsEmails";
                stmntClientEmail = "Visa_Platinum_B2B_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                //stmntDate = datStmntData.Value; // DateTime.Now;
                strFileName = "_";
                break;

            case 3002:   // 1) NSGB VISA Business Corporate >> Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Visa_Platinum_Single_Limit_in_EGB_Emails";
                strStatementType = "QNB_Visa_Platinum_Single_Limit_in_EGB_ClientsEmails";
                stmntClientEmail = "Visa_Platinum_Single_Limit_in_EGB_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                //stmntDate = datStmntData.Value; // DateTime.Now;
                strFileName = "_";
                break;

            case 3003:   // 1) NSGB VISA Business Corporate >> Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Visa_Platinum_SME_Emails";
                strStatementType = "QNB_ALAHLI_Visa_Platinum_SME_ClientsEmails";
                stmntClientEmail = "Visa_Platinum_SME_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                //stmntDate = datStmntData.Value; // DateTime.Now;
                strFileName = "_";
                break;

            case 56:   //1) NSGB Credit >> Business Individual Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Individual_Emails";
                strStatementType = "QNB_ALAHLI_Individual_ClientsEmails";
                stmntClientEmail = "Individual_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                //stmntDate = datStmntData.Value; // DateTime.Now;
                strFileName = "_";
                break;

            case 57:   // 37) KBL Keystone Bank >> Default Text 5/m
                bankCode = 37; bankName = "KBL"; stmntType = "Credit";
                strStatementType = "KBL_Credit";
                stmntMail = "KBL";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 58:   // 41) BDCA Banque Du Caire Gold >> Text 5/m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_Gold";
                strStatementType = "BDCA_Credit_Gold";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05";// DateTime.Now;
                break;

            case 139: // 41) BDCA Banque Du Caire MasterCard Gold >> Text 1/m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_Gold";
                strStatementType = "BDCA_MasterCard_Credit_Gold";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 173: // 41) BDCA Banque Du Caire MasterCard Installment >> Text 5/m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_Installment";
                strStatementType = "BDCA_MasterCard_Credit_Installment";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 273: // 41) BDCA Banque Du Caire MasterCard Titanium >> Text 5m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_Titanium";
                strStatementType = "BDCA_MasterCard_Credit_Titanium";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 318: // 41) BDCA Banque Du Caire MasterCard World Elite >> Text 5/m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_World_Elite";
                strStatementType = "BDCA_MasterCard_Credit_World_Elite";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 513: // 41) BDCA Banque Du Caire MasterCard platinum >> Text 5m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_Platinum";
                strStatementType = "BDCA_MasterCard_Credit_Platinum";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 515: // 41) BDCA Banque Du Caire MasterCard Corporate >> Text 1m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_Corporate";
                strStatementType = "BDCA_MasterCard_Corporate";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 413:   // 41) BDCA Banque Du Caire Classic staff >> Text 1/m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_Classic_Staff";
                strStatementType = "BDCA_Credit_Classic_Staff";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 415: // 41) BDCA Banque Du Caire MasterCard Standard staff >> Text 1/m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_Standard_Staff";
                strStatementType = "BDCA_MasterCard_Credit_Standard_Staff";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;



            case 414:   // 41) BDCA Banque Du Caire Gold staff >> Text 5/m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_Gold_Staff";
                strStatementType = "BDCA_Credit_Gold_Staff";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05";// DateTime.Now;
                break;

            case 416: // 41) BDCA Banque Du Caire MasterCard Gold staff >> Text 1/m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_Gold_Staff";
                strStatementType = "BDCA_MasterCard_Credit_Gold_Staff";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 417: // 41) BDCA Banque Du Caire MasterCard Installment staff >> Text 5/m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_Installment_Staff";
                strStatementType = "BDCA_MasterCard_Credit_Installment_Staff";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 418: // 41) BDCA Banque Du Caire MasterCard Titanium staff >> Text 5m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_Titanium_Staff";
                strStatementType = "BDCA_MasterCard_Credit_Titanium_Staff";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 419: // 41) BDCA Banque Du Caire MasterCard World Elite staff >> Text 5/m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_World_Elite_Staff";
                strStatementType = "BDCA_MasterCard_Credit_World_Elite_Staff";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 514: // 41) BDCA Banque Du Caire MasterCard platinum Staff >> Text 5m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_Platinum_Staff";
                strStatementType = "BDCA_MasterCard_Credit_Platinum_Staff";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 516: // 41) BDCA_MasterCard_Credit_Corporate_Staff >> Text 1m
                bankCode = 41; bankName = "BDCA"; stmntType = "Credit_Corporate_Staff";
                strStatementType = "BDCA_MasterCard_Credit_Corporate_Staff";
                stmntMail = "BDCA";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 502: // [146] Fidelity Bank Ghana Limited  >> PrePaid text 1/m
                bankCode = 146; bankName = "FBPG"; stmntType = "PrePaid";
                strStatementType = "FBPG_PrePaid";
                stmntMail = "FBPG";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 503: // [146] Fidelity Bank Ghana Limited  >> Credit text 1/m
                bankCode = 146; bankName = "FBPG"; stmntType = "Credit";
                strStatementType = "FBPG_Credit";
                stmntMail = "FBPG";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;
            //case 59:   // Test Statement
            //  bankCode = 10; bankName = "BNP"; stmntType = "Corporate";//C
            //  strStatementType = "BNP_Corporate";
            //  stmntMail = "BNP";
            //  break;

            case 60:   // 3) NBK VISA Gold >> Text 1/m
                bankCode = 3; bankName = "NBK"; stmntType = "Credit_Gold";
                strStatementType = "NBK_Credit_Gold";
                stmntMail = "NBK";
                break;

            case 61:   // 16) ABP Access Bank Plc with Label Infinite Dual Currency >> Text 1/m
                bankCode = 16; bankName = "ABP"; stmntType = "Infinite";//
                strStatementType = "ABP_Infinite";
                stmntMail = "ABP";//ABP_Infinite
                break;

            case 471:   // 16) ABP Access Bank Plc with Label AMEX Dual Currency >> Text 1/m
                bankCode = 16; bankName = "ABP"; stmntType = "AMEX";//
                strStatementType = "ABP_AMEX";
                stmntMail = "ABP";//ABP_AMEX
                break;

            case 62:   //37) KBL Keystone Bank >> Emails 1/m
                bankCode = 37; bankName = "KBL"; stmntType = "Credit_Emails";
                strStatementType = "KBL_ClientsEmails";
                stmntClientEmail = "ClientsEmails";
                stmntMail = "KBL";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                strFileName = "_";
                break;

            case 63:   // 45) BOCD Bank Of Commerce and Development Debit >> PDF 1/m
                bankCode = 45; bankName = "BOCD"; stmntType = "Debit";
                strStatementType = "BOCD_Debit";
                stmntMail = "BOCD";
                reportFleName += "Statement_BOCD.rpt";
                break;

            case 65:   // 18) BOAB Bank of Africa Benin >> Debit >> PDF 1/m
                bankCode = 18; bankName = "BOAB"; stmntType = "Debit";
                strStatementType = "BOAB_Debit";
                stmntMail = "BOAB";
                reportFleName += "Statement_Debit_BOAB.rpt";
                break;

            case 68:   // 39) MBL Mainstreet Bank Limited >> Default Debit Text 1/m
                bankCode = 39; bankName = "MBL"; stmntType = "Debit";
                strStatementType = "MBL_Debit";
                stmntMail = "MBL";
                break;

            case 69:   // 27) BCA BANCO COMERCIAL DO ATLANTICO >> Debit >> Default Debit Text 1/m
                bankCode = 27; bankName = "BCA"; stmntType = "Debit";
                strStatementType = "BCA_Debit";
                stmntMail = "BCA";
                break;

            case 188:   // 5) AIB Debit >> Text 1/m
                bankCode = 5; bankName = "AIB"; stmntType = "Debit";
                strStatementType = "AIB_Debit";
                stmntMail = "AIB";
                break;

            case 303: // 5) AIB Prepaid EGP >> Prepaid Raw 1/m
                bankCode = 5; bankName = "AIB"; stmntType = "Prepaid";
                strStatementType = "AIB_Prepaid";
                stmntMail = "AIB";
                break;
            case 1009: // 174) ABPK Prepaid >> Prepaid Raw 1/m
                bankCode = 174; bankName = "ABPK"; stmntType = "Prepaid";
                strStatementType = "ABPK_Prepaid";
                stmntMail = "ABPK";
                break;

            case 70:   // 27) BCA BANCO COMERCIAL DO ATLANTICO >> Corporate >> Default Text 1/m
                bankCode = 27; bankName = "BCA"; stmntType = "Corporate";
                strStatementType = "BCA_Corporate";
                stmntMail = "BCA";
                //stmntDate = DateTime.Now;
                break;

            case 71:   // 40) ICB INTERCONTINENTAL BANK PLC >> Corporate Default Text 1/m
                bankCode = 40; bankName = "ICB"; stmntType = "Corporate";
                strStatementType = "ICB_Corporate";
                stmntMail = "ICB";
                //stmntDate = DateTime.Now;
                break;

            case 73:   // 30) BMG Banque Misr Gulf >> Text 1/m
                bankCode = 30; bankName = "BMG"; stmntType = "Credit";
                strStatementType = "BMG_Credit";
                stmntMail = "BMG";
                break;

            case 74:   // 15) SSB SG-SSB Ltd Debit >> PDF 1/m
                bankCode = 15; bankName = "SSB"; stmntType = "Debit";
                strStatementType = "SSB_Debit";
                stmntMail = "SSB";
                reportFleName += "Statement_Common_English_Debit.rpt";
                break;

            case 75:   // 51) TMB Trust Merchant Bank >> Raw VISA Text 20/m
                bankCode = 51; bankName = "TMB"; stmntType = "Credit";
                strStatementType = "TMB_VISA_Credit";
                stmntMail = "TMB";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 205:   // 6) AAIB Arab African International Bank >> Raw Text 1/m
                bankCode = 6; bankName = "AAIB"; stmntType = "Credit";
                strStatementType = "AAIB_Credit";
                stmntMail = "AAIB";
                break;

            case 431:   // 130) UNB Union National Bank >> Raw Text 1/m
                bankCode = 130; bankName = "UNB"; stmntType = "Credit";
                strStatementType = "UNB_Credit";
                stmntMail = "UNB"; statPeriod = "10";
                break;


            case 469:   // 153) BRKA AL Baraka Bank of Egypt >> Raw Text 30/m
                bankCode = 153; bankName = "BRKA"; stmntType = "Credit";
                strStatementType = "BRKA_Credit";
                stmntMail = "BRKA"; statPeriod = "10";
                break;

            case 437: //130) UNB Union National Bank >> Credit Text 10th/m
                bankCode = 130; bankName = "UNB"; stmntType = "Credit";
                strStatementType = "UNB_Credit_txt";
                statPeriod = "10";
                stmntMail = "UNB";
                break;


            case 1999: //130) UNB Union National Bank >> Credit Text 10th/m
                bankCode = 130; bankName = "UNB"; stmntType = "Credit";
                strStatementType = "ADBC_Corporate";
                stmntMail = "UNB";
                break;

            case 439: //94) FABG First Atlantic Bank Ghana >> Credit Text 1/m
                bankCode = 94; bankName = "FABG"; stmntType = "Credit";
                strStatementType = "FABG_Credit_txt";
                stmntMail = "FABG";
                break;

            case 467:   //94) FABG First Atlantic Bank Ghana >> Credit PDF 1/m
                bankCode = 94; bankName = "FABG"; stmntType = "Credit";
                strStatementType = "FABG_Credit_PDF";
                stmntMail = "FABG";
                reportFleName += "Statement_FABG_Credit.rpt";
                break;

            case 441: //154) UBP Unity Bank PlC >> Credit Text 15/m
                bankCode = 154; bankName = "UBP"; stmntType = "Credit";
                strStatementType = "UBP_Credit_txt";
                stmntMail = "UBP"; statPeriod = "15";
                break;

            case 459: //158) RBGH Republic Bank(Ghana) Limited >> Default Credit Text 23rd / m
                bankCode = 158; bankName = "RBGH"; stmntType = "Credit";
                strStatementType = "RBGH_Credit_txt";
                stmntMail = "RBGH";
                break;

            case 468:   //158) RBGH Republic Bank (Ghana) Limited >> Default Credit Emails 23rd/m
                bankCode = 158; bankName = "RBGH"; stmntType = "Credit_Emails";
                strStatementType = "RBGH_Credit_ClientsEmails";
                stmntMail = "RBGH";
                stmntClientEmail = "Credit_ClientsEmails";
                //stmntDate = datStmntData.Value; // DateTime.Now;
                strFileName = "_";
                reportFleName += "Statement_RBGH_Credit_Email.rpt";
                break;

            case 217:   // 6) AAIB Arab African International Bank >> Text Spool 1/m
                bankCode = 6; bankName = "AAIB"; stmntType = "Credit";
                strStatementType = "AAIB_Credit_txt";
                stmntMail = "AAIB";
                break;

            case 298:   // 122) ALXB ALEXBANK  >> Credit Raw 30/m
                bankCode = 122; bankName = "ALXB"; stmntType = "Credit";
                strStatementType = "ALXB_Credit";
                stmntMail = "ALXB";
                break;

            case 344:   // 122) ALXB ALEXBANK  >> Credit MF Raw 30/m
                bankCode = 122; bankName = "ALXB"; stmntType = "Credit";
                strStatementType = "ALXB_MF_Credit";
                stmntMail = "ALXB";
                break;

            case 299:   // 122) ALXB ALEXBANK  >> Credit Text 30/m
                bankCode = 122; bankName = "ALXB"; stmntType = "Credit";
                strStatementType = "ALXB_Credit_txt";
                stmntMail = "ALXB";
                break;

            case 343:   // 122) ALXB ALEXBANK  >> Credit MF Text 30/m
                bankCode = 122; bankName = "ALXB"; stmntType = "Credit";
                strStatementType = "ALXB_Credit_MF_txt";
                stmntMail = "ALXB";
                break;

            case 304:   // 122) ALXB ALEXBANK  >> Corporate Raw 30/m
                bankCode = 122; bankName = "ALXB"; stmntType = "Corporate";
                strStatementType = "ALXB_Corporate";
                stmntMail = "ALXB";
                break;

            case 305:   // 122) ALXB ALEXBANK  >> Corporate Text 30/m
                bankCode = 122; bankName = "ALXB"; stmntType = "Corporate";
                strStatementType = "ALXB_Corporate_txt";
                stmntMail = "ALXB";
                break;

            case 292:   // 127) AIBK Arab Investment Bank of Egypt  >> Credit PDF 30/m
                bankCode = 127; bankName = "AIBK"; stmntType = "Credit";
                strStatementType = "AIBK_Credit";
                stmntMail = "AIBK";
                reportFleName += "AIBStatement.rpt";
                break;

            case 412:   // 128) EGB   >> Credit Raw 30/m
                bankCode = 128; bankName = "EGB"; stmntType = "Credit_Raw";
                strStatementType = "EGB_Credit_Raw";
                stmntMail = "EGB";
                break;

            case 293:   // 127) AIBK Arab Investment Bank of Egypt  >> Credit Customer PDF 30/m
                bankCode = 127; bankName = "AIBK"; stmntType = "Credit";
                strStatementType = "AIBK_Credit_Customer";
                stmntMail = "AIBK";
                reportFleName += "AIBStatement.rpt";
                break;

            case 306:   // 127) AIBK Arab Investment Bank of Egypt  >> Credit Staff PDF 30/m
                bankCode = 127; bankName = "AIBK"; stmntType = "Credit";
                strStatementType = "AIBK_Credit_Staff";
                stmntMail = "AIBK";
                reportFleName += "AIBStatement.rpt";
                break;

            case 307:   // 127) AIBK Arab Investment Bank of Egypt  >> Installment PDF 30/m
                bankCode = 127; bankName = "AIBK"; stmntType = "Credit";
                strStatementType = "AIBK_Installment";
                stmntMail = "AIBK";
                reportFleName += "AIBInstallmentStatement.rpt";
                break;

            case 308:   // 127) AIBK Arab Investment Bank of Egypt  >> Installment Customer PDF 30/m
                bankCode = 127; bankName = "AIBK"; stmntType = "Credit";
                strStatementType = "AIBK_Installment_Customer";
                stmntMail = "AIBK";
                reportFleName += "AIBInstallmentStatement.rpt";
                break;

            case 309:   // 127) AIBK Arab Investment Bank of Egypt  >> Installment Staff PDF 30/m
                bankCode = 127; bankName = "AIBK"; stmntType = "Credit";
                strStatementType = "AIBK_Installment_Staff";
                stmntMail = "AIBK";
                reportFleName += "AIBInstallmentStatement.rpt";
                break;

            case 310:   // 127) AIBK Arab Investment Bank of Egypt  >> Prepaid Customer PDF 30/m
                bankCode = 127; bankName = "AIBK"; stmntType = "Prepaid";
                strStatementType = "AIBK_Prepaid_Customer";
                stmntMail = "AIBK";
                reportFleName += "AIBStatement.rpt";
                break;

            case 311:   // 127) AIBK Arab Investment Bank of Egypt  >> Prepaid Staff PDF 30/m
                bankCode = 127; bankName = "AIBK"; stmntType = "Prepaid";
                strStatementType = "AIBK_Prepaid_Staff";
                stmntMail = "AIBK";
                reportFleName += "AIBStatement.rpt";
                break;

            case 204:   // 51) TMB Trust Merchant Bank >> Raw MasterCard Text 20/m
                bankCode = 51; bankName = "TMB"; stmntType = "MasterCard_Credit";
                strStatementType = "TMB_MasterCard_Credit";
                stmntMail = "TMB";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 76:   // 24) BOAL Bank of Africa Mali >> Debit >> Default PDF 1/m
                bankCode = 24; bankName = "BOAL"; stmntType = "Debit";
                strStatementType = "BOAL_Debit";
                stmntMail = "BOAL";
                reportFleName += "Statement_Common_English_Debit_EMP.rpt";
                break;

            case 168:   // 24) BOAL Bank of Africa Mali >> Credit >> French PDF 1/m
                bankCode = 24; bankName = "BOAL"; stmntType = "Credit";
                strStatementType = "BOAL_Credit";
                stmntMail = "BOAL";
                reportFleName += "Statement_BOAL_French_Credit.rpt";
                break;

            case 198:   // 35) BOAL Bank of Africa Cote D'Ivoire >> Credit >> French PDF 1/m
                bankCode = 35; bankName = "BOAC"; stmntType = "Credit";
                strStatementType = "BOAC_Credit";
                stmntMail = "BOAC";
                reportFleName += "Statement_BOAC_French_Credit.rpt";
                break;

            case 199:   // 35) BOAS Bank of Africa Senegal >> Credit >> French PDF 1/m
                bankCode = 26; bankName = "BOAS"; stmntType = "Credit";
                strStatementType = "BOAS_Credit";
                stmntMail = "BOAS";
                reportFleName += "Statement_BOAS_French_Credit.rpt";
                break;

            case 226:   // 56) NCB National Commercial Bank >> Prepaid >> PDF 1/m
                bankCode = 56; bankName = "NCB"; stmntType = "Prepaid";
                strStatementType = "NCB_Prepaid";
                stmntMail = "NCB";
                reportFleName += "Statement_Debit_NCB.rpt";
                break;

            case 77:   // 56) NCB National Commercial Bank >> Credit >> PDF 1/m
                bankCode = 56; bankName = "NCB"; stmntType = "Credit";
                strStatementType = "NCB_Credit";
                stmntMail = "NCB";
                reportFleName += "Statement_NCB.rpt";
                break;

            case 78:   // clsStatementNSGBbusiness SME
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Corporate";
                strStatementType = "QNB_ALAHLI_Business_SME";
                //strFileName = "_Statement_File_";//_Business
                stmntMail = "QNB_ALAHLI";
                break;

            case 148: // NSGB Mastercard Business - SME 
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Corporate";
                strStatementType = "QNB_ALAHLI_MasterCard_Business_SME";
                //strFileName = "_Statement_File_";//_Business
                stmntMail = "QNB_ALAHLI";
                break;

            case 432: // NSGB Corporate Contract - FEDCOC 
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Corporate";
                strStatementType = "QNB_ALAHLI_Corporate_Contract_FEDCOC";
                stmntMail = "QNB_ALAHLI";
                break;

            case 433: // NSGB Cardholder Corporate - B2B 
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Corporate";
                strStatementType = "QNB_ALAHLI_Cardholder_Corporate_B2B";
                stmntMail = "QNB_ALAHLI";
                break;

            case 434: // NSGB Corporate Contract - B2B 
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Corporate";
                strStatementType = "QNB_ALAHLI_Corporate_Contract_B2B";
                stmntMail = "QNB_ALAHLI";
                break;

            case 92:   // 58) SBP SKYE BANK PLC >> Default Text 1/m
                bankCode = 58; bankName = "SBP"; stmntType = "Credit";
                strStatementType = "SBP_Credit_Gold_Platinum";
                stmntMail = "SBP";
                break;

            case 79:   // 58) SBP SKYE BANK PLC >> Default Text 16/m
                bankCode = 58; bankName = "SBP"; stmntType = "Credit";
                stmntMail = "SBP";
                strStatementType = "SBP_Credit_Classic_Premium";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 93:   //58] SBP SKYE BANK PLC >> Emails 1/m
                bankCode = 58; bankName = "SBP"; stmntType = "Credit_Gold_Platinum_Clients_Emails";
                strStatementType = "SBP_Credit_Gold_PlatinumClientsEmails";
                stmntClientEmail = "Credit_Gold_PlatinumClientsEmails";
                stmntMail = "SBP";
                strFileName = "_";
                break;

            case 80:   //58] SBP SKYE BANK PLC >> Emails 16/m
                bankCode = 58; bankName = "SBP"; stmntType = "Credit_Classic_Premium_Clients_Emails";
                strStatementType = "SBP_Credit_Classic_PremiumClientsEmails";
                stmntClientEmail = "Credit_Classic_PremiumClientsEmails";
                stmntMail = "SBP";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 253: //"[58] SBP SKYE BANK PLC >> MasterCard Platinum Credit Emails 16/m
                bankCode = 58; bankName = "SBP"; stmntType = "MasterCard_Platinum_Credit_Clients_Emails";
                strStatementType = "SBP_MasterCard_Platinum_CreditClientsEmails";
                stmntClientEmail = "MasterCard_Platinum_CreditClientsEmails";
                stmntMail = "SBP";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 126:   //58] SBP SKYE BANK PLC >> Corporate Emails 1/m
                bankCode = 58; bankName = "SBP"; stmntType = "Corporate_Emails";
                strStatementType = "SBP_CorporateEmails";
                stmntClientEmail = "CorporateEmails";
                stmntMail = "SBP";
                strFileName = "_";
                break;

            case 176:   //58] SBP SKYE BANK PLC >> Prepaid VISA Emails 1/m
                bankCode = 58; bankName = "SBP"; stmntType = "VISA_Prepaid_Emails";
                strStatementType = "SBP_VISA_PrepaidEmails";
                stmntClientEmail = "VISA_PrepaidEmails";
                stmntMail = "SBP";
                strFileName = "_";
                break;

            case 184:   //58] SBP SKYE BANK PLC >> Prepaid VISA-NTDC Emails 1/m
                bankCode = 58; bankName = "SBP"; stmntType = "VISA_Prepaid_NTDC_Emails";
                strStatementType = "SBP_VISA_Prepaid_NTDC_Emails";
                stmntClientEmail = "VISA_Prepaid_NTDC_Emails";
                stmntMail = "SBP";
                strFileName = "_";
                break;

            case 177:   //58] SBP SKYE BANK PLC >> Prepaid MasterCard Emails 1/m
                bankCode = 58; bankName = "SBP"; stmntType = "MasterCard_Prepaid_Emails";
                strStatementType = "SBP_MasterCard_PrepaidEmails";
                stmntClientEmail = "MasterCard_PrepaidEmails";
                stmntMail = "SBP";
                strFileName = "_";
                break;

            case 81:   // 16) ABP Access Bank Plc with Label Credit >> Reward Default Text 1/m
                bankCode = 16; bankName = "ABP"; stmntType = "Credit";//
                strStatementType = "ABP_Credit_Reward";
                stmntMail = "ABP";
                break;

            case 82:   // 53) RCB ROKEL COMMERCIAL BANK S/L LTD >> PDF 1/m
                bankCode = 53; bankName = "RCB"; stmntType = "Debit";
                strStatementType = "RCB_Debit";
                stmntMail = "RCB";
                reportFleName += "Statement_RCB_Debit.rpt";
                break;

            //case 85:   //7) BAI Prducts >> Email 1/m
            //    bankCode = 7; bankName = "BAI"; stmntType = "Credit_Emails";
            //    strStatementType = "BAI_Credit_ClientsEmails";
            //    stmntMail = "BAI";
            //    stmntClientEmail = "Credit_ClientsEmails";
            //    //stmntDate = datStmntData.Value; // DateTime.Now;
            //    strFileName = "_";
            //    break;

            case 460:   //7) BAI Visa Classic >> Email 1/m
                bankCode = 7; bankName = "BAI"; stmntType = "Classic_Emails";
                strStatementType = "BAI_Classic_ClientsEmails";
                stmntMail = "BAI";
                stmntClientEmail = "Classic_ClientsEmails";
                //stmntDate = datStmntData.Value; // DateTime.Now;
                strFileName = "_";
                reportFleName += "Statement_BAI_Portuguese_Email_Classic.rpt";
                break;

            case 461:   //7) BAI Visa Gold >> Email 1/m
                bankCode = 7; bankName = "BAI"; stmntType = "Gold_Emails";
                strStatementType = "BAI_Gold_ClientsEmails";
                stmntMail = "BAI";
                stmntClientEmail = "Gold_ClientsEmails";
                //stmntDate = datStmntData.Value; // DateTime.Now;
                strFileName = "_";
                reportFleName += "Statement_BAI_Portuguese_Email_Gold.rpt";
                break;

            case 462:   //7) BAI Visa Platinum >> Email 1/m
                bankCode = 7; bankName = "BAI"; stmntType = "Platinum_Emails";
                strStatementType = "BAI_Platinum_ClientsEmails";
                stmntMail = "BAI";
                stmntClientEmail = "Platinum_ClientsEmails";
                //stmntDate = datStmntData.Value; // DateTime.Now;
                strFileName = "_";
                reportFleName += "Statement_BAI_Portuguese_Email_Platinum.rpt";
                break;

            case 180:   // [87] WEMA BANK PLC NIGERIA >> Credit Emails 15/m
                bankCode = 87; bankName = "WEMA"; stmntType = "Credit_Emails";
                strStatementType = "WEMA_Credit_ClientsEmails";
                stmntMail = "WEMA";
                stmntClientEmail = "Credit_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                break;

            //case 223:   // [87] SBN Sterling BANK NIGERIA >> Credit Emails 1/m
            //    bankCode = 47; bankName = "SBN"; stmntType = "Credit_Emails";
            //    strStatementType = "SBN_Credit_ClientsEmails";
            //    stmntMail = "SBN";
            //    stmntClientEmail = "Credit_ClientsEmails";
            //    strFileName = "_";
            //    //stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
            //    break;

            //case 316:   // [87] SBN Sterling BANK NIGERIA >> Credit Emails 15/m
            //    bankCode = 47; bankName = "SBN"; stmntType = "Credit_SD15th_Emails";
            //    strStatementType = "SBN_Credit_SD15th_ClientsEmails";
            //    stmntMail = "SBN";
            //    stmntClientEmail = "Credit_SD15th_ClientsEmails";
            //    strFileName = "_";
            //    stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
            //    break;

            case 223:   // [87] SBN Sterling BANK NIGERIA >> Credit Emails 15/m
                bankCode = 47; bankName = "SBN"; stmntType = "Credit_Emails";
                strStatementType = "SBN_Credit_ClassicPlatinumGold_ClientsEmails";
                stmntMail = "SBN";
                stmntClientEmail = "Credit_ClassicPlatinumGold_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                break;

            case 316:   // [87] SBN Sterling BANK NIGERIA >> Credit Emails 15/m
                bankCode = 47; bankName = "SBN"; stmntType = "Credit_SD15th_Emails";
                strStatementType = "SBN_Credit_Infinite_ClientsEmails";
                stmntMail = "SBN";
                stmntClientEmail = "Credit_Infinite_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                break;
            case 401:   // [87] SBN Sterling BANK NIGERIA >> Credit Emails 15/m
                bankCode = 47; bankName = "SBN"; stmntType = "Credit_SD15th_Emails";
                strStatementType = "SBN_Credit_Signature_ClientsEmails";
                stmntMail = "SBN";
                stmntClientEmail = "Credit_Signature_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                break;

            case 224:   // 56) NCB National Commercial Bank >> Credit >> Email  1/m
                bankCode = 56; bankName = "NCB"; stmntType = "Credit_Emails";
                strStatementType = "NCB_Credit_ClientsEmails";
                stmntMail = "NCB";
                stmntClientEmail = "Credit_ClientsEmails";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                break;

            case 227:   // 56) NCB National Commercial Bank >> Prepaid >> Email  1/m
                bankCode = 56; bankName = "NCB"; stmntType = "Prepaid_Emails";
                strStatementType = "NCB_Prepaid_ClientsEmails";
                stmntMail = "NCB";
                stmntClientEmail = "Prepaid_ClientsEmails";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                break;

            case 183:   // [87] WEMA BANK PLC NIGERIA >> Prepaid Emails 15/m
                bankCode = 87; bankName = "WEMA"; stmntType = "Prepaid_Emails";
                strStatementType = "WEMA_Prepaid_ClientsEmails";
                stmntMail = "WEMA";
                stmntClientEmail = "Prepaid_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 128: //    [7] BAI Prepaid >> Email 1/m
                bankCode = 7; bankName = "BAI"; stmntType = "Prepaid_Emails";
                strStatementType = "BAI_Prepaid_ClientsEmails";
                stmntMail = "BAI";
                stmntClientEmail = "Prepaid_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_BAI_Portuguese_Email.rpt";
                break;

            //case 181: //    [7] BAI Corporate >> Email 1/m
            //    bankCode = 7; bankName = "BAI"; stmntType = "Corporate_Emails";
            //    strStatementType = "BAI_Corporate_ClientsEmails";
            //    stmntMail = "BAI";
            //    stmntClientEmail = "Corporate_ClientsEmails";
            //    strFileName = "_";
            //    break;

            case 463:   //7) BAI Visa Bussiness >> Email 1/m
                bankCode = 7; bankName = "BAI"; stmntType = "Bussiness_Emails";
                strStatementType = "BAI_Bussiness_ClientsEmails";
                stmntMail = "BAI";
                stmntClientEmail = "Bussiness_ClientsEmails";
                //stmntDate = datStmntData.Value; // DateTime.Now;
                strFileName = "_";
                reportFleName += "Statement_BAI_Portuguese_Email_Business.rpt";
                break;

            case 86:   // 10) BNP Credit Debit >> Text 1/m
                bankCode = 10; bankName = "BNP"; stmntType = "Debit";//R
                strStatementType = "BNP_Debit";
                stmntMail = "BNP";
                break;

            case 87:   // 50) ICBG Intercontinental Bank Ghana Limited Debit >> Text 1/m
                bankCode = 50; bankName = "ICBG"; stmntType = "Debit";
                strStatementType = "ICBG_Debit";
                stmntMail = "ICBG";
                break;

            case 113:   // 50) ICBG Intercontinental Bank Ghana Limited Prepaid >> PDF 1/m
                bankCode = 50; bankName = "ICBG"; stmntType = "Prepaid";
                strStatementType = "ICBG_Prepaid";
                stmntMail = "ICBG";
                reportFleName += "Statement_ICBG_Debit.rpt";
                break;

            case 130: // 67) ABPC Access Bank Plc Congo >> PDF 1/m
                bankCode = 67; bankName = "ABPC"; stmntType = "Debit";
                strStatementType = "ABPC_Debit";
                stmntMail = "ABPC";
                reportFleName += "Statement_ABPC_English_Debit.rpt";
                break;

            case 88:   // 4) BMSR Bank Misr Credit Gold >> Text 1/m
                bankCode = 4; bankName = "BMSR"; stmntType = "Credit";
                strStatementType = "BMSR_Credit_Gold";
                stmntMail = "BMSR_Credit";
                //strFileName = "_Statement_Classic_File_";
                whereCond = " contracttype in ('Visa Credit Gold - Secured','Visa Credit Gold - UnSecured')";
                break;

            case 89:   // 4) BMSR Bank Misr Credit Mobinet >> Text 1/m
                bankCode = 4; bankName = "BMSR"; stmntType = "Credit";
                strStatementType = "BMSR_Credit_Mobinet";
                stmntMail = "BMSR_Credit";
                //strFileName = "_Statement_Classic_File_";
                whereCond = " contracttype = 'Visa Classic Mobinet'";
                break;

            case 94:   // 58) SBP SKYE BANK PLC >> Default Debit Text 1/m
            case 90:   // 58) SBP SKYE BANK PLC >> Default Debit Text 16/m
                bankCode = 58; bankName = "SBP"; stmntType = "Debit";
                strStatementType = "SBP_Debit";
                stmntMail = "SBP";
                if (curIndx == 90)
                {
                    stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                }
                break;

            case 125: // 58) SBP SKYE BANK PLC >> Default Corporate Text 1/m
                bankCode = 58; bankName = "SBP"; stmntType = "Corporate";
                strStatementType = "SBP_Corporate";
                stmntMail = "SBP";
                break;

            case 297: // 5) AIB Arab International Bank >> Default Corporate Text 1/m
                bankCode = 5; bankName = "AIB"; stmntType = "Corporate";
                strStatementType = "AIB_Corporate-txt";
                stmntMail = "AIB";
                break;

            case 175: // 58) SBP SKYE BANK PLC >> Default Prepaid Text 1/m
                bankCode = 58; bankName = "SBP"; stmntType = "Prepaid";
                strStatementType = "SBP_Prepaid";
                stmntMail = "SBP";
                break;

            case 95:   // 55) FBP Fidelity Bank PLC >> Default Text 16/m
                bankCode = 55; bankName = "FBP"; stmntType = "Credit";
                strStatementType = "FBP_Credit";
                stmntMail = "FBP";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 218:   // 55) FBP Fidelity Bank PLC >> Default Text 1/m
                bankCode = 55; bankName = "FBP"; stmntType = "Credit_Payroll";
                strStatementType = "FBP_Credit_Payroll";
                stmntMail = "FBP";
                break;

            case 151:   // 55) FBP Fidelity Bank PLC >> Default Text Debit 16/m
                bankCode = 55; bankName = "FBP"; stmntType = "Debit";
                strStatementType = "FBP_Debit";
                stmntMail = "FBP";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 96:   //55] FBP Fidelity Bank PLC >> Emails 16/m
                bankCode = 55; bankName = "FBP"; stmntType = "Credit_Emails";
                strStatementType = "FBP_ClientsEmails";
                stmntClientEmail = "ClientsEmails";
                stmntMail = "FBP";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 219:   //55] FBP Fidelity Bank PLC >> Emails 1/m
                bankCode = 55; bankName = "FBP"; stmntType = "Payroll_Credit_Emails";
                strStatementType = "FBP_Payroll_ClientsEmails";
                stmntClientEmail = "Payroll_ClientsEmails";
                stmntMail = "FBP";
                strFileName = "_";
                break;

            //case 99:   // 50) ICBG Intercontinental Bank Ghana Limited Debit >> PDF 1/m
            //    bankCode = 50; bankName = "ICBG"; stmntType = "Credit";
            //    strStatementType = "ICBG_Credit";
            //    stmntMail = "ICBG";
            //    break;

            case 99:   // 50) ICBG Intercontinental Bank Ghana Limited Credit >> PDF 1/m
                bankCode = 50; bankName = "ICBG"; stmntType = "Credit";
                strStatementType = "ICBG_Credit";
                stmntMail = "ICBG";
                reportFleName += "Statement_ICBG_Credit.rpt";
                break;

            case 100:   // 13) BK Bank of Kigali >> Default Prepaid Text 1/m
                bankCode = 13; bankName = "BK"; stmntType = "Prepaid";
                strStatementType = "BK_Prepaid";
                stmntMail = "BK";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 101:   // 13] BK Bank of Kigali >> Emails 1/m
                bankCode = 13; bankName = "BK"; stmntType = "Prepaid_Emails";
                strStatementType = "BK_Prepaid_ClientsEmails";
                stmntClientEmail = "Prepaid_ClientsEmails";
                stmntMail = "BK";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 154:   // 13) BK Bank of Kigali >> Default Credit Text 1/m
                bankCode = 13; bankName = "BK"; stmntType = "Credit";
                strStatementType = "BK_Credit";
                stmntMail = "BK";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 328:   // 13) BK Bank of Kigali >> Default Credit VIP Text 5/m
                bankCode = 13; bankName = "BK"; stmntType = "Credit";
                strStatementType = "BK_Credit_VIP";
                stmntMail = "BK";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 329:   // 13) BK Bank of Kigali >> Default Corporate Text 1/m
                bankCode = 13; bankName = "BK"; stmntType = "Corporate";
                strStatementType = "BK_Corporate";
                stmntMail = "BK";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 330:   // 13) BK Bank of Kigali >> Default Corporate VIP Text 5/m
                bankCode = 13; bankName = "BK"; stmntType = "Corporate";
                strStatementType = "BK_Corporate_VIP";
                stmntMail = "BK";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 155:   // 13] BK Bank of Kigali >> VISA RWF Credit Emails 1/m
                bankCode = 13; bankName = "BK"; stmntType = "VISA_RWF_Credit_Emails";
                strStatementType = "BK_VISA_RWF_ClientsEmails";
                stmntClientEmail = "VISA_RWF_ClientsEmails";
                stmntMail = "BK";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 287:   // 13] BK Bank of Kigali >> VISA USD Credit Emails 1/m
                bankCode = 13; bankName = "BK"; stmntType = "VISA_USD_Credit_Emails";
                strStatementType = "BK_VISA_USD_ClientsEmails";
                stmntClientEmail = "VISA_USD_ClientsEmails";
                stmntMail = "BK";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 313:   // 13] BK Bank of Kigali >> MasterCard RWF Credit Emails 1/m
                bankCode = 13; bankName = "BK"; stmntType = "MasterCard_RWF_Credit_Emails";
                strStatementType = "BK_MasterCard_RWF_ClientsEmails";
                stmntClientEmail = "MasterCard_RWF_ClientsEmails";
                stmntMail = "BK";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 314:   // 13] BK Bank of Kigali >> MasterCard USD Credit Emails 1/m
                bankCode = 13; bankName = "BK"; stmntType = "MasterCard_USD_Credit_Emails";
                strStatementType = "BK_MasterCard_USD_ClientsEmails";
                stmntClientEmail = "MasterCard_USD_ClientsEmails";
                stmntMail = "BK";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 331:   // 13] BK Bank of Kigali >> VISA RWF Credit VIP Emails 5/m
                bankCode = 13; bankName = "BK"; stmntType = "VISA_VIP_RWF_Credit_Emails";
                strStatementType = "BK_VISA_VIP_RWF_ClientsEmails";
                stmntClientEmail = "VISA_VIP_RWF_ClientsEmails";
                stmntMail = "BK";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 332:   // 13] BK Bank of Kigali >> VISA USD Credit VIP Emails 5/m
                bankCode = 13; bankName = "BK"; stmntType = "VISA_VIP_USD_Credit_Emails";
                strStatementType = "BK_VISA_VIP_USD_ClientsEmails";
                stmntClientEmail = "VISA_VIP_USD_ClientsEmails";
                stmntMail = "BK";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 333:   // 13] BK Bank of Kigali >> MasterCard RWF Credit VIP Emails 5/m
                bankCode = 13; bankName = "BK"; stmntType = "MasterCard_VIP_RWF_Credit_Emails";
                strStatementType = "BK_MasterCard_VIP_RWF_ClientsEmails";
                stmntClientEmail = "MasterCard_VIP_RWF_ClientsEmails";
                stmntMail = "BK";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 334:   // 13] BK Bank of Kigali >> MasterCard USD Credit VIP Emails 5/m
                bankCode = 13; bankName = "BK"; stmntType = "MasterCard_VIP_USD_Credit_Emails";
                strStatementType = "BK_MasterCard_VIP_USD_ClientsEmails";
                stmntClientEmail = "MasterCard_VIP_USD_ClientsEmails";
                stmntMail = "BK";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 335:   // 13] BK Bank of Kigali >> MasterCard RWF Corporate Cardholder VIP Emails 5/m
                bankCode = 13; bankName = "BK"; stmntType = "MasterCard_Corporate Cardholder_VIP_RWF_Credit_Emails";
                strStatementType = "BK_MasterCard_Corporate Cardholder_VIP_RWF_ClientsEmails";
                stmntClientEmail = "MasterCard_Corporate Cardholder_VIP_RWF_ClientsEmails";
                stmntMail = "BK";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 336:   // 13] BK Bank of Kigali >> MasterCard USD Corporate Cardholder VIP Emails 5/m
                bankCode = 13; bankName = "BK"; stmntType = "MasterCard_Corporate Cardholder_VIP_USD_Credit_Emails";
                strStatementType = "BK_MasterCard_Corporate Cardholder_VIP_USD_ClientsEmails";
                stmntClientEmail = "MasterCard_Corporate Cardholder_VIP_USD_ClientsEmails";
                stmntMail = "BK";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "05"; // DateTime.Now;
                break;

            case 337:   // 13] BK Bank of Kigali >> MasterCard RWF Corporate Cardholder Emails 1/m
                bankCode = 13; bankName = "BK"; stmntType = "MasterCard_Corporate Cardholder_RWF_Credit_Emails";
                strStatementType = "BK_MasterCard_Corporate Cardholder_RWF_ClientsEmails";
                stmntClientEmail = "MasterCard_Corporate Cardholder_RWF_ClientsEmails";
                stmntMail = "BK";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 338:   // 13] BK Bank of Kigali >> MasterCard USD Corporate Cardholder Emails 1/m
                bankCode = 13; bankName = "BK"; stmntType = "MasterCard_Corporate Cardholder_USD_Credit_Emails";
                strStatementType = "BK_MasterCard_Corporate Cardholder_USD_ClientsEmails";
                stmntClientEmail = "MasterCard_Corporate Cardholder_USD_ClientsEmails";
                stmntMail = "BK";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 274:   // [50] ICBG Access Bank Ghana Limited >> Credit Emails 1/m
                bankCode = 50; bankName = "ICBG"; stmntType = "Credit_Emails";
                strStatementType = "ICBG_ClientsEmails";
                stmntClientEmail = "ClientsEmails";
                stmntMail = "ICBG";
                strFileName = "_";
                reportFleName += "Statement_ICBG_Credit.rpt";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 102:   // [36] DBN Diamond Bank Nigeria - VIP 1,2,5 Supplementary>> Email 1/m
                bankCode = 36; bankName = "DBN"; stmntType = "Credit_VIP1-2-5_Emails_sup"; //_VIP
                strStatementType = "DBN_VIP1-2-5_ClientsEmails_sup";//_Credit_VIP
                stmntClientEmail = "VIP1-2-5_ClientsEmails_sup";
                stmntMail = "DBN";//_Credit_VIP
                strFileName = "_";
                //reportFleName += "Statement_ABP_DBN_VIP1_2_5_Sup.rpt";
                break;

            case 103:   // [36] DBN Diamond Bank Nigeria - VIP Supplementary>> Email 16/m
                bankCode = 36; bankName = "DBN"; stmntType = "Credit_VIP_Emails_sup";
                strStatementType = "DBN_VIP_ClientsEmails_sup";//DBN_Credit_ClientsEmails
                stmntClientEmail = "VIP_ClientsEmails_sup";
                stmntMail = "DBN";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                strFileName = "_";
                break;

            case 104:   // [36] DBN Diamond Bank Nigeria Supplementary>> Email 16/m
                bankCode = 36; bankName = "DBN"; stmntType = "Credit_Emails_sup";
                strStatementType = "DBN_Credit_ClientsEmails_sup";//DBN_Credit_ClientsEmails
                stmntClientEmail = "Credit_ClientsEmails_sup";
                stmntMail = "DBN";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                strFileName = "_";
                break;

            case 105:   // 51] TMB Trust Merchant Bank >> Corporate 20/m Text
                bankCode = 51; bankName = "TMB"; stmntType = "Corporate";
                strStatementType = "TMB_Corporate";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                stmntMail = "TMB";
                //stmntDate = DateTime.Now;
                break;

            case 250: //112) GTBR Guaranty trust bank Rwanda >> Credit Text 1/m
                bankCode = 112; bankName = "GTBR"; stmntType = "Credit";
                strStatementType = "GTBR_Credit";
                stmntMail = "GTBR";
                break;

            case 258: //114) IDBE Industrial Development & Workers Bank of Egypt >> Credit Text 1/m
                bankCode = 114; bankName = "IDBE"; stmntType = "Credit";
                strStatementType = "IDBE_Credit";
                stmntMail = "IDBE";
                break;

            case 319: //136) BPG Bank Prestigo >> Credit Test 1/m
                bankCode = 136; bankName = "BPG"; stmntType = "Credit";
                strStatementType = "BPG_Credit";
                stmntMail = "BPG";
                break;

            case 400: //11) GTB Guaranty Trust Bank Ghana >> Credit Text 1/m
                bankCode = 11; bankName = "GTB"; stmntType = "Credit";
                strStatementType = "GTB_Credit";

                stmntMail = "GTB";
                break;


            case 280: //114) IDBE Industrial Development & Workers Bank of Egypt >> Credit XML 1/m
                bankCode = 114; bankName = "IDBE"; stmntType = "Credit_XML";
                strStatementType = "IDBE_Credit_XML";
                stmntMail = "IDBE";
                break;

            case 251: //112) GTBR Guaranty trust bank Rwanda >> Debit Text 1/m
                bankCode = 112; bankName = "GTBR"; stmntType = "Debit";
                strStatementType = "GTBR_Debit";
                stmntMail = "GTBR";
                break;

            case 339: //12) BSIC Alwaha Bank Libya Inc (Oasis Bank) >> Default Text 1/m
                bankCode = 12; bankName = "WAHA"; stmntType = "Debit";
                strStatementType = "WAHA_Debit";
                stmntMail = "BSIC";
                break;

            case 252: //112) GTBR Guaranty trust bank Rwanda >> Corporate Text 1/m
                bankCode = 112; bankName = "GTBR"; stmntType = "Corporate";
                strStatementType = "GTBR_Corporate";
                stmntMail = "GTBR";
                break;

            //case 106:   // 50] ICBG Intercontinental Bank Ghana Limited Corporate >> Text 1/m
            //    bankCode = 50; bankName = "ICBG"; stmntType = "Corporate";
            //    strStatementType = "ICBG_Corporate";
            //    stmntMail = "ICBG";
            //    //stmntDate = DateTime.Now;
            //    break;

            case 106:   // 50] ICBG Intercontinental Bank Ghana Limited Corporate >> PDF 1/m
                bankCode = 50; bankName = "ICBG"; stmntType = "Corporate";
                strStatementType = "ICBG_Corporate";
                stmntMail = "ICBG";
                //stmntDate = DateTime.Now;
                break;

            case 107:   // 51] TMB Trust Merchant Bank >> Corporate 20/m Text
                bankCode = 5; bankName = "AIB"; stmntType = "Credit";
                strStatementType = "AIB_PDF";
                stmntMail = "AIB";
                //reportFleName += "Statement_Common_English_Nice1.rpt";
                reportFleName += "Statement_Common_English_AIB.rpt";
                //stmntDate = DateTime.Now;
                break;

            case 108:   // 51] TMB Trust Merchant Bank >> Credit 20/m PDF
                bankCode = 51; bankName = "TMB"; stmntType = "Credit";
                strStatementType = "TMB_PDF";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                stmntMail = "TMB";
                //reportFleName += "Statement_Common_English_Nice1.rpt";
                reportFleName += "Statement_Credit_TMB.rpt";
                //stmntDate = DateTime.Now;
                break;

            case 109:   // 69) First Bank of Nigeria  >> Credit Default Text 1/m
            case 123:   // 69) First Bank of Nigeria  >> Credit Default Text 16/m
                //bankCode = 69; bankName = "FBN"; stmntType = "Credit_Gold";
                bankCode = 69; bankName = "FBN"; stmntType = "Credit";
                //strStatementType = "FBN_Credit_Gold";
                strStatementType = "FBN_Credit";
                stmntMail = "FBN";
                //if (curIndx == 123)
                //{
                //strStatementType = "FBN_Credit";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                //}
                break;

            case 142:  //69) FBN First Bank of Nigeria  >> MasterCard Credit Standard Default Text 1/m
                bankCode = 69; bankName = "FBN"; stmntType = "Credit_Standard";
                strStatementType = "FBN_MasterCard_Credit_Standard";
                stmntMail = "FBN";
                break;

            //case 115:   // 13) 69) First Bank of Nigeria  >> Prepaid Default Text 1/m
            case 115:   // 13) 69) First Bank of Nigeria  >> Prepaid Default Text 16/m
                bankCode = 69; bankName = "FBN"; stmntType = "Prepaid";
                strStatementType = "FBN_Prepaid";
                stmntMail = "FBN";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 116:   // 13) 69) First Bank of Nigeria  >> Corporate Default Text 1/m
                bankCode = 69; bankName = "FBN"; stmntType = "Corporate";
                strStatementType = "FBN_Corporate";
                stmntMail = "FBN";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 110:   // 69] First Bank of Nigeria  >> Credit Emails 1/m
            case 124:   //69] First Bank of Nigeria  >> Credit Emails 16/m
                //bankCode = 69; bankName = "FBN"; stmntType = "Credit_Gold_Emails";
                //strStatementType = "FBN_Credit_Gold_ClientsEmails";
                //stmntClientEmail = "CreditGoldClientsEmails";
                bankCode = 69; bankName = "FBN"; stmntType = "Credit_Emails";
                strStatementType = "FBN_CreditClientsEmails";
                stmntClientEmail = "CreditClientsEmails";
                stmntMail = "FBN";
                strFileName = "_";
                //if (curIndx == 124)
                //{
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                //}
                break;

            case 321:   //69] First Bank of Nigeria  >> Emails Classic 16/m
                bankCode = 69; bankName = "FBN"; stmntType = "Credit_Classic_Emails";
                strStatementType = "FBN_Credit_Classic_ClientsEmails";
                stmntClientEmail = "Credit_Classic_ClientsEmails";
                stmntMail = "FBN";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 322:   //69] First Bank of Nigeria  >> Emails Gold 16/m
                bankCode = 69; bankName = "FBN"; stmntType = "Credit_Gold_Emails";
                strStatementType = "FBN_Credit_Gold_ClientsEmails";
                stmntClientEmail = "Credit_Gold_ClientsEmails";
                stmntMail = "FBN";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 323:   //69] First Bank of Nigeria  >> Emails Infinite 16/m
                bankCode = 69; bankName = "FBN"; stmntType = "Credit_Infinite_Emails";
                strStatementType = "FBN_Credit_Infinite_ClientsEmails";
                stmntClientEmail = "Credit_Infinite_ClientsEmails";
                stmntMail = "FBN";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 324:   //69] First Bank of Nigeria  >> Emails Classic Platinum 16/m
                bankCode = 69; bankName = "FBN"; stmntType = "Credit_Classic_Platinum_Emails";
                strStatementType = "FBN_Credit_Classic_Platinum_ClientsEmails";
                stmntClientEmail = "Credit_Classic_Platinum_ClientsEmails";
                stmntMail = "FBN";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 152:   //[69] FBN First Bank of Nigeria  >> Supplementary Credit Emails 16/m
            case 153:   //[69] FBN First Bank of Nigeria  >> Supplementary Credit Emails 1/m
                //bankCode = 69; bankName = "FBN"; stmntType = "Credit_Gold_Emails_Sup";
                //strStatementType = "FBN_Credit_Gold_ClientsEmails_Sup";
                //stmntClientEmail = "CreditGoldClientsEmails_Sup";
                bankCode = 69; bankName = "FBN"; stmntType = "Credit_Emails_Sup";
                strStatementType = "FBN_CreditClientsEmails_Sup";
                stmntClientEmail = "CreditClientsEmails_Sup";
                stmntMail = "FBN";
                strFileName = "_";
                //if (curIndx == 152)
                //{
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                //}
                break;

            case 147:   //69) FBN First Bank of Nigeria  >> MasterCard Credit Standard Default Emails 1/m
                bankCode = 69; bankName = "FBN"; stmntType = "MasterCard_Credit_Emails";
                strStatementType = "FBN_MasterCard_CreditClientsEmails";
                stmntClientEmail = "MasterCard_CreditClientsEmails";
                stmntMail = "FBN";
                strFileName = "_";
                break;

            case 122:   // 69] First Bank of Nigeria  >> Corporate Cardholder Emails 1/m
                bankCode = 69; bankName = "FBN"; stmntType = "CorporateCardholder_Emails";
                strStatementType = "FBN_CorporateCardholderClientsEmails";
                stmntClientEmail = "CorporateCardholderClientsEmails";
                stmntMail = "FBN";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 117:   // 69] First Bank of Nigeria  >> Prepaid Emails 16/m
                bankCode = 69; bankName = "FBN"; stmntType = "Prepaid_Emails";
                strStatementType = "FBN_PrepaidClientsEmails";
                stmntClientEmail = "PrepaidClientsEmails";
                stmntMail = "FBN";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 203:   // 69] First Bank of Nigeria  >> MasterCard Prepaid Emails 16/m
                bankCode = 69; bankName = "FBN"; stmntType = "MC_Prepaid_Emails";
                strStatementType = "FBN_MC_PrepaidClientsEmails";
                stmntClientEmail = "MC_PrepaidClientsEmails";
                stmntMail = "FBN";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 190:   // [6] AAIB Arab African International Bank >> Prepaid Email 1/m
                bankCode = 6; bankName = "AAIB"; stmntType = "Prepaid_Emails";
                strStatementType = "AAIB_PrepaidClientsEmails";
                stmntClientEmail = "PrepaidClientsEmails";
                stmntMail = "AAIB";
                reportFleName += "Statement_AAIB_Prepaid.rpt";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 118:   // 69] First Bank of Nigeria  >> Corporate Company Emails 1/m
                bankCode = 69; bankName = "FBN"; stmntType = "CorporateCompany_Emails";
                strStatementType = "FBN_CorporateCompanyClientsEmails";
                stmntClientEmail = "CorporateCompanyClientsEmails";
                stmntMail = "FBN";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 114:   // [1] NSGB MasterCard Salary Prepaid >> PDF 1/m 
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Prepaid";
                strStatementType = "QNB_ALAHLI_Salary_Prepaid";
                stmntMail = "QNB_ALAHLI";
                reportFleName += "Statement_MasterCard_Prepaid_NSGB.rpt";
                break;

            case 295:   // 128) EGB Egyptian Gulf Bank of Egypt  >> Credit OMR PDF 30/m 
                bankCode = 128; bankName = "EGB"; stmntType = "Credit";
                strStatementType = "EGB_Credit";
                stmntMail = "EGB";
                reportFleName += "Statement_Credit_EGB.rpt";
                break;

            case 470:   // 153) BRKA AL Baraka Bank of Egypt >> Credit PDF 30/m
                bankCode = 153; bankName = "BRKA"; stmntType = "Credit";
                strStatementType = "BRKA_Credit_PDF";
                stmntMail = "BRKA";
                reportFleName += "Statement_Credit_BARKA.rpt";
                break;

            case 440: //23) Suez Canal Bank(SCB) >> Corporate Text 1/m
                bankCode = 23; bankName = "Suez"; stmntType = "Corporate";
                strStatementType = "Suez_Corporate_txt";
                stmntMail = "Suez";
                break;

            case 119: //42) OBI Oceanic Bank International >> Default Debit Text 1/m
                bankCode = 42; bankName = "OBI"; stmntType = "Debit";
                strStatementType = "OBI_Debit";
                stmntMail = "OBI";
                break;

            case 206: //42) ENG Echo Bank Nigeria >> Default Credit Text 20/m
                bankCode = 42; bankName = "ENG"; stmntType = "Credit_SD20th";
                strStatementType = "ENG_Credit_SD20th";
                stmntDate = datStmntData.Value; statPeriod = "20"; // DateTime.Now;
                stmntMail = "ENG";
                break;

            case 120: //47) SBN Sterling Bank Nigeria >> Default Debit Text 1/m
                bankCode = 47; bankName = "SBN"; stmntType = "Debit";
                strStatementType = "SBN_Debit";
                stmntMail = "SBN";
                break;

            case 248: //44) BOAK Bank of Africa Kenya >> Default Debit Text 1/m
                bankCode = 44; bankName = "BOAK"; stmntType = "Debit";
                strStatementType = "BOAK_Debit";
                stmntMail = "BOAK";
                break;

            case 254: //113) GTBU Guaranty trust bank Uganda  >> Debit Text 1/m
                bankCode = 113; bankName = "GTBU"; stmntType = "Debit";
                strStatementType = "GTBU_Debit";
                stmntMail = "GTBU";
                break;

            case 294: //129) SBL SAHARA BANK BNP PARIBAS  >> Prepaid Text 1/m
                bankCode = 129; bankName = "SBL"; stmntType = "Prepaid";
                strStatementType = "SBL_Prepaid";
                stmntMail = "SBL";
                break;

            case 222: //47) SBN Sterling Bank Nigeria >> Default Credit Text 1/m
                bankCode = 47; bankName = "SBN"; stmntType = "Credit";
                strStatementType = "SBN_Credit";
                stmntMail = "SBN";
                break;

            case 315: //47) SBN Sterling Bank Nigeria >> Default Credit Text 15/m
                bankCode = 47; bankName = "SBN"; stmntType = "Credit";
                strStatementType = "SBN_Credit_SD15th";
                stmntMail = "SBN";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 121: //65) SBPG SKYE BANK PLC Gambia >> Default Debit Text 1/m
                bankCode = 65; bankName = "SBPG"; stmntType = "Debit";
                strStatementType = "SBPG_Debit";
                stmntMail = "SBPG";
                break;

            case 129: // 66) ABPG Access Bank Plc Gambia >> Default Debit Text 1/m
                bankCode = 66; bankName = "ABPG"; stmntType = "Debit";
                strStatementType = "ABPG_Debit";
                stmntMail = "ABPG";
                break;

            case 131: //72) ABPR ACCESS BANK RWANDA SA >> Default Credit Text 1/m
                bankCode = 72; bankName = "ABPR"; stmntType = "Credit";
                strStatementType = "ABPR_Credit";
                stmntMail = "ABPR";
                break;

            case 212: //73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 7th/m
                bankCode = 73; bankName = "FCMB"; stmntType = "Visa_Credit_SD7th";
                strStatementType = "FCMB_Visa_Credit_SD7th";
                stmntDate = datStmntData.Value; statPeriod = "07"; // DateTime.Now;
                stmntMail = "FCMB";
                break;

            case 213: //73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 12th/m
                bankCode = 73; bankName = "FCMB"; stmntType = "Visa_Credit_SD12th";
                strStatementType = "FCMB_Visa_Credit_SD12th";
                stmntDate = datStmntData.Value; statPeriod = "12"; // DateTime.Now;
                stmntMail = "FCMB";
                break;

            case 214: //73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 17th/m
                bankCode = 73; bankName = "FCMB"; stmntType = "Visa_Credit_SD17th";
                strStatementType = "FCMB_Visa_Credit_SD17th";
                stmntDate = datStmntData.Value; statPeriod = "17"; // DateTime.Now;
                stmntMail = "FCMB";
                break;

            case 133: //73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 20th/m
                bankCode = 73; bankName = "FCMB"; stmntType = "Visa_Credit_SD20th";
                strStatementType = "FCMB_Visa_Credit_SD20th";
                stmntDate = datStmntData.Value; statPeriod = "20"; // DateTime.Now;
                stmntMail = "FCMB";
                break;

            case 215: //73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 23rd/m
                bankCode = 73; bankName = "FCMB"; stmntType = "Visa_Credit_SD23rd";
                strStatementType = "FCMB_Visa_Credit_SD23rd";
                stmntDate = datStmntData.Value; statPeriod = "23"; // DateTime.Now;
                stmntMail = "FCMB";
                break;

            case 216: //73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 27th/m
                bankCode = 73; bankName = "FCMB"; stmntType = "Visa_Credit_SD27th";
                strStatementType = "FCMB_Visa_Credit_SD27th";
                stmntDate = datStmntData.Value; statPeriod = "27"; // DateTime.Now;
                stmntMail = "FCMB";
                break;

            case 267: //73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 7th/m
                bankCode = 73; bankName = "FCMB"; stmntType = "MasterCard_Credit_SD7th";
                strStatementType = "FCMB_MasterCard_Credit_SD7th";
                stmntDate = datStmntData.Value; statPeriod = "07"; // DateTime.Now;
                stmntMail = "FCMB";
                break;

            case 268: //73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 12th/m
                bankCode = 73; bankName = "FCMB"; stmntType = "MasterCard_Credit_SD12th";
                strStatementType = "FCMB_MasterCard_Credit_SD12th";
                stmntDate = datStmntData.Value; statPeriod = "12"; // DateTime.Now;
                stmntMail = "FCMB";
                break;

            case 269: //73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 17th/m
                bankCode = 73; bankName = "FCMB"; stmntType = "MasterCard_Credit_SD17th";
                strStatementType = "FCMB_MasterCard_Credit_SD17th";
                stmntDate = datStmntData.Value; statPeriod = "17"; // DateTime.Now;
                stmntMail = "FCMB";
                break;

            case 270: //73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 20th/m
                bankCode = 73; bankName = "FCMB"; stmntType = "MasterCard_Credit_SD20th";
                strStatementType = "FCMB_MasterCard_Credit_SD20th";
                stmntDate = datStmntData.Value; statPeriod = "20"; // DateTime.Now;
                stmntMail = "FCMB";
                break;

            case 271: //73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 23rd/m
                bankCode = 73; bankName = "FCMB"; stmntType = "MasterCard_Credit_SD23rd";
                strStatementType = "FCMB_MasterCard_Credit_SD23rd";
                stmntDate = datStmntData.Value; statPeriod = "23"; // DateTime.Now;
                stmntMail = "FCMB";
                break;

            case 272: //73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 27th/m
                bankCode = 73; bankName = "FCMB"; stmntType = "MasterCard_Credit_SD27th";
                strStatementType = "FCMB_MasterCard_Credit_SD27th";
                stmntDate = datStmntData.Value; statPeriod = "27"; // DateTime.Now;
                stmntMail = "FCMB";
                break;

            case 447: //73) FCMB First City Monument Bank Nigeria >> Visa Credit Emails 7/m

                bankCode = 73; bankName = "FCMB"; stmntType = "Visa_Credit_SD7th";
                strStatementType = "FCMB_Visa_Credit_SD7th_ClientsEmails";
                stmntMail = "FCMB";
                stmntClientEmail = "Visa_Credit_SD7th_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "7";
                break;

            case 449: //73) FCMB First City Monument Bank Nigeria >> Visa Credit Emails 12/m

                bankCode = 73; bankName = "FCMB"; stmntType = "Visa_Credit_SD12th";
                strStatementType = "FCMB_Visa_Credit_SD12th_ClientsEmails";
                stmntMail = "FCMB";
                stmntClientEmail = "Visa_Credit_SD12th_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "12";// DateTime.Now;
                break;

            case 451: //73) FCMB First City Monument Bank Nigeria >> Visa Credit Emails 17/m

                bankCode = 73; bankName = "FCMB"; stmntType = "Visa_Credit_SD17th";
                strStatementType = "FCMB_Visa_Credit_SD17th_ClientsEmails";
                stmntMail = "FCMB";
                stmntClientEmail = "Visa_Credit_SD17th_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "17";// DateTime.Now;
                break;

            case 453: //73) FCMB First City Monument Bank Nigeria >> Visa Credit Emails 20/m

                bankCode = 73; bankName = "FCMB"; stmntType = "Visa_Credit_SD20th";
                strStatementType = "FCMB_Visa_Credit_SD20th_ClientsEmails";
                stmntMail = "FCMB";
                stmntClientEmail = "Visa_Credit_SD20th_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "20";// DateTime.Now;
                break;

            case 455: //73) FCMB First City Monument Bank Nigeria >> Visa Credit Emails 23/m

                bankCode = 73; bankName = "FCMB"; stmntType = "Visa_Credit_SD23rd";
                strStatementType = "FCMB_Visa_Credit_SD23rd_ClientsEmails";
                stmntMail = "FCMB";
                stmntClientEmail = "Visa_Credit_SD23rd_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "23";// DateTime.Now;
                break;

            case 457: //73) FCMB First City Monument Bank Nigeria >> Visa Credit Emails 27/m

                bankCode = 73; bankName = "FCMB"; stmntType = "Visa_Credit_SD27th";
                strStatementType = "FCMB_Visa_Credit_SD27th_ClientsEmails";
                stmntMail = "FCMB";
                stmntClientEmail = "Visa_Credit_SD27th_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "27";// DateTime.Now;
                break;

            case 448: //73) FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 7/m

                bankCode = 73; bankName = "FCMB"; stmntType = "MasterCard_Credit_SD7th";
                strStatementType = "FCMB_MasterCard_Credit_SD7th_ClientsEmails";
                stmntMail = "FCMB";
                stmntClientEmail = "MasterCard_Credit_SD7th_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "7";
                break;

            case 450: //73) FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 12/m

                bankCode = 73; bankName = "FCMB"; stmntType = "MasterCard_Credit_SD12th";
                strStatementType = "FCMB_MasterCard_Credit_SD12th_ClientsEmails";
                stmntMail = "FCMB";
                stmntClientEmail = "MasterCard_Credit_SD12th_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "12";// DateTime.Now;
                break;

            case 452: //73) FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 17/m

                bankCode = 73; bankName = "FCMB"; stmntType = "MasterCard_Credit_SD17th";
                strStatementType = "FCMB_MasterCard_Credit_SD17th_ClientsEmails";
                stmntMail = "FCMB";
                stmntClientEmail = "MasterCard_Credit_SD17th_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "17";// DateTime.Now;
                break;

            case 454: //73) FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 20/m

                bankCode = 73; bankName = "FCMB"; stmntType = "MasterCard_Credit_SD20th";
                strStatementType = "FCMB_MasterCard_Credit_SD20th_ClientsEmails";
                stmntMail = "FCMB";
                stmntClientEmail = "MasterCard_Credit_SD20th_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "20";// DateTime.Now;
                break;

            case 456: //73) FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 23/m

                bankCode = 73; bankName = "FCMB"; stmntType = "MasterCard_Credit_SD23rd";
                strStatementType = "FCMB_MasterCard_Credit_SD23rd_ClientsEmails";
                stmntMail = "FCMB";
                stmntClientEmail = "MasterCard_Credit_SD23rd_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "23";// DateTime.Now;
                break;

            case 458: //73) FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 27/m

                bankCode = 73; bankName = "FCMB"; stmntType = "MasterCard_Credit_SD27th";
                strStatementType = "FCMB_MasterCard_Credit_SD27th_ClientsEmails";
                stmntMail = "FCMB";
                stmntClientEmail = "MasterCard_Credit_SD27th_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "27";// DateTime.Now;
                break;

            case 276: //115) UNBN UNION BANK NIGERIA  >> Credit Text 15/m
                bankCode = 115; bankName = "UNBN"; stmntType = "Visa_Credit_SD15th";
                strStatementType = "UNBN_Visa_Credit_SD15th";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                stmntMail = "UNBN";
                break;

            case 277: //115) UNBN UNION BANK NIGERIA  >> Credit Text 20/m
                bankCode = 115; bankName = "UNBN"; stmntType = "Visa_Credit_SD20th";
                strStatementType = "UNBN_Visa_Credit_SD20th";
                stmntDate = datStmntData.Value; statPeriod = "20"; // DateTime.Now;
                stmntMail = "UNBN";
                break;

            case 278: //115) UNBN UNION BANK NIGERIA  >> Credit Text 27/m
                bankCode = 115; bankName = "UNBN"; stmntType = "Visa_Credit_SD27th";
                strStatementType = "UNBN_Visa_Credit_SD27th";
                stmntDate = datStmntData.Value; statPeriod = "27"; // DateTime.Now;
                stmntMail = "UNBN";
                break;

            case 279: //115) UNBN UNION BANK NIGERIA  >> Credit Text 30/m
                bankCode = 115; bankName = "UNBN"; stmntType = "Visa_Credit_SD30th";
                strStatementType = "UNBN_Visa_Credit_SD30th";
                //stmntDate = datStmntData.Value; statPeriod = "27"; // DateTime.Now;
                stmntMail = "UNBN";
                break;

            case 288: //[115] UNBN UNION BANK NIGERIA  >> Credit Emails 15/m
                bankCode = 115; bankName = "UNBN"; stmntType = "Visa_Credit_SD15th_Emails";
                strStatementType = "UNBN_Visa_Credit_SD15th_ClientsEmails";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                stmntMail = "UNBN";
                stmntClientEmail = "Visa_Credit_SD15th_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_UNBN_Credit.rpt";
                break;

            case 289: //[115] UNBN UNION BANK NIGERIA  >> Credit Emails 20/m
                bankCode = 115; bankName = "UNBN"; stmntType = "Visa_Credit_SD20th_Emails";
                strStatementType = "UNBN_Visa_Credit_SD20th_ClientsEmails";
                stmntDate = datStmntData.Value; statPeriod = "20"; // DateTime.Now;
                stmntMail = "UNBN";
                stmntClientEmail = "Visa_Credit_SD20th_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_UNBN_Credit.rpt";
                break;

            case 290: //[115] UNBN UNION BANK NIGERIA  >> Credit Emails 27/m
                bankCode = 115; bankName = "UNBN"; stmntType = "Visa_Credit_SD27th_Emails";
                strStatementType = "UNBN_Visa_Credit_SD27th_ClientsEmails";
                stmntDate = datStmntData.Value; statPeriod = "27"; // DateTime.Now;
                stmntMail = "UNBN";
                stmntClientEmail = "Visa_Credit_SD27th_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_UNBN_Credit.rpt";
                break;

            case 291: //[115] UNBN UNION BANK NIGERIA  >> Credit Emails 30/m
                bankCode = 115; bankName = "UNBN"; stmntType = "Visa_Credit_SD30th_Emails";
                strStatementType = "UNBN_Visa_Credit_SD30th_ClientsEmails";
                //stmntDate = datStmntData.Value; statPeriod = "27"; // DateTime.Now;
                stmntMail = "UNBN";
                stmntClientEmail = "Visa_Credit_SD30th_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_UNBN_Credit.rpt";
                break;

            case 135: //74) BLME Blom Bank Egypt VISA>> Credit PDF 1/m
                bankCode = 74; bankName = "BLME"; stmntType = "VISA_Credit";
                strStatementType = "ABC_Egypt_VISA_Credit_PDF";
                stmntMail = "BLME";
                //reportFleName += "Statement_BLME_VISA_Credit.rpt";
                reportFleName += "Statement_BLME_Credit.rpt";
                //reportFleName += "Statement_BLME_Credit - Copy.rpt";
                break;

            case 178: //74) BLME Blom Bank Egypt MasterCard>> Credit PDF 1/m
                bankCode = 74; bankName = "BLME"; stmntType = "MasterCard_CCredit";
                strStatementType = "ABC_Egypt_MasterCard_Credit_PDF";
                stmntMail = "BLME";
                reportFleName += "Statement_BLME_MasterCard_Credit.rpt";
                //reportFleName += "Statement_BLME_Credit.rpt";
                break;

            case 195:  //90) UBG UniBank Ghana Limited  >> Debit PDF 1/m
                bankCode = 90; bankName = "UBG"; stmntType = "Debit";
                strStatementType = "UBG_Debit";
                stmntMail = "UBG";
                reportFleName += "Statement_UBG_Debit.rpt";
                break;

            case 167:  //85) UTBG UT Bank Limited Ghana >> Default Credit PDF 1/m
                bankCode = 85; bankName = "UTBG"; stmntType = "Credit";
                strStatementType = "UTBG_Credit";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                stmntMail = "UTBG";
                reportFleName += "Statement_UTBG.rpt";
                break;

            case 143: //74) BLME Blom Bank Egypt VISA>> Credit Text 1/m
                bankCode = 74; bankName = "BLME"; stmntType = "VISA_Credit";
                strStatementType = "BLME_VISA_Credit_Text";
                stmntMail = "BLME";
                break;

            case 141: //74) BLME Blom Bank Egypt MasterCard >> Credit Text 1/m
                bankCode = 74; bankName = "BLME"; stmntType = "MasterCard_Credit";
                strStatementType = "BLME_MasterCard_Credit_Text";
                stmntMail = "BLME";
                break;

            case 132: // 72) ABPR ACCESS BANK RWANDA SA >> Default Debit Text 1/m
                bankCode = 72; bankName = "ABPR"; stmntType = "Debit";
                strStatementType = "ABPR_Debit";
                stmntMail = "ABPR";
                break;

            case 134: //73) FCMB First City Monument Bank Nigeria >> Default Debit Text 20/m
                bankCode = 73; bankName = "FCMB"; stmntType = "Debit";
                strStatementType = "FCMB_Debit";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                stmntMail = "FCMB";
                break;

            case 136: // 71) BOAD BANK OF AFRICA DEMOCRATIC REPUBLIC OF CONGO >> Default Debit Text 1/m
                bankCode = 71; bankName = "BOAD"; stmntType = "Debit";
                strStatementType = "BOAD_Debit";
                stmntMail = "BOAD";
                break;

            case 137: //75) UMB UNIVERSAL MERCHANT BANK GHANA >> Default Prepaid Text 1/m
                bankCode = 75; bankName = "UMB"; stmntType = "Prepaid";
                strStatementType = "UMB_Prepaid";
                stmntMail = "UMB";
                break;

            //case 140: // 7) BAI Corporate >> Text 1/m
            case 140: // 7) BAI Corporate >> PDF 1/m
                bankCode = 7; bankName = "BAI"; stmntType = "Corporate";
                strStatementType = "BAI_Corporate";
                stmntMail = "BAI";
                break;

            case 144:   // [22] BPC BANCO DE POUPANCA E CREDITO, SARL Credit >> Email 16/m
                bankCode = 22; bankName = "BPC"; stmntType = "Credit_Emails";
                strStatementType = "BPC_Credit_ClientsEmails";
                stmntMail = "BPC";
                stmntClientEmail = "Credit_ClientsEmails";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                strFileName = "_";
                break;

            case 192:   // 22) BPC BANCO DE POUPANCA E CREDITO, SARL Debit >> Email 1/m
                bankCode = 22; bankName = "BPC"; stmntType = "Debit_Emails";
                strStatementType = "BPC_Debit_ClientsEmails";
                stmntMail = "BPC";
                stmntClientEmail = "Debit_ClientsEmails";
                //stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                strFileName = "_";
                break;

            case 194:   // //32)UBAG UNITED BANK FOR AFRICA GHANA LIMITED >> Debit Email 16/m
                bankCode = 32; bankName = "UBAG"; stmntType = "Debit_Emails";
                strStatementType = "UBAG_Debit_ClientsEmails";
                stmntMail = "UBAG";
                stmntClientEmail = "Debit_ClientsEmails";
                strFileName = "_";
                break;

            case 186:   // [22] BPC BANCO DE POUPANCA E CREDITO, SARL Credit >> Cardholder Email 16/m
                bankCode = 22; bankName = "BPC"; stmntType = "Corporate_Cardholder_Emails";
                strStatementType = "BPC_Corporate_Cardholder_ClientsEmails";
                stmntMail = "BPC";
                stmntClientEmail = "Corporate_Cardholder_ClientsEmails";
                strFileName = "_";
                break;

            case 145:   // 31) NBS NBS Bank Limited >> Email 16/m
                bankCode = 31; bankName = "NBS"; stmntType = "Credit_Emails";
                strStatementType = "NBS_Credit_ClientsEmails";
                stmntMail = "NBS";
                stmntClientEmail = "Credit_ClientsEmails";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                strFileName = "_";
                reportFleName += "Statement_NBS_Credit.rpt";

                break;

            case 146: // 76) EDBE EXPORT DEVELOPMENT BANK OF EGYPT >> Default Credit Text 1/m
                bankCode = 76; bankName = "EDBE"; stmntType = "Credit";
                strStatementType = "EDBE_MasterCard_Credit";
                stmntMail = "EDBE";
                break;

            case 156: //81) SIBN STANBIC IBTC BANK NIGERIA  >> Credit Default Text 1/m
                bankCode = 81; bankName = "SIBN"; stmntType = "Credit";
                strStatementType = "SIBN_Credit";
                stmntMail = "SIBN";
                break;

            case 179: //87) WEMA BANK PLC NIGERIA >> Credit Default 15/m
                bankCode = 87; bankName = "WEMA"; stmntType = "Credit";
                strStatementType = "WEMA_Credit";
                stmntMail = "WEMA";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 182: //87) WEMA BANK PLC NIGERIA >> Prepaid Default 15/m
                bankCode = 87; bankName = "WEMA"; stmntType = "Prepaid";
                strStatementType = "WEMA_Prepaid";
                stmntMail = "WEMA";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                break;

            case 157: //[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Emails 1/m
                bankCode = 81; bankName = "SIBN"; stmntType = "Credit_Emails";
                strStatementType = "SIBN_Credit_ClientsEmails";
                stmntMail = "SIBN";
                stmntClientEmail = "Credit_ClientsEmails";
                reportFleName += "Statement_SIBN_Credit.rpt";
                strFileName = "_";
                break;

            case 208: //[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Platinum Emails 1/m
                bankCode = 81; bankName = "SIBN"; stmntType = "Credit_Platinum_Emails";
                strStatementType = "SIBN_Credit_Platinum_ClientsEmails";
                stmntMail = "SIBN";
                stmntClientEmail = "Credit_Platinum_ClientsEmails";
                reportFleName += "Statement_SIBN_Credit.rpt";
                strFileName = "_";
                break;

            case 209: //[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Gold Emails 1/m
                bankCode = 81; bankName = "SIBN"; stmntType = "Credit_Gold_Emails";
                strStatementType = "SIBN_Credit_Gold_ClientsEmails";
                stmntMail = "SIBN";
                stmntClientEmail = "Credit_Gold_ClientsEmails";
                reportFleName += "Statement_SIBN_Credit.rpt";
                strFileName = "_";
                break;

            case 210: //[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Infinite Emails 1/m
                bankCode = 81; bankName = "SIBN"; stmntType = "Credit_Infinite_Emails";
                strStatementType = "SIBN_Credit_Infinite_ClientsEmails";
                stmntMail = "SIBN";
                stmntClientEmail = "Credit_Infinite_ClientsEmails";
                reportFleName += "Statement_SIBN_Credit.rpt";
                strFileName = "_";
                break;

            case 259: //[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Silver Emails 1/m
                bankCode = 81; bankName = "SIBN"; stmntType = "Credit_Silver_Emails";
                strStatementType = "SIBN_Credit_Silver_ClientsEmails";
                stmntMail = "SIBN";
                stmntClientEmail = "Credit_Silver_ClientsEmails";
                reportFleName += "Statement_SIBN_Credit.rpt";
                strFileName = "_";
                break;

            case 158: //81) SIBN STANBIC IBTC BANK NIGERIA  >> Corporate Default Text 1/m
                bankCode = 81; bankName = "SIBN"; stmntType = "Corporate";
                strStatementType = "SIBN_Corporate";
                stmntMail = "SIBN";
                break;

            case 159: //[81] STANBIC IBTC BANK NIGERIA >> Corporate Cardholder Emails 1/m
                bankCode = 81; bankName = "SIBN"; stmntType = "CorporateCardholder_Emails";
                strStatementType = "SIBN_CorporateCardholder_ClientsEmails";
                stmntMail = "SIBN";
                stmntClientEmail = "CorporateCardholder_ClientsEmails";
                reportFleName += "Statement_SIBN_Credit.rpt";
                strFileName = "_";
                break;

            case 160:   // 16) ABP Access Bank Plc with Label Corporate >> PDF 1/m
                bankCode = 16; bankName = ""; stmntType = "Corporate";
                strStatementType = "Corporate";
                stmntMail = "";
                //stmntDate = DateTime.Now;
                break;

            //case 161:  //38) GTBN GUARANTY TRUST BANK PLC NIGERIA  >> Credit Default Text 1/m
            //    bankCode = 38; bankName = "GTBN"; stmntType = "Credit";
            //    strStatementType = "GTBN_Credit";
            //    stmntMail = "GTBN";
            //    stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
            //    break;

            case 161:  //38) GTBN GUARANTY TRUST BANK PLC NIGERIA  >> Credit Default PDF 16/m
                bankCode = 38; bankName = "GTBN"; stmntType = "Credit";
                strStatementType = "GTBN_Credit";
                stmntMail = "GTBN";
                reportFleName += "Statement_GTBN_Credit.rpt";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                break;

            case 326:  //38) GTBN GUARANTY TRUST BANK PLC NIGERIA  >> Corporate Default Text 16/m
                bankCode = 38; bankName = "GTBN"; stmntType = "Corporate";
                strStatementType = "GTBN_Corporate";
                stmntMail = "GTBN";
                //reportFleName += "Statement_GTBN_Credit.rpt";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                break;

            case 185: // 18) BOAB Bank of Africa Benin >> Default Credit Text 1/m
                bankCode = 18; bankName = "BOAB"; stmntType = "Credit";
                strStatementType = "BOAB_Credit";
                stmntMail = "BOAB";
                break;

            case 162:  //[38] GTBN GUARANTY TRUST BANK PLC NIGERIA  >> Credit Emails 16/m
                bankCode = 38; bankName = "GTBN"; stmntType = "Credit_Emails";
                strStatementType = "GTBN_Credit_ClientsEmails";
                stmntMail = "GTBN";
                stmntClientEmail = "Credit_ClientsEmails";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                reportFleName += "Statement_GTBN_Credit.rpt";
                strFileName = "_";
                break;

            case 327:  //[38] GTBN GUARANTY TRUST BANK PLC NIGERIA  >> Corporate Cardholder Emails 16/m
                bankCode = 38; bankName = "GTBN"; stmntType = "CorporateCardholder_Emails";
                strStatementType = "GTBN_CorporateCardholder_ClientsEmails";
                stmntMail = "GTBN";
                stmntClientEmail = "CorporateCardholder_ClientsEmails";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                reportFleName += "Statement_GTBN_Credit.rpt";
                strFileName = "_";
                break;

            case 169:   //1) NSGB Credit >> Classic Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit_Classic_Emails";
                strStatementType = "QNB_ALAHLI_Credit_Classic_ClientsEmails";
                stmntClientEmail = "Credit_Classic_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                strFileName = "_";
                break;

            case 170:   //1) NSGB Credit >> Gold Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit_Gold_Emails";
                strStatementType = "QNB_ALAHLI_Credit_Gold_ClientsEmails";
                stmntClientEmail = "Credit_Gold_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                strFileName = "_";
                break;

            case 171:   //1) NSGB Credit >> Platinum Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit_Platinum_Emails";
                strStatementType = "QNB_ALAHLI_Credit_Platinum_ClientsEmails";
                stmntClientEmail = "Credit_Platinum_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                strFileName = "_";
                break;

            case 172:   //1) NSGB Credit >> Infinite Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit_Infinite_Emails";
                strStatementType = "QNB_ALAHLI_Credit_Infinite_ClientsEmails";
                stmntClientEmail = "Credit_Infinite_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                strFileName = "_";
                break;

            case 174:   //1) NSGB Credit >> MasterCard Standard Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit_MasterCard_Standard_Emails";
                strStatementType = "QNB_ALAHLI_Credit_MasterCard_Standard_ClientsEmails";
                stmntClientEmail = "Credit_MasterCard_Standard_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                strFileName = "_";
                break;

            case 403:   //1) NSGB Credit >> MasterCard Titanium Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit_MasterCard_Titanium_Emails";
                strStatementType = "QNB_ALAHLI_Credit_MasterCard_Titanium_ClientsEmails";
                stmntClientEmail = "Credit_MasterCard_Titanium_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                strFileName = "_";
                break;

            case 207: //   60) GUM GUMHOURIA Bank  >> Prepaid PDF 1/m
                bankCode = 60; bankName = "GUM"; stmntType = "Prepaid";
                strStatementType = "GUM_Prepaid";
                stmntMail = "GUM";
                reportFleName += "Statement_Debit_GUM.rpt";
                //stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                break;

            //case 200: //  62) WHDA WHDA Bank  >> Prepaid PDF 1/m
            //    bankCode = 62; bankName = "WHDA"; stmntType = "Prepaid";
            //    strStatementType = "WHDA_Prepaid";
            //    stmntMail = "WHDA";
            //    reportFleName += "Statement_WHDA.rpt";
            //    //stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
            //    break;

            case 201: //  77) I&M Bank Rwanda Limited  >> Credit PDF 15/m
                bankCode = 77; bankName = "I&M"; stmntType = "Credit";
                strStatementType = "I&M_Credit";
                stmntMail = "I&M";
                reportFleName += "Statement_BCR.rpt";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                break;

            case 320: //  [77] I&M Bank Rwanda Limited  >> Credit Emails 15/m
                bankCode = 77; bankName = "I&M"; stmntType = "Credit_Emails";
                strStatementType = "I&M_Credit_ClientsEmails";
                stmntClientEmail = "Credit_ClientsEmails";
                stmntMail = "I&M";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                break;

            case 202: //  77) I&M Bank Rwanda Limited  >> Prepaid PDF 15/m
                bankCode = 77; bankName = "I&M"; stmntType = "Prepaid";
                strStatementType = "I&M_Prepaid";
                stmntMail = "I&M";
                reportFleName += "Statement_Debit_BCR.rpt";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                break;

            case 325: //  [77] I&M Bank Rwanda Limited  >> Prepaid Emails 15/m
                bankCode = 77; bankName = "I&M"; stmntType = "Prepaid_Emails";
                strStatementType = "I&M_Prepaid_ClientsEmails";
                stmntClientEmail = "Prepaid_ClientsEmails";
                stmntMail = "I&M";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                break;

            case 229: // 95) GTBL Guaranty trust bank Liberia  >> Debit Text 1/m
                bankCode = 95; bankName = "GTBL"; stmntType = "Debit";
                strStatementType = "GTBL_Debit";
                stmntMail = "GTBL";
                break;

            case 232: //104) DSBS Dahabshil Bank International Somalia  >> Debit Text 1/m
                bankCode = 104; bankName = "DSBS"; stmntType = "Debit";
                strStatementType = "DSBS_Debit";
                stmntMail = "DSBS";
                break;

            case 233: //106) DSBJ EAST AFRICA BANK Djibouti  >> Debit Text 1/m
                bankCode = 106; bankName = "DSBJ"; stmntType = "Debit";
                strStatementType = "DSBJ_Debit";
                stmntMail = "DSBJ";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                break;

            case 234: //106) DSBJ EAST AFRICA BANK Djibouti  >> Credit Islamic Text 1/m
                bankCode = 106; bankName = "DSBJ"; stmntType = "Credit_Islamic";
                strStatementType = "DSBJ_Credit_Islamic";
                stmntMail = "DSBJ";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                break;

            case 231: //102) VCBK Victoria Commercial Bank Ltd. Kenya  >> Credit Text 1/m
                bankCode = 102; bankName = "VCBK"; stmntType = "Credit";
                strStatementType = "VCBK_MasterCard_Credit-txt";
                stmntMail = "VCBK";
                break;

            case 235: //102) VCBK Victoria Commercial Bank Ltd. Kenya  >> Raw Text 1/m
                bankCode = 102; bankName = "VCBK"; stmntType = "Credit";
                strStatementType = "VCBK_MasterCard_Credit";
                stmntMail = "VCBK";
                break;

            case 228: //  98) GTBL Guaranty trust bank Kenya  >> Credit PDF 15/m
                bankCode = 98; bankName = "GTBK"; stmntType = "Credit";
                strStatementType = "GTBK_Credit";
                stmntMail = "GTBK";
                stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                reportFleName += "Statement_GTBK_Credit.rpt";
                break;

            case 256: //  98) GTBL Guaranty trust bank Kenya  >> Prepaid PDF 15/m
                bankCode = 98; bankName = "GTBK"; stmntType = "Prepaid";
                strStatementType = "GTBK_Prepaid";
                stmntMail = "GTBK";
                //stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                reportFleName += "Statement_GTBK_Prepaid.rpt";
                break;

            //iatta
            case 300: //27) BCA BANCO COMERCIAL DO ATLANTICO >> Credit >> Default Text 1/m
                bankCode = 27; bankName = "BCA"; stmntType = "Credit";
                strStatementType = "BCA_Credit_Text";
                stmntMail = "BCA";
                break;

            case 341: //27) BCA BANCO COMERCIAL DO ATLANTICO >> Corporate Text 30/m
                bankCode = 27; bankName = "BCA"; stmntType = "Corporate";
                strStatementType = "BCA_Corporate_Text";
                stmntMail = "BCA";
                break;

            case 405:   //[41] BDCA Banque Du Caire Classic  >> Credit Emails 5/m
                bankCode = 41; bankName = "BDCA"; stmntType = "Classic_Credit_Emails";
                strStatementType = "BDCA_Credit_Classic_ClientsEmails";
                stmntMail = "BDCA";
                stmntClientEmail = "Credit_Classic_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "5";// DateTime.Now;
                reportFleName += "BDCAStatementEmailingClassic.rpt";
                break;

            case 406:   //[41] BDCA Banque Du Caire Gold  >> Credit Emails 5/m
                bankCode = 41; bankName = "BDCA"; stmntType = "Gold_Credit_Emails";
                strStatementType = "BDCA_Credit_Gold_ClientsEmails";
                stmntMail = "BDCA";
                stmntClientEmail = "Credit_Gold_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "5";// DateTime.Now;
                reportFleName += "BDCAStatementEmailingGold.rpt";
                break;

            case 407:   //[41] BDCA Banque Du Caire MasterCard Standard  >> Credit Emails 5/m
                bankCode = 41; bankName = "BDCA"; stmntType = "MCStandard_Credit_Emails";
                strStatementType = "BDCA_Credit_MCStandard_ClientsEmails";
                stmntMail = "BDCA";
                stmntClientEmail = "Credit_MCStandard_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "5";// DateTime.Now;
                reportFleName += "BDCAStatementEmailingMCStandard.rpt";
                break;

            case 408:   //[41] BDCA Banque Du Caire MasterCard Gold  >> Credit Emails 5/m
                bankCode = 41; bankName = "BDCA"; stmntType = "MCGold_Credit_Emails";
                strStatementType = "BDCA_Credit_MCGold_ClientsEmails";
                stmntMail = "BDCA";
                stmntClientEmail = "Credit_MCGold_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "5";// DateTime.Now;
                reportFleName += "BDCAStatementEmailingMCGold.rpt";
                break;

            case 409:   //[41] BDCA Banque Du Caire MasterCard Installment  >> Credit Emails 5/m
                bankCode = 41; bankName = "BDCA"; stmntType = "MCInst_Credit_Emails";
                strStatementType = "BDCA_Credit_MCInst_ClientsEmails";
                stmntMail = "BDCA";
                stmntClientEmail = "Credit_MCInst_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "5";// DateTime.Now;
                reportFleName += "BDCAStatementEmailingInstallment.rpt";
                //reportFleName += "BDCAStatementEmailingInstallment_oldDesignJan21.rpt";
                break;

            case 410:   //[41] BDCA Banque Du Caire MasterCard Titanium  >> Credit Emails 5/m
                bankCode = 41; bankName = "BDCA"; stmntType = "MCTitanium_Credit_Emails";
                strStatementType = "BDCA_Credit_MCTitanium_ClientsEmails";
                stmntMail = "BDCA";
                stmntClientEmail = "Credit_MCTitanium_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "5";// DateTime.Now;
                reportFleName += "BDCAStatementEmailingTITANIUM.rpt";
                break;

            case 411:   //[41] BDCA Banque Du Caire MasterCard World Elite  >> Credit Emails 5/m
                bankCode = 41; bankName = "BDCA"; stmntType = "MCWrldElt_Credit_Emails";
                strStatementType = "BDCA_Credit_MCWrldElt_ClientsEmails";
                stmntMail = "BDCA";
                stmntClientEmail = "Credit_MCWrldElt_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "5";// DateTime.Now;
                reportFleName += "BDCAStatementEmailingWorldElite.rpt";
                break;

            case 511:   //[41] BDCA Banque Du Caire MasterCard Platinum  >> Credit Emails 5/m
                bankCode = 41; bankName = "BDCA"; stmntType = "MCplatinum_Credit_Emails";
                strStatementType = "BDCA_Credit_MCplatinum_ClientsEmails";
                stmntMail = "BDCA";
                stmntClientEmail = "Credit_MCplatinum_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "5";
                reportFleName += "BDCAStatementEmailingMCPlatinum.rpt";  //change this
                break;

            case 512:   //[41] BDCA Banque Du Caire MasterCard Corporate  >> Credit Emails 1/m
                bankCode = 41; bankName = "BDCA"; stmntType = "MCcorporate_Credit_Emails";
                strStatementType = "BDCA_Credit_MCcorporate_ClientsEmails";
                stmntMail = "BDCA";
                stmntClientEmail = "Credit_MCcorporate_ClientsEmails";
                strFileName = "_";
                stmntDate = datStmntData.Value; statPeriod = "5";
                reportFleName += "BDCAStatementEmailingMCCorporate.rpt";  //change this
                break;


            //ALXB iatta
            case 301:   //[122] ALXB ALEXBANK  >> Credit Emails 1/m
                bankCode = 122; bankName = "ALXB"; stmntType = "Credit_Emails";
                strStatementType = "ALXB_Credit_ClientsEmails";
                stmntMail = "ALXB";
                stmntClientEmail = "Credit_ClientsEmails";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                reportFleName += "ALEXBANKEStatement.rpt";
                break;

            case 342:   //[122] ALXB ALEXBANK  >> Credit MF Emails 1/m
                bankCode = 122; bankName = "ALXB"; stmntType = "Credit_Emails";
                strStatementType = "ALXB_Credit_MF_ClientsEmails";
                stmntMail = "ALXB";
                stmntClientEmail = "Credit_MF_ClientsEmails";
                strFileName = "_";
                reportFleName += "ALEXBANKEStatementMF.rpt";
                break;

            case 438:   //[122] ALXB ALEXBANK  >> Corporate Emails 1/m
                bankCode = 122; bankName = "ALXB"; stmntType = "Corporate_Emails";
                strStatementType = "ALXB_Corporate_ClientsEmails";
                stmntMail = "ALXB";
                stmntClientEmail = "Corporate_ClientsEmails";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                reportFleName += "ALEXBANKEStatement_Corp.rpt";
                break;

            //AIBK iatta
            case 302:   //[127] AIBK Arab Investment Bank of Egypt  >> Credit Emails 1/m
                bankCode = 127; bankName = "AIBK"; stmntType = "Credit_Emails";
                strStatementType = "AIBK_Credit_ClientsEmails";
                stmntMail = "AIBK";
                stmntClientEmail = "Credit_ClientsEmails";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                reportFleName += "AIBStatementEmailing.rpt";
                break;

            case 3097:   //[127] AIBK Arab Investment Bank of Egypt  >> Credit Emails 1/m
                bankCode = 127; bankName = "AIBK"; stmntType = "Valu_Credit_Emails";
                strStatementType = "AIBK_Credit_Valu_ClientsEmails";
                stmntMail = "AIBK";
                stmntClientEmail = "Credit_Valu_ClientsEmails";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                reportFleName += "AIBStatementEmailing_Valu.rpt";
                break;

            case 312:   //[127] AIBK Arab Investment Bank of Egypt  >> Installment Emails 1/m
                bankCode = 127; bankName = "AIBK"; stmntType = "Installment_Emails";
                strStatementType = "AIBK_Installment_ClientsEmails";
                stmntMail = "AIBK";
                stmntClientEmail = "Installment_ClientsEmails";
                strFileName = "_";
                //stmntDate = datStmntData.Value; statPeriod = "15";// DateTime.Now;
                reportFleName += "AIBInstallmentStatementEmailing.rpt";
                break;

            case 404: //1) NSGB Credit >> Platinum PDF Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit_Platinum_PDF_Emails";
                strStatementType = "QNB_ALAHLI_Credit_Platinum_PDF_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                stmntClientEmail = "Credit_Platinum_PDF_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_QNB_Credit_Platinum.rpt";
                break;

            case 420: //1) NSGB Credit >> Classic PDF Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit_Classic_PDF_Emails";
                strStatementType = "QNB_ALAHLI_Credit_Classic_PDF_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                stmntClientEmail = "Credit_Classic_PDF_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_QNB_Credit_Classic.rpt";
                break;

            case 421: //1) NSGB Credit >> Gold PDF Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit_Gold_PDF_Emails";
                strStatementType = "QNB_ALAHLI_Credit_Gold_PDF_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                stmntClientEmail = "Credit_Gold_PDF_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_QNB_Credit_Gold.rpt";
                break;

            case 422: //1) NSGB Credit >> Infinite PDF Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit_Infinite_PDF_Emails";
                strStatementType = "QNB_ALAHLI_Credit_Infinite_PDF_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                stmntClientEmail = "Credit_Infinite_PDF_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_QNB_Credit_Infinite.rpt";
                break;

            case 423: //1) NSGB Credit >> Individual PDF Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit_Individual_PDF_Emails";
                strStatementType = "QNB_ALAHLI_Credit_Individual_PDF_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                stmntClientEmail = "Credit_Individual_PDF_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_QNB_Credit_Individuals.rpt";
                break;

            case 424: //1) NSGB Credit >> MC_Standard PDF Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit_MC_Standard_PDF_Emails";
                strStatementType = "QNB_ALAHLI_Credit_MC_Standard_PDF_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                stmntClientEmail = "Credit_MC_Standard_PDF_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_QNB_Credit_MC_Standard.rpt";
                break;

            case 425: //1) NSGB Credit >> Titanium PDF Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit_Titanium_PDF_Emails";
                strStatementType = "QNB_ALAHLI_Credit_Titanium_PDF_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                stmntClientEmail = "Credit_Titanium_PDF_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_QNB_Credit_Titanium.rpt";
                break;

            case 446: //1) NSGB Credit >> MasterCard World Elite PDF Email 1/m
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit_MC_World_Elite_PDF_Emails";
                strStatementType = "QNB_ALAHLI_Credit_MC_World_Elite_PDF_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                stmntClientEmail = "Credit_MC_World_Elite_PDF_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_QNB_Credit_MC_World_Elite.rpt";
                break;

            case 472: //1) QNB ALAHLI Credit >> VISA Signature PDF Email 1/m", 472));//472
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "Credit_MC_VISA_Signature_PDF_Emails";
                strStatementType = "QNB_ALAHLI_Credit_MC_VISA_Signature_PDF_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                stmntClientEmail = "Credit_MC_VISA_Signature_PDF_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_QNB_Credit_MC_VISA_Signature.rpt";
                break;
            case 473: //1) QNB ALAHLI Credit >> bebasata Gold Credit Card PDF Email 1/m", 473));//473
                bankCode = 1; bankName = "QNB_ALAHLI"; stmntType = "bebasata_Gold_Credit_PDF_Emails";
                strStatementType = "QNB_ALAHLI_bebasata_Gold_Credit_PDF_ClientsEmails";
                stmntMail = "QNB_ALAHLI";
                stmntClientEmail = "bebasata_Gold_Credit_PDF_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_QNB_bebasata_Gold_Credit.rpt";
                break;

            case 402: //[139] Coronation Merchant Bank  >> Credit Emails EOM/15/m
                bankCode = 139; bankName = "CMB"; stmntType = "Credit_Emails";
                strStatementType = "CMB_Credit_ClientsEmails";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                stmntMail = "Coronation Merchant Bank – Credit Card Account Statement";
                stmntClientEmail = "Credit_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_CMB_Credit.rpt";
                break;

            case 500: //[146] Fidelity Bank Ghana Limited  >> PrePaid email 1/m
                bankCode = 146;
                bankName = "FBPG"; stmntType = "PrePaid_Emails";
                strStatementType = "FBPG_Prepaid_ClientsEmails";//FBPG
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                stmntMail = "FBPG";
                //stmntMail = "Fidelity Bank Ghana Limited – PrePaid Account Statement";
                stmntClientEmail = "PrePaid_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_FBPG_Prepaid_Email.rpt"; //change to PrePaid
                break;

            case 501: //[146] Fidelity Bank Ghana Limited  >> Credit email 1/m
                bankCode = 146;
                bankName = "FBPG"; stmntType = "Credit_Emails";
                strStatementType = "FBPG_Credit_ClientsEmails";
                stmntDate = datStmntData.Value; statPeriod = "15"; // DateTime.Now;
                stmntMail = "FBPG";
                stmntClientEmail = "Credit_ClientsEmails";
                strFileName = "_";
                reportFleName += "Statement_FBPG_Credit_Email.rpt"; // change to CreditCard
                break;

        } // end switch
          //BeginInvoke(statusDelegate, new object[] { bankCode, "<< End -- " + strStatementType });//strStatementType
          //txtRunResult.Text = "<< End -- " + strStatementType + "\r\n\r\n" + basText.Left(txtRunResult.Text, 2000);
          //if (chkBackupStatement.Checked == false)
          //  stmntType = "No";
    }
}
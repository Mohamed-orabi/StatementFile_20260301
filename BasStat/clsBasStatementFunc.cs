using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;

// Base Data for Statement
public class clsBasStatementFunc
{
    protected decimal calculateCrDb(decimal pBasVal, decimal pVal, string pCrDb)
    {
        if (pCrDb.ToUpper() == "CR")
            return pBasVal + pVal;
        else
            return pBasVal - pVal; ;
    }

    protected string CrDb(decimal pVal)
    {
        if (pVal > 0)
            return "CR";
        else
            return "  ";
    }

    protected string CrDbMinus(decimal pVal)
    {
        if (pVal < 0)
            return "CR";
        else
            return "  ";
    }

    protected string DbCr(decimal pVal)
    {
        if (pVal < 0)
            return "DB";
        else
            return "  ";
    }

    protected string CrDb(string pVal)
    {
        if (pVal == "CR")
            return "CR";
        else
            return "  ";
    }

    protected string printCrDb(string pVal)
        {
        if (pVal == "CR")
            return "CR";
        else
            return "DB";
        }

    protected string isSupplementCard(string pIsPrimary)
    {
        if (pIsPrimary.ToUpper() == "N")
            return "SUP";
        else
            return "   ";
    }

    protected string ValidateArbic(string pStrArbic)
    {
        string rtnStr = string.Empty, tmpStr = string.Empty;
        if (pStrArbic == null || pStrArbic.Trim() == "")
            rtnStr = "";
        else
        {
            rtnStr = pStrArbic;
            if (pStrArbic.Length > 3 && pStrArbic.Substring(0, 3) == "ÑÞã")
                rtnStr = pStrArbic.Substring(3);
            if (basText.isContainArbic(rtnStr))
                rtnStr = pStrArbic;
        }
        rtnStr = rtnStr.Trim();
        return rtnStr;
    }

    protected string ValidateArbicNum(string pStrArbic)
    {
        string rtnStr = string.Empty, tmpStr = string.Empty;
        if (string.IsNullOrEmpty(pStrArbic))
            rtnStr = "";
        else
        {
            rtnStr = pStrArbic = pStrArbic.Trim();
            if (basText.isContainArbic(rtnStr))
            {
                //rtnStr = pStrArbic;
                if (pStrArbic.Length > 3 && pStrArbic.Substring(0, 3) != "ÑÞã" && Char.IsNumber(pStrArbic, 0))
                {
                    rtnStr = "\u200F" + pStrArbic;
                }
            }
        }
        //rtnStr = rtnStr.Trim();
        return rtnStr;
    }

    protected bool isValidateCard(string pCardStat)
    {
        //if (pCardStat == "CLSB")
        //if (pCardStat == "Given" || pCardStat == "Embossed" || pCardStat == "New" || pCardStat == "New Pin Generated Only")// || pCardStat == "Pin Generated Only" || pCardStat == "Pin Generated"  Pin Genereted Only
    if (pCardStat == "Given" || pCardStat == "Embossed" || pCardStat == "New" || pCardStat == "New Pin Generation Only" || pCardStat == "Embossing" || pCardStat == "New Pin Generated Only" || pCardStat == "Pin generation" || pCardStat == "Pin generated" || pCardStat == "Entered")
        //EDT-918 => BDCA-1890 
    //if (pCardStat == "Given" || pCardStat == "Embossed" || pCardStat == "New" || pCardStat == "New Pin Generation Only" || pCardStat == "Embossing" || pCardStat == "New Pin Generated Only" || pCardStat == "Pin generation" || pCardStat == "Pin generated" || pCardStat == "Entered" || pCardStat == "Compromised")
            return true;
        else
            return false;
    }

    protected bool isValidateAccount(string pCardStat)
    {
        //if (pCardStat == "CLSB")
        if (pCardStat == "Open" || pCardStat == "Primary Open")
            return true;
        else
            return false;
    }

    protected decimal valDbCr(decimal pVal, string pDbCr)
    {
        if (pDbCr.ToUpper() == "DB")
            return pVal * -1;
        else
            return pVal;
    }

    protected bool CalcTransInterest(string pEntry)
    {
        bool interest = false;
        if (pEntry == "Charge interest for 0 operations group" ||
            pEntry == "Charge interest for 1" ||
            pEntry == "Charge interest for 2" ||
            pEntry == "Charge interest for 3" ||
            pEntry == "Charge interest for 4" ||
            pEntry == "Charge interest for 5" ||
            pEntry == "Charge interest for 6" ||
            pEntry == "Finanace Charges" ||
            pEntry == "Finance Charge" ||
            pEntry == "Finance Charges" ||
            pEntry == "Interest" ||
            pEntry == "INTEREST CHARGE" ||
            pEntry == "Juros" ||
            pEntry == "Charge interest for Installment" ||
            pEntry == "Installment Finance Charges" ||
            pEntry == "Interest Charges" ||
            pEntry == "Finance charges" ||
            pEntry == "Interest Charge" ||
            pEntry == "Acceleration Finance Charges" ||
            pEntry == "Charge interest for Acceleration")
            interest = true;
        else
            interest = false;
        return interest;
    }

    public clsBasStatementFunc()
    {
    }
}

using System;
//using System.Data;
//using System.IO;
//using System.Text;
//using System.Windows.Forms;
//using Oracle.DataAccess.Client;
//using System.Xml;

//
public class clsOMR
{
    private string OmrLine = "____";
    private string OmrBlankLine = "    ";

    public clsOMR()
    {
    }

    public string fixLine()
    {
        //Line 1 & 5
        return OmrLine;
    }

    public string Line2LastPage(int pCurPage, int pTotPage)
    {
        //Line 1 & 5
        if ((pCurPage % 4) == 0 || pCurPage == pTotPage)
            return OmrLine;
        else
            return OmrBlankLine;
    }

    public string LastPage(int pCurPage, int pTotPage)
    {
        //Line 1 & 5
        if (pCurPage == pTotPage)
            return OmrLine;
        else
            return OmrBlankLine;
    }

    public string Line3(int pCurPage)
    {
        //Line 3
        if (pCurPage == 3 || pCurPage == 4 || (pCurPage > 4 && ((pCurPage % 4) == 3 || (pCurPage % 4) == 0)))
            return OmrLine;
        else
            return OmrBlankLine;
    }

    public string Line4(int pCurPage)
    {
        //Line 4
        if (pCurPage == 2 || pCurPage == 4 || (pCurPage > 4 && ((pCurPage % 4) == 2 || (pCurPage % 4) == 0)))
            return OmrLine;
        else
            return OmrBlankLine;
    }

    public string Asterisk(int pCurPage)
    {
        string rtrnStr = string.Empty;
        //Line 4
        if (pCurPage == 1 || (pCurPage > 4 && pCurPage % 4 == 1)) //Page 1
            rtrnStr = "* ";
        else if (pCurPage == 2 || (pCurPage > 4 && pCurPage % 4 == 2)) //Page 2
            rtrnStr = " *";
        else if (pCurPage == 3 || (pCurPage > 4 && pCurPage % 4 == 3)) //Page 3
            rtrnStr = "**";
        else if (pCurPage == 4 || (pCurPage > 4 && pCurPage % 4 == 0)) //Page 4
            rtrnStr = "  ";

        return rtrnStr;
    }

    public string AsteriskPage4(int pCurPage)
    {
        string rtrnStr = string.Empty;
        //Line 4
        if (pCurPage == 4 || (pCurPage > 4 && pCurPage % 4 == 0)) //Page 4
            rtrnStr = "*";
        else
            rtrnStr = " ";

        return rtrnStr;
    }

    public string AsteriskPage4(int pCurPage, int pTotPages)
    {
        string rtrnStr = string.Empty;
        //Line 4
        if (pCurPage == pTotPages || pCurPage == 4 || (pCurPage > 4 && pCurPage % 4 == 0)) //Page 4
            rtrnStr = "*";
        else
            rtrnStr = " ";

        return rtrnStr;
    }

    public string AsteriskLastPage(int pCurPage, int pTotPages)
    {
        string rtrnStr = string.Empty;

        if (pCurPage == pTotPages) //Page 4
            rtrnStr = "*";
        else
            rtrnStr = " ";

        return rtrnStr;
    }

    public string AsteriskBMSR(int pCurPage)
    {
        string rtrnStr = string.Empty;
        //Line 7
        if (pCurPage == 0 || (pCurPage > 7 && pCurPage % 7 == 0)) //Page 1
            rtrnStr = "***";
        else if (pCurPage == 1 || (pCurPage > 7 && pCurPage % 7 == 1)) //Page 2
            rtrnStr = "*  ";
        else if (pCurPage == 2 || (pCurPage > 7 && pCurPage % 7 == 2)) //Page 3
            rtrnStr = " * ";
        else if (pCurPage == 3 || (pCurPage > 7 && pCurPage % 7 == 3)) //Page 4
            rtrnStr = "** ";
        else if (pCurPage == 4 || (pCurPage > 7 && pCurPage % 7 == 4)) //Page 5
            rtrnStr = "  *";
        else if (pCurPage == 5 || (pCurPage > 7 && pCurPage % 7 == 5)) //Page 6
            rtrnStr = "* *";
        else if (pCurPage == 6 || (pCurPage > 7 && pCurPage % 7 == 6)) //Page 7
            rtrnStr = " **";
        else if (pCurPage == 7 || (pCurPage > 7 && pCurPage % 7 == 7)) //Page 8
            rtrnStr = "***";

        return rtrnStr;
    }

    public string ParityCheck(int pCurPage, string omr)
    {
        string rtrnStr = string.Empty;
        int countomr = 0;
        int pageomr = 0;
        //if (pCurPage % 7 == 1 || pCurPage % 7 == 3 || pCurPage % 7 == 5 || pCurPage % 7 == 6)
        switch (pCurPage % 7)
        {
            case 0:
                pageomr = 3;
                break;
            case 1:
                pageomr = 1;
                break;
            case 2:
                pageomr = 1;
                break;
            case 3:
                pageomr = 2;
                break;
            case 4:
                pageomr = 1;
                break;
            case 5:
                pageomr = 2;
                break;
            case 6:
                pageomr = 2;
                break;
            case 7:
                pageomr = 3;
                break;
        }
        if (omr == "*-*")
            countomr = 3;
        else
            countomr = 0;
        if ((countomr + 5 + pageomr) % 2 == 0)
            rtrnStr = "   ";
        else
            rtrnStr = "***";

        return rtrnStr;
    }

    public string AsteriskLastPageBMSR(int pCurPage, int pTotPages)
    {
        string rtrnStr = string.Empty;

        if (pCurPage == pTotPages) //Page 7
            rtrnStr = "*";
        else
            rtrnStr = " ";

        return rtrnStr;
    }

    public string AsteriskLastPageBLME(int pCurPage, int pTotPages)
    {
        string rtrnStr = string.Empty;

        if (pCurPage == pTotPages) //Page 7
            rtrnStr = "*****";
        else
            rtrnStr = "     ";

        return rtrnStr;
    }

    public string dash1Line(int pCurPage, int pTotPage)
    {
        string rtrnStr = string.Empty;

        rtrnStr = OmrLine = "- ";  // fixed first line
        OmrBlankLine = "  ";
        rtrnStr = rtrnStr + LastPage(pCurPage, pTotPage);
        rtrnStr = rtrnStr + Line3(pCurPage);
        rtrnStr = rtrnStr + Line4(pCurPage);
        rtrnStr = rtrnStr + "-"; // fixed last line

        return rtrnStr;
    }

    public string setOmrLine { set { OmrLine = value; } }
    public string setOmrBlankLine { set { OmrBlankLine = value; } }

    ~clsOMR()
    {
    }
}

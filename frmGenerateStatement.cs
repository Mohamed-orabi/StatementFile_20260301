using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;


public partial class frmGenerateStatement : Form
{
//  private System.Timers.Timer aTimer;
  private int bankCode;
  private string strStatementType = string.Empty;
  //private delegate void StatusDelegate(int pBankCode, string pStat);
  //public delegate void SetProgressDelegate(int pValue);
  //public delegate void SetMinMaxProgressDelegate(int pValue);
  
  //private StatusDelegate statusDelegate;
  //public SetProgressDelegate setProgressDelegate;
  //public SetMinMaxProgressDelegate setMinMaxProgressDelegate;

  public frmGenerateStatement()
  {
    InitializeComponent();
  }

  private void frmGenerateSatatement_Load(object sender, EventArgs e)
  {
    //statusDelegate = new StatusDelegate(statusUpdate);
    //setProgressDelegate = new SetProgressDelegate(setProgress);
    //setMinMaxProgressDelegate = new SetMinMaxProgressDelegate(setMinMaxProgress);
    
    //BeginInvoke(statusDelegate, new object[] { bankCode, "Start Creat Statement for " + strStatementType });//strStatementType
    if (System.IO.File.Exists(Application.StartupPath + "\\frmBackground.jpg"))
    {
      this.BackgroundImage = Image.FromFile(Application.StartupPath + "\\frmBackground.jpg");
    }
    //this.Show();
  }


  private void statusUpdate(int pBankCode, string pStat)
  {
    lblStatus.Text = pStat + " \r\n";//"Statement File Creation Done for " + pStat
    txtRunResult.Text = lblStatus.Text + "\r\n\r\n" + basText.Left(txtRunResult.Text, 2000);
  }

  public void setProgress(int pValue)
  {
    progBarStatus.Value = pValue;
    txtCurProgress.Text = pValue.ToString();
    lblStatus.Text = Convert.ToString((int)(((float)pValue / (float)progBarStatus.Maximum) * 100.00)) + " %";
  }// setProgress

  public void setMinMaxProgress(int pValue)
  {
    progBarStatus.Minimum = 0;
    progBarStatus.Maximum = pValue;
    txtTotProgress.Text = pValue.ToString();
  }

  private void btnExit_Click(object sender, EventArgs e)
  {
    this.Close();

  }


}

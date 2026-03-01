using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Globalization;
using System.Reflection;
using POPMailMessage;
using POPMail;
using System.Data;
using Oracle.DataAccess.Client;
using StatementFile.StatementFile;
using System.Diagnostics;
using StatementFile;
using StatementFile.StatementFile.mailConfiguration;

/// Summary description for frmStatementFile.
public partial class frmStatementFile : System.Windows.Forms.Form
{
    private System.Windows.Forms.Label lblStatementPath;
    private System.Windows.Forms.Button btnExit;
    private System.Windows.Forms.TextBox txtFileName;
    private System.Windows.Forms.Button btnSaveFileName;
    private System.Windows.Forms.Button btnSaveStatement;
    private System.Windows.Forms.Label lblVer;
    private System.Windows.Forms.Label lblStatus;
    private System.Windows.Forms.ProgressBar progBarStatus;
    private IContainer components;
    private string strServer = string.Empty, strUserName = string.Empty;
    public static bool Internal = false;
    //Islam Atta    
    public static DateTime StDate = DateTime.Now;
    public static ArrayList InternalMailTo = new ArrayList();
    public static ArrayList InternalMailCC = new ArrayList();
    public static ArrayList InternalMailBCC = new ArrayList();
    public static string InternalMailFrom = string.Empty;
    public static string InternalMailFromName = string.Empty;

    private bool cancellationPending = false;
    private bool isBusy = false;
    private System.ComponentModel.BackgroundWorker backgroundWorker;
    //private System.ComponentModel.BackgroundWorker generateAllSelected;
    //private int curCmbProdInx;
    private TextBox txtFileMD5;
    private Button btnGetFileMD5;
    private TextBox txtFileName2;
    private Button btnSendOprMail;
    private string strStatementType = string.Empty, stmntType = string.Empty, stmntMail = string.Empty;// N Normal statement C Corperate Statement 
    private int bankCode;
    private string bankName, strFileName, reportFleName;
    private Button btnSendBankMail;
    private CheckBox chkBackupStatement;
    private CheckBox chkAppendData;
    private DateTime stmntDate;
    private DateTimePicker datStmntData;
    private bool appendData;
    private TextBox txtRunResult;
    private Button btnSelectFile;
    private Button btnUpdateData;
    private Button btnMaintainBank;
    private CheckBox chkSendOprMail;
    private CheckBox chkSendBankMail;
    private CheckedListBox chkLstProducts;
    private Button btnAll;
    private CheckBox chkSaveStatement;
    private string checkErrRslt = string.Empty;
    private int curIndx, curIndxVal;
    private System.Timers.Timer aTimer;
    private CheckBox chkDontPrompt;
    private CheckBox chkCheckEmail;
    private TextBox txtPassword;
    private TextBox txtUserName;
    private Label lblPassword;
    private Label lblUserName;
    private string sStr;
    private CheckBox chkMaintainData;
    private CheckBox chkExitAfterComplete;
    private CheckBox chkDontFixBranchData;
    private ArrayList emailArray = new ArrayList();
    public delegate void StatusDelegate(int pBankCode, string pStat);
    public delegate void SetProgressDelegate(int pValue);
    public delegate void SetMinMaxProgressDelegate(int pValue);
    private string stmntClientEmail;

    public StatusDelegate setStatusDelegate;
    public SetProgressDelegate setProgressDelegate;
    private TextBox txtTotProgress;
    private TextBox txtCurProgress;
    private Button btnUnSelectAll;
    private Label lblBank;
    private TextBox txtBank;
    private Label lblBankCode;
    private TextBox txtBankCode;
    private CheckBox chkStatGenAfterMonth;
    private CheckBox chkDevelopers;
    private CheckBox chkValidateBankEmail;
    private Button btnLoad;
    private CheckBox chkMonitorFtp;
    private CheckBox chkUpdateSummary;
    private CheckBox chkGenerateAtTime;
    private DateTimePicker datGenerateAtTime;
    public SetMinMaxProgressDelegate setMinMaxProgressDelegate;
    //private EventHandler onCreateComplete;
    private System.Windows.Forms.NotifyIcon notifyIcon1;
    private System.Windows.Forms.ContextMenu contextMenu1;
    private System.Windows.Forms.MenuItem menuItem1;
    private TextBox txtDbSchema;
    private TextBox txtTblMaster;
    private TextBox txtTblDetail;
    private System.Windows.Forms.MenuItem menuItem2;
    private string whereCond = string.Empty;
    private CheckBox chkStartCycle;
    private CheckBox chkEndCycle;
    private Panel panel1;
    private CheckBox chkTclientRestor;
    private CheckBox chkMaintain_Data;
    private CheckBox chkStatPlanExec;
    private CheckBox chkExportData;
    private TextBox txtTablePrefix;
    private Button btnDumpName;
    private CheckBox chkImportData;
    private CheckBox chkRenameTables;
    private CheckBox chkUpdateSummaryPart;
    private string statPeriod = string.Empty;
    private const int CP_NOCLOSE_BUTTON = 0x200;
    private string MainSchema = clsCnfg.readSetting("MainSchema");
    string mstrCR = "TSTATEMENTMASTERCR", dtlCR = "TSTATEMENTDETAILCR";
    string mstrDB = "TSTATEMENTMASTERDB", dtlDB = "TSTATEMENTDETAILDB";
    string mstrCP = "TSTATEMENTMASTERCP", dtlCP = "TSTATEMENTDETAILCP";
    private Label label1;
    private CheckBox chkTest;
    private ComboBox comboBox1;
    private Label label3;
    private ComboBox comboBox2;
    private Label label4;
    private TextBox txtPrefixRun;
    private Label label5;
    private Label label2;
    string[] CyclesCR = { "5th", "7th", "10th", "12th", "15th", "17th", "20th", "23rd", "27th", "EOM" };
    string[] EStmtCR = { "5th", "7th", "12th", "15th", "17th", "20th", "23rd", "27th", "EOM" };
    string[] CyclesDB = { "15th", "20th", "EOM" };
    string[] EStmtDB = { "15th", "EOM" };
    string[] CyclesCP = { "5th", "15th", "20th", "EOM" };
    string[] EStmtCP = { "5th", "15th", "EOM" };
    private MenuStrip menuStrip1;
    private ToolStripMenuItem utilitiesToolStripMenuItem;
    private ToolStripMenuItem changeTypeToolStripMenuItem;
    private ToolStripMenuItem cReditToolStripMenuItem;
    private ToolStripMenuItem corporateToolStripMenuItem;
    private ToolStripMenuItem debitToolStripMenuItem;
    public Button button1;

    public static int MailCount = 0;
    AssemblyInfo inf;
    public static string internalAccNo;
    private Button btnBanksMails;
    private Button CombineFiles;
    private bool dontSend = false;
    private string biliingFileFinalName = "StatementFinalBilling_20";

    public static PathConfiguration pathConfig = ConfigurationReader.LoadPathConfiguration();
    public string stmtPath = pathConfig.stmtPath;

    public frmStatementFile()
    {
        //
        // Required for Windows Form Designer support
        //

        InitializeComponent();

        //
        // TODO: Add any constructor code after InitializeComponent call
        //
    }

    //
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            if (components != null)
            {
                components.Dispose();
            }
        }
        base.Dispose(disposing);
        Application.Exit();
    }

    protected override CreateParams CreateParams
    {
        get
        {
            CreateParams myCp = base.CreateParams;
            myCp.ClassStyle = myCp.ClassStyle | CP_NOCLOSE_BUTTON;
            return myCp;
        }
    }
    #region Windows Form Designer generated code
    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmStatementFile));
            this.lblStatementPath = new System.Windows.Forms.Label();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.btnSaveFileName = new System.Windows.Forms.Button();
            this.btnSaveStatement = new System.Windows.Forms.Button();
            this.lblVer = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            this.progBarStatus = new System.Windows.Forms.ProgressBar();
            this.txtFileMD5 = new System.Windows.Forms.TextBox();
            this.btnGetFileMD5 = new System.Windows.Forms.Button();
            this.txtFileName2 = new System.Windows.Forms.TextBox();
            this.btnSendOprMail = new System.Windows.Forms.Button();
            this.btnSendBankMail = new System.Windows.Forms.Button();
            this.chkBackupStatement = new System.Windows.Forms.CheckBox();
            this.chkAppendData = new System.Windows.Forms.CheckBox();
            this.datStmntData = new System.Windows.Forms.DateTimePicker();
            this.txtRunResult = new System.Windows.Forms.TextBox();
            this.btnUpdateData = new System.Windows.Forms.Button();
            this.btnMaintainBank = new System.Windows.Forms.Button();
            this.chkSendOprMail = new System.Windows.Forms.CheckBox();
            this.chkSendBankMail = new System.Windows.Forms.CheckBox();
            this.chkLstProducts = new System.Windows.Forms.CheckedListBox();
            this.btnAll = new System.Windows.Forms.Button();
            this.chkSaveStatement = new System.Windows.Forms.CheckBox();
            this.chkDontPrompt = new System.Windows.Forms.CheckBox();
            this.chkCheckEmail = new System.Windows.Forms.CheckBox();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtUserName = new System.Windows.Forms.TextBox();
            this.lblPassword = new System.Windows.Forms.Label();
            this.lblUserName = new System.Windows.Forms.Label();
            this.chkMaintainData = new System.Windows.Forms.CheckBox();
            this.chkExitAfterComplete = new System.Windows.Forms.CheckBox();
            this.chkDontFixBranchData = new System.Windows.Forms.CheckBox();
            this.txtTotProgress = new System.Windows.Forms.TextBox();
            this.txtCurProgress = new System.Windows.Forms.TextBox();
            this.btnUnSelectAll = new System.Windows.Forms.Button();
            this.lblBank = new System.Windows.Forms.Label();
            this.txtBank = new System.Windows.Forms.TextBox();
            this.lblBankCode = new System.Windows.Forms.Label();
            this.txtBankCode = new System.Windows.Forms.TextBox();
            this.chkStatGenAfterMonth = new System.Windows.Forms.CheckBox();
            this.chkDevelopers = new System.Windows.Forms.CheckBox();
            this.chkValidateBankEmail = new System.Windows.Forms.CheckBox();
            this.btnLoad = new System.Windows.Forms.Button();
            this.chkMonitorFtp = new System.Windows.Forms.CheckBox();
            this.chkUpdateSummary = new System.Windows.Forms.CheckBox();
            this.chkGenerateAtTime = new System.Windows.Forms.CheckBox();
            this.datGenerateAtTime = new System.Windows.Forms.DateTimePicker();
            this.txtDbSchema = new System.Windows.Forms.TextBox();
            this.txtTblMaster = new System.Windows.Forms.TextBox();
            this.txtTblDetail = new System.Windows.Forms.TextBox();
            this.chkStartCycle = new System.Windows.Forms.CheckBox();
            this.chkEndCycle = new System.Windows.Forms.CheckBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.chkTclientRestor = new System.Windows.Forms.CheckBox();
            this.chkMaintain_Data = new System.Windows.Forms.CheckBox();
            this.chkStatPlanExec = new System.Windows.Forms.CheckBox();
            this.chkExportData = new System.Windows.Forms.CheckBox();
            this.txtTablePrefix = new System.Windows.Forms.TextBox();
            this.btnDumpName = new System.Windows.Forms.Button();
            this.chkImportData = new System.Windows.Forms.CheckBox();
            this.chkRenameTables = new System.Windows.Forms.CheckBox();
            this.chkUpdateSummaryPart = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.chkTest = new System.Windows.Forms.CheckBox();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtPrefixRun = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.utilitiesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.changeTypeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cReditToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.corporateToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.debitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.button1 = new System.Windows.Forms.Button();
            this.btnBanksMails = new System.Windows.Forms.Button();
            this.CombineFiles = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblStatementPath
            // 
            this.lblStatementPath.BackColor = System.Drawing.Color.Transparent;
            this.lblStatementPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            this.lblStatementPath.Location = new System.Drawing.Point(708, 326);
            this.lblStatementPath.Name = "lblStatementPath";
            this.lblStatementPath.Size = new System.Drawing.Size(119, 20);
            this.lblStatementPath.TabIndex = 29;
            this.lblStatementPath.Text = "Statement Path";
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(708, 343);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(192, 20);
            this.txtFileName.TabIndex = 30;
            this.txtFileName.Text = stmtPath;
            // 
            // btnSaveFileName
            // 
            this.btnSaveFileName.Image = global::StatementFile.Properties.Resources._264;
            this.btnSaveFileName.Location = new System.Drawing.Point(903, 346);
            this.btnSaveFileName.Name = "btnSaveFileName";
            this.btnSaveFileName.Size = new System.Drawing.Size(22, 18);
            this.btnSaveFileName.TabIndex = 31;
            this.btnSaveFileName.Click += new System.EventHandler(this.btnSaveFileName_Click);
            // 
            // btnSaveStatement
            // 
            this.btnSaveStatement.Location = new System.Drawing.Point(598, 7);
            this.btnSaveStatement.Name = "btnSaveStatement";
            this.btnSaveStatement.Size = new System.Drawing.Size(120, 20);
            this.btnSaveStatement.TabIndex = 14;
            this.btnSaveStatement.Text = "Create Statement File";
            this.btnSaveStatement.Click += new System.EventHandler(this.btnSaveStatement_Click);
            // 
            // lblVer
            // 
            this.lblVer.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblVer.BackColor = System.Drawing.Color.Transparent;
            this.lblVer.ForeColor = System.Drawing.Color.Blue;
            this.lblVer.Location = new System.Drawing.Point(888, 15);
            this.lblVer.Name = "lblVer";
            this.lblVer.Size = new System.Drawing.Size(25, 16);
            this.lblVer.TabIndex = 100;
            this.lblVer.Text = "Ver";
            this.lblVer.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.lblVer.Visible = false;
            // 
            // lblStatus
            // 
            this.lblStatus.BackColor = System.Drawing.Color.Transparent;
            this.lblStatus.ForeColor = System.Drawing.Color.Brown;
            this.lblStatus.Location = new System.Drawing.Point(-1, 264);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(701, 32);
            this.lblStatus.TabIndex = 8;
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // progBarStatus
            // 
            this.progBarStatus.Location = new System.Drawing.Point(65, 389);
            this.progBarStatus.Name = "progBarStatus";
            this.progBarStatus.Size = new System.Drawing.Size(569, 17);
            this.progBarStatus.Step = 0;
            this.progBarStatus.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progBarStatus.TabIndex = 11;
            // 
            // txtFileMD5
            // 
            this.txtFileMD5.Location = new System.Drawing.Point(708, 388);
            this.txtFileMD5.Name = "txtFileMD5";
            this.txtFileMD5.Size = new System.Drawing.Size(165, 20);
            this.txtFileMD5.TabIndex = 100;
            this.txtFileMD5.Visible = false;
            // 
            // btnGetFileMD5
            // 
            this.btnGetFileMD5.Location = new System.Drawing.Point(807, 386);
            this.btnGetFileMD5.Name = "btnGetFileMD5";
            this.btnGetFileMD5.Size = new System.Drawing.Size(76, 20);
            this.btnGetFileMD5.TabIndex = 100;
            this.btnGetFileMD5.Text = "Get File MD5";
            this.btnGetFileMD5.UseVisualStyleBackColor = true;
            this.btnGetFileMD5.Visible = false;
            this.btnGetFileMD5.Click += new System.EventHandler(this.btnGetFileMD5_Click);
            // 
            // txtFileName2
            // 
            this.txtFileName2.Location = new System.Drawing.Point(708, 364);
            this.txtFileName2.Name = "txtFileName2";
            this.txtFileName2.Size = new System.Drawing.Size(237, 20);
            this.txtFileName2.TabIndex = 100;
            this.txtFileName2.Text = "D:\\TEMP\\P20Files\\Statement\\_MerchantStatement\\SSB\\SSB_Mer_Stmt.xml";
            this.txtFileName2.Visible = false;
            // 
            // btnSendOprMail
            // 
            this.btnSendOprMail.Location = new System.Drawing.Point(598, 27);
            this.btnSendOprMail.Name = "btnSendOprMail";
            this.btnSendOprMail.Size = new System.Drawing.Size(120, 20);
            this.btnSendOprMail.TabIndex = 27;
            this.btnSendOprMail.Text = "Send Operation Mail";
            this.btnSendOprMail.UseVisualStyleBackColor = true;
            // 
            // btnSendBankMail
            // 
            this.btnSendBankMail.Location = new System.Drawing.Point(598, 49);
            this.btnSendBankMail.Name = "btnSendBankMail";
            this.btnSendBankMail.Size = new System.Drawing.Size(120, 20);
            this.btnSendBankMail.TabIndex = 28;
            this.btnSendBankMail.Text = "Send Bank Mail";
            this.btnSendBankMail.UseVisualStyleBackColor = true;
            this.btnSendBankMail.Click += new System.EventHandler(this.btnSendBankMail_Click);
            // 
            // chkBackupStatement
            // 
            this.chkBackupStatement.AutoSize = true;
            this.chkBackupStatement.BackColor = System.Drawing.Color.Transparent;
            this.chkBackupStatement.Location = new System.Drawing.Point(413, 29);
            this.chkBackupStatement.Name = "chkBackupStatement";
            this.chkBackupStatement.Size = new System.Drawing.Size(113, 17);
            this.chkBackupStatement.TabIndex = 29;
            this.chkBackupStatement.Text = "Backup Statement";
            this.chkBackupStatement.UseVisualStyleBackColor = false;
            // 
            // chkAppendData
            // 
            this.chkAppendData.AutoSize = true;
            this.chkAppendData.BackColor = System.Drawing.Color.Transparent;
            this.chkAppendData.Location = new System.Drawing.Point(413, 13);
            this.chkAppendData.Name = "chkAppendData";
            this.chkAppendData.Size = new System.Drawing.Size(89, 17);
            this.chkAppendData.TabIndex = 30;
            this.chkAppendData.Text = "Append Data";
            this.chkAppendData.UseVisualStyleBackColor = false;
            // 
            // datStmntData
            // 
            this.datStmntData.CustomFormat = "yyyy-MMMM-dd-dddd";
            this.datStmntData.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            this.datStmntData.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.datStmntData.Location = new System.Drawing.Point(449, 26);
            this.datStmntData.Name = "datStmntData";
            this.datStmntData.Size = new System.Drawing.Size(251, 20);
            this.datStmntData.TabIndex = 6;
            this.datStmntData.ValueChanged += new System.EventHandler(this.datStmntData_ValueChanged);
            // 
            // txtRunResult
            // 
            this.txtRunResult.Location = new System.Drawing.Point(0, 300);
            this.txtRunResult.Multiline = true;
            this.txtRunResult.Name = "txtRunResult";
            this.txtRunResult.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtRunResult.Size = new System.Drawing.Size(701, 85);
            this.txtRunResult.TabIndex = 9;
            // 
            // btnUpdateData
            // 
            this.btnUpdateData.Location = new System.Drawing.Point(728, 386);
            this.btnUpdateData.Name = "btnUpdateData";
            this.btnUpdateData.Size = new System.Drawing.Size(76, 20);
            this.btnUpdateData.TabIndex = 100;
            this.btnUpdateData.Text = "Update Data";
            this.btnUpdateData.UseVisualStyleBackColor = true;
            this.btnUpdateData.Visible = false;
            // 
            // btnMaintainBank
            // 
            this.btnMaintainBank.Location = new System.Drawing.Point(829, 323);
            this.btnMaintainBank.Name = "btnMaintainBank";
            this.btnMaintainBank.Size = new System.Drawing.Size(84, 20);
            this.btnMaintainBank.TabIndex = 100;
            this.btnMaintainBank.Text = "Maintain Bank";
            this.btnMaintainBank.UseVisualStyleBackColor = true;
            this.btnMaintainBank.Visible = false;
            // 
            // chkSendOprMail
            // 
            this.chkSendOprMail.BackColor = System.Drawing.Color.Transparent;
            this.chkSendOprMail.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkSendOprMail.ForeColor = System.Drawing.Color.DarkSlateGray;
            this.chkSendOprMail.Location = new System.Drawing.Point(726, 138);
            this.chkSendOprMail.Name = "chkSendOprMail";
            this.chkSendOprMail.Size = new System.Drawing.Size(146, 20);
            this.chkSendOprMail.TabIndex = 18;
            this.chkSendOprMail.Text = "Send Operation Mail";
            this.chkSendOprMail.UseVisualStyleBackColor = true;
            this.chkSendOprMail.Visible = false;
            // 
            // chkSendBankMail
            // 
            this.chkSendBankMail.BackColor = System.Drawing.Color.Transparent;
            this.chkSendBankMail.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkSendBankMail.ForeColor = System.Drawing.Color.DarkSlateGray;
            this.chkSendBankMail.Location = new System.Drawing.Point(726, 159);
            this.chkSendBankMail.Name = "chkSendBankMail";
            this.chkSendBankMail.Size = new System.Drawing.Size(133, 20);
            this.chkSendBankMail.TabIndex = 19;
            this.chkSendBankMail.Text = "Send Bank Mail";
            this.chkSendBankMail.UseVisualStyleBackColor = false;
            // 
            // chkLstProducts
            // 
            this.chkLstProducts.FormattingEnabled = true;
            this.chkLstProducts.Location = new System.Drawing.Point(0, 50);
            this.chkLstProducts.Name = "chkLstProducts";
            this.chkLstProducts.Size = new System.Drawing.Size(701, 214);
            this.chkLstProducts.TabIndex = 7;
            // 
            // btnAll
            // 
            this.btnAll.Location = new System.Drawing.Point(726, 182);
            this.btnAll.Name = "btnAll";
            this.btnAll.Size = new System.Drawing.Size(50, 20);
            this.btnAll.TabIndex = 23;
            this.btnAll.Text = "Go";
            this.btnAll.UseVisualStyleBackColor = true;
            this.btnAll.Click += new System.EventHandler(this.btnAll_Click);
            // 
            // chkSaveStatement
            // 
            this.chkSaveStatement.BackColor = System.Drawing.Color.Transparent;
            this.chkSaveStatement.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkSaveStatement.ForeColor = System.Drawing.Color.DarkSlateGray;
            this.chkSaveStatement.Location = new System.Drawing.Point(726, 96);
            this.chkSaveStatement.Name = "chkSaveStatement";
            this.chkSaveStatement.Size = new System.Drawing.Size(132, 20);
            this.chkSaveStatement.TabIndex = 16;
            this.chkSaveStatement.Text = "Create Statement";
            this.chkSaveStatement.UseVisualStyleBackColor = false;
            // 
            // chkDontPrompt
            // 
            this.chkDontPrompt.AutoSize = true;
            this.chkDontPrompt.BackColor = System.Drawing.Color.Transparent;
            this.chkDontPrompt.Checked = true;
            this.chkDontPrompt.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDontPrompt.Location = new System.Drawing.Point(726, 78);
            this.chkDontPrompt.Name = "chkDontPrompt";
            this.chkDontPrompt.Size = new System.Drawing.Size(87, 17);
            this.chkDontPrompt.TabIndex = 20;
            this.chkDontPrompt.Text = "don\'t prompt";
            this.chkDontPrompt.UseVisualStyleBackColor = false;
            // 
            // chkCheckEmail
            // 
            this.chkCheckEmail.AutoSize = true;
            this.chkCheckEmail.BackColor = System.Drawing.Color.Transparent;
            this.chkCheckEmail.Location = new System.Drawing.Point(328, 49);
            this.chkCheckEmail.Name = "chkCheckEmail";
            this.chkCheckEmail.Size = new System.Drawing.Size(82, 17);
            this.chkCheckEmail.TabIndex = 51;
            this.chkCheckEmail.Text = "Check Email";
            this.chkCheckEmail.UseVisualStyleBackColor = false;
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(78, 46);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(189, 20);
            this.txtPassword.TabIndex = 55;
            // 
            // txtUserName
            // 
            this.txtUserName.Location = new System.Drawing.Point(78, 7);
            this.txtUserName.Name = "txtUserName";
            this.txtUserName.Size = new System.Drawing.Size(189, 20);
            this.txtUserName.TabIndex = 54;
            // 
            // lblPassword
            // 
            this.lblPassword.BackColor = System.Drawing.Color.Transparent;
            this.lblPassword.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPassword.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lblPassword.Location = new System.Drawing.Point(11, 48);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(66, 17);
            this.lblPassword.TabIndex = 52;
            this.lblPassword.Text = "Password:";
            // 
            // lblUserName
            // 
            this.lblUserName.BackColor = System.Drawing.Color.Transparent;
            this.lblUserName.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblUserName.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lblUserName.Location = new System.Drawing.Point(11, 9);
            this.lblUserName.Name = "lblUserName";
            this.lblUserName.Size = new System.Drawing.Size(66, 17);
            this.lblUserName.TabIndex = 53;
            this.lblUserName.Text = "User Name:";
            // 
            // chkMaintainData
            // 
            this.chkMaintainData.AutoSize = true;
            this.chkMaintainData.BackColor = System.Drawing.Color.Transparent;
            this.chkMaintainData.Checked = true;
            this.chkMaintainData.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkMaintainData.Location = new System.Drawing.Point(411, 49);
            this.chkMaintainData.Name = "chkMaintainData";
            this.chkMaintainData.Size = new System.Drawing.Size(92, 17);
            this.chkMaintainData.TabIndex = 56;
            this.chkMaintainData.Text = "Maintain Data";
            this.chkMaintainData.UseVisualStyleBackColor = false;
            // 
            // chkExitAfterComplete
            // 
            this.chkExitAfterComplete.AutoSize = true;
            this.chkExitAfterComplete.BackColor = System.Drawing.Color.Transparent;
            this.chkExitAfterComplete.Location = new System.Drawing.Point(903, 75);
            this.chkExitAfterComplete.Name = "chkExitAfterComplete";
            this.chkExitAfterComplete.Size = new System.Drawing.Size(120, 17);
            this.chkExitAfterComplete.TabIndex = 100;
            this.chkExitAfterComplete.Text = "Exit After Complete";
            this.chkExitAfterComplete.UseVisualStyleBackColor = false;
            this.chkExitAfterComplete.Visible = false;
            // 
            // chkDontFixBranchData
            // 
            this.chkDontFixBranchData.AutoSize = true;
            this.chkDontFixBranchData.BackColor = System.Drawing.Color.Transparent;
            this.chkDontFixBranchData.Location = new System.Drawing.Point(903, 93);
            this.chkDontFixBranchData.Name = "chkDontFixBranchData";
            this.chkDontFixBranchData.Size = new System.Drawing.Size(127, 17);
            this.chkDontFixBranchData.TabIndex = 100;
            this.chkDontFixBranchData.Text = "don\'t fix Branch Data";
            this.chkDontFixBranchData.UseVisualStyleBackColor = false;
            this.chkDontFixBranchData.Visible = false;
            // 
            // txtTotProgress
            // 
            this.txtTotProgress.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtTotProgress.Location = new System.Drawing.Point(640, 388);
            this.txtTotProgress.Name = "txtTotProgress";
            this.txtTotProgress.Size = new System.Drawing.Size(60, 18);
            this.txtTotProgress.TabIndex = 12;
            // 
            // txtCurProgress
            // 
            this.txtCurProgress.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCurProgress.Location = new System.Drawing.Point(0, 388);
            this.txtCurProgress.Name = "txtCurProgress";
            this.txtCurProgress.Size = new System.Drawing.Size(60, 18);
            this.txtCurProgress.TabIndex = 10;
            // 
            // btnUnSelectAll
            // 
            this.btnUnSelectAll.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            this.btnUnSelectAll.Location = new System.Drawing.Point(8, 26);
            this.btnUnSelectAll.Name = "btnUnSelectAll";
            this.btnUnSelectAll.Size = new System.Drawing.Size(72, 20);
            this.btnUnSelectAll.TabIndex = 0;
            this.btnUnSelectAll.Text = "Unselect All";
            this.btnUnSelectAll.UseVisualStyleBackColor = true;
            this.btnUnSelectAll.Click += new System.EventHandler(this.btnUnSelectAll_Click);
            // 
            // lblBank
            // 
            this.lblBank.BackColor = System.Drawing.Color.Transparent;
            this.lblBank.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            this.lblBank.Location = new System.Drawing.Point(85, 2);
            this.lblBank.Name = "lblBank";
            this.lblBank.Size = new System.Drawing.Size(40, 16);
            this.lblBank.TabIndex = 1;
            this.lblBank.Text = "Bank";
            this.lblBank.Visible = false;
            // 
            // txtBank
            // 
            this.txtBank.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.txtBank.Location = new System.Drawing.Point(129, 0);
            this.txtBank.Name = "txtBank";
            this.txtBank.Size = new System.Drawing.Size(56, 20);
            this.txtBank.TabIndex = 3;
            this.txtBank.Visible = false;
            // 
            // lblBankCode
            // 
            this.lblBankCode.BackColor = System.Drawing.Color.Transparent;
            this.lblBankCode.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            this.lblBankCode.Location = new System.Drawing.Point(187, 2);
            this.lblBankCode.Name = "lblBankCode";
            this.lblBankCode.Size = new System.Drawing.Size(35, 16);
            this.lblBankCode.TabIndex = 4;
            this.lblBankCode.Text = "Code";
            this.lblBankCode.Visible = false;
            // 
            // txtBankCode
            // 
            this.txtBankCode.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.txtBankCode.Location = new System.Drawing.Point(220, 0);
            this.txtBankCode.Name = "txtBankCode";
            this.txtBankCode.Size = new System.Drawing.Size(44, 20);
            this.txtBankCode.TabIndex = 5;
            this.txtBankCode.Visible = false;
            // 
            // chkStatGenAfterMonth
            // 
            this.chkStatGenAfterMonth.AutoSize = true;
            this.chkStatGenAfterMonth.BackColor = System.Drawing.Color.Transparent;
            this.chkStatGenAfterMonth.Location = new System.Drawing.Point(903, 111);
            this.chkStatGenAfterMonth.Name = "chkStatGenAfterMonth";
            this.chkStatGenAfterMonth.Size = new System.Drawing.Size(141, 17);
            this.chkStatGenAfterMonth.TabIndex = 33;
            this.chkStatGenAfterMonth.Text = "BPC Debit Forward Only";
            this.chkStatGenAfterMonth.UseVisualStyleBackColor = false;
            this.chkStatGenAfterMonth.Visible = false;
            this.chkStatGenAfterMonth.CheckedChanged += new System.EventHandler(this.chkStatGenAfterMonth_CheckedChanged);
            // 
            // chkDevelopers
            // 
            this.chkDevelopers.BackColor = System.Drawing.Color.Transparent;
            this.chkDevelopers.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkDevelopers.ForeColor = System.Drawing.Color.DarkSlateGray;
            this.chkDevelopers.Location = new System.Drawing.Point(726, 118);
            this.chkDevelopers.Name = "chkDevelopers";
            this.chkDevelopers.Size = new System.Drawing.Size(174, 20);
            this.chkDevelopers.TabIndex = 17;
            this.chkDevelopers.Text = "Send Settlement Dev Mail";
            this.chkDevelopers.UseVisualStyleBackColor = false;
            // 
            // chkValidateBankEmail
            // 
            this.chkValidateBankEmail.AutoSize = true;
            this.chkValidateBankEmail.BackColor = System.Drawing.Color.Transparent;
            this.chkValidateBankEmail.Location = new System.Drawing.Point(726, 60);
            this.chkValidateBankEmail.Name = "chkValidateBankEmail";
            this.chkValidateBankEmail.Size = new System.Drawing.Size(117, 17);
            this.chkValidateBankEmail.TabIndex = 21;
            this.chkValidateBankEmail.Text = "Validate Bank Email";
            this.chkValidateBankEmail.UseVisualStyleBackColor = false;
            // 
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(970, 364);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(45, 20);
            this.btnLoad.TabIndex = 100;
            this.btnLoad.Text = "Load";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Visible = false;
            // 
            // chkMonitorFtp
            // 
            this.chkMonitorFtp.AutoSize = true;
            this.chkMonitorFtp.BackColor = System.Drawing.Color.Transparent;
            this.chkMonitorFtp.Location = new System.Drawing.Point(888, 386);
            this.chkMonitorFtp.Name = "chkMonitorFtp";
            this.chkMonitorFtp.Size = new System.Drawing.Size(83, 17);
            this.chkMonitorFtp.TabIndex = 100;
            this.chkMonitorFtp.Text = "Monitor FTP";
            this.chkMonitorFtp.UseVisualStyleBackColor = false;
            this.chkMonitorFtp.Visible = false;
            // 
            // chkUpdateSummary
            // 
            this.chkUpdateSummary.AutoSize = true;
            this.chkUpdateSummary.BackColor = System.Drawing.Color.Transparent;
            this.chkUpdateSummary.Checked = true;
            this.chkUpdateSummary.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkUpdateSummary.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.chkUpdateSummary.ForeColor = System.Drawing.Color.DarkSlateGray;
            this.chkUpdateSummary.Location = new System.Drawing.Point(726, 138);
            this.chkUpdateSummary.Name = "chkUpdateSummary";
            this.chkUpdateSummary.Size = new System.Drawing.Size(147, 17);
            this.chkUpdateSummary.TabIndex = 22;
            this.chkUpdateSummary.Text = "Update Summary Full";
            this.chkUpdateSummary.UseVisualStyleBackColor = false;
            // 
            // chkGenerateAtTime
            // 
            this.chkGenerateAtTime.AutoSize = true;
            this.chkGenerateAtTime.BackColor = System.Drawing.Color.Transparent;
            this.chkGenerateAtTime.Location = new System.Drawing.Point(904, 173);
            this.chkGenerateAtTime.Name = "chkGenerateAtTime";
            this.chkGenerateAtTime.Size = new System.Drawing.Size(15, 14);
            this.chkGenerateAtTime.TabIndex = 100;
            this.chkGenerateAtTime.UseVisualStyleBackColor = false;
            this.chkGenerateAtTime.Visible = false;
            // 
            // datGenerateAtTime
            // 
            this.datGenerateAtTime.CustomFormat = "yyyy-MM-dd hh:mm:ss";
            this.datGenerateAtTime.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            this.datGenerateAtTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.datGenerateAtTime.Location = new System.Drawing.Point(917, 169);
            this.datGenerateAtTime.Name = "datGenerateAtTime";
            this.datGenerateAtTime.ShowUpDown = true;
            this.datGenerateAtTime.Size = new System.Drawing.Size(125, 20);
            this.datGenerateAtTime.TabIndex = 100;
            this.datGenerateAtTime.Visible = false;
            // 
            // txtDbSchema
            // 
            this.txtDbSchema.BackColor = System.Drawing.SystemColors.Window;
            this.txtDbSchema.Location = new System.Drawing.Point(711, 208);
            this.txtDbSchema.Name = "txtDbSchema";
            this.txtDbSchema.ReadOnly = true;
            this.txtDbSchema.Size = new System.Drawing.Size(70, 20);
            this.txtDbSchema.TabIndex = 26;
            this.txtDbSchema.TextChanged += new System.EventHandler(this.txtDbSchema_TextChanged);
            this.txtDbSchema.Leave += new System.EventHandler(this.txtDbSchema_Leave);
            // 
            // txtTblMaster
            // 
            this.txtTblMaster.BackColor = System.Drawing.SystemColors.Window;
            this.txtTblMaster.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtTblMaster.Location = new System.Drawing.Point(711, 234);
            this.txtTblMaster.Name = "txtTblMaster";
            this.txtTblMaster.ReadOnly = true;
            this.txtTblMaster.Size = new System.Drawing.Size(206, 20);
            this.txtTblMaster.TabIndex = 24;
            this.txtTblMaster.Text = "TSTATEMENTMASTERTABLE";
            this.txtTblMaster.TextChanged += new System.EventHandler(this.txtTblMaster_TextChanged);
            // 
            // txtTblDetail
            // 
            this.txtTblDetail.BackColor = System.Drawing.SystemColors.Window;
            this.txtTblDetail.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtTblDetail.Location = new System.Drawing.Point(711, 262);
            this.txtTblDetail.Name = "txtTblDetail";
            this.txtTblDetail.ReadOnly = true;
            this.txtTblDetail.Size = new System.Drawing.Size(208, 20);
            this.txtTblDetail.TabIndex = 25;
            this.txtTblDetail.Text = "TSTATEMENTDETAILTABLE";
            this.txtTblDetail.TextChanged += new System.EventHandler(this.txtTblDetail_TextChanged);
            // 
            // chkStartCycle
            // 
            this.chkStartCycle.AutoSize = true;
            this.chkStartCycle.BackColor = System.Drawing.Color.Transparent;
            this.chkStartCycle.Location = new System.Drawing.Point(726, 7);
            this.chkStartCycle.Name = "chkStartCycle";
            this.chkStartCycle.Size = new System.Drawing.Size(79, 17);
            this.chkStartCycle.TabIndex = 100;
            this.chkStartCycle.Text = "Start Cycle";
            this.chkStartCycle.UseVisualStyleBackColor = false;
            this.chkStartCycle.Visible = false;
            // 
            // chkEndCycle
            // 
            this.chkEndCycle.AutoSize = true;
            this.chkEndCycle.BackColor = System.Drawing.Color.Transparent;
            this.chkEndCycle.Location = new System.Drawing.Point(726, 144);
            this.chkEndCycle.Name = "chkEndCycle";
            this.chkEndCycle.Size = new System.Drawing.Size(73, 17);
            this.chkEndCycle.TabIndex = 100;
            this.chkEndCycle.Text = "End Cycle";
            this.chkEndCycle.UseVisualStyleBackColor = false;
            this.chkEndCycle.Visible = false;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.chkCheckEmail);
            this.panel1.Controls.Add(this.chkBackupStatement);
            this.panel1.Controls.Add(this.btnSendOprMail);
            this.panel1.Controls.Add(this.btnSendBankMail);
            this.panel1.Controls.Add(this.chkAppendData);
            this.panel1.Controls.Add(this.btnSaveStatement);
            this.panel1.Controls.Add(this.lblUserName);
            this.panel1.Controls.Add(this.lblPassword);
            this.panel1.Controls.Add(this.txtUserName);
            this.panel1.Controls.Add(this.txtPassword);
            this.panel1.Controls.Add(this.chkMaintainData);
            this.panel1.Location = new System.Drawing.Point(2, 412);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(740, 79);
            this.panel1.TabIndex = 69;
            this.panel1.Visible = false;
            // 
            // chkTclientRestor
            // 
            this.chkTclientRestor.AutoSize = true;
            this.chkTclientRestor.BackColor = System.Drawing.Color.Transparent;
            this.chkTclientRestor.Location = new System.Drawing.Point(903, 58);
            this.chkTclientRestor.Name = "chkTclientRestor";
            this.chkTclientRestor.Size = new System.Drawing.Size(116, 17);
            this.chkTclientRestor.TabIndex = 100;
            this.chkTclientRestor.Text = "tClient Restoration";
            this.chkTclientRestor.UseVisualStyleBackColor = false;
            this.chkTclientRestor.Visible = false;
            // 
            // chkMaintain_Data
            // 
            this.chkMaintain_Data.BackColor = System.Drawing.Color.Transparent;
            this.chkMaintain_Data.Location = new System.Drawing.Point(904, 195);
            this.chkMaintain_Data.Name = "chkMaintain_Data";
            this.chkMaintain_Data.Size = new System.Drawing.Size(112, 20);
            this.chkMaintain_Data.TabIndex = 100;
            this.chkMaintain_Data.Text = "Maintain Data";
            this.chkMaintain_Data.UseVisualStyleBackColor = false;
            this.chkMaintain_Data.Visible = false;
            // 
            // chkStatPlanExec
            // 
            this.chkStatPlanExec.AutoSize = true;
            this.chkStatPlanExec.BackColor = System.Drawing.Color.Transparent;
            this.chkStatPlanExec.Location = new System.Drawing.Point(726, 141);
            this.chkStatPlanExec.Name = "chkStatPlanExec";
            this.chkStatPlanExec.Size = new System.Drawing.Size(151, 17);
            this.chkStatPlanExec.TabIndex = 100;
            this.chkStatPlanExec.Text = "Check Stat Plan Execution";
            this.chkStatPlanExec.UseVisualStyleBackColor = false;
            this.chkStatPlanExec.Visible = false;
            // 
            // chkExportData
            // 
            this.chkExportData.BackColor = System.Drawing.Color.Transparent;
            this.chkExportData.Location = new System.Drawing.Point(726, 62);
            this.chkExportData.Name = "chkExportData";
            this.chkExportData.Size = new System.Drawing.Size(112, 20);
            this.chkExportData.TabIndex = 14;
            this.chkExportData.Text = "Export Data";
            this.chkExportData.UseVisualStyleBackColor = false;
            this.chkExportData.Visible = false;
            // 
            // txtTablePrefix
            // 
            this.txtTablePrefix.BackColor = System.Drawing.Color.White;
            this.txtTablePrefix.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtTablePrefix.ForeColor = System.Drawing.Color.Black;
            this.txtTablePrefix.Location = new System.Drawing.Point(711, 288);
            this.txtTablePrefix.Name = "txtTablePrefix";
            this.txtTablePrefix.ReadOnly = true;
            this.txtTablePrefix.Size = new System.Drawing.Size(46, 21);
            this.txtTablePrefix.TabIndex = 27;
            this.txtTablePrefix.TextChanged += new System.EventHandler(this.txtTablePrefix_TextChanged);
            this.txtTablePrefix.Leave += new System.EventHandler(this.txtTablePrefix_Leave);
            // 
            // btnDumpName
            // 
            this.btnDumpName.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDumpName.Location = new System.Drawing.Point(992, 290);
            this.btnDumpName.Name = "btnDumpName";
            this.btnDumpName.Size = new System.Drawing.Size(16, 18);
            this.btnDumpName.TabIndex = 28;
            this.btnDumpName.Visible = false;
            // 
            // chkImportData
            // 
            this.chkImportData.BackColor = System.Drawing.Color.Transparent;
            this.chkImportData.Location = new System.Drawing.Point(726, 65);
            this.chkImportData.Name = "chkImportData";
            this.chkImportData.Size = new System.Drawing.Size(112, 20);
            this.chkImportData.TabIndex = 15;
            this.chkImportData.Text = "Import Data";
            this.chkImportData.UseVisualStyleBackColor = false;
            this.chkImportData.Visible = false;
            // 
            // chkRenameTables
            // 
            this.chkRenameTables.BackColor = System.Drawing.Color.Transparent;
            this.chkRenameTables.Location = new System.Drawing.Point(726, 24);
            this.chkRenameTables.Name = "chkRenameTables";
            this.chkRenameTables.Size = new System.Drawing.Size(112, 20);
            this.chkRenameTables.TabIndex = 13;
            this.chkRenameTables.Text = "Rename Tables";
            this.chkRenameTables.UseVisualStyleBackColor = false;
            this.chkRenameTables.CheckedChanged += new System.EventHandler(this.chkRenameTables_CheckedChanged);
            // 
            // chkUpdateSummaryPart
            // 
            this.chkUpdateSummaryPart.AutoSize = true;
            this.chkUpdateSummaryPart.BackColor = System.Drawing.Color.Transparent;
            this.chkUpdateSummaryPart.Location = new System.Drawing.Point(903, 148);
            this.chkUpdateSummaryPart.Name = "chkUpdateSummaryPart";
            this.chkUpdateSummaryPart.Size = new System.Drawing.Size(141, 17);
            this.chkUpdateSummaryPart.TabIndex = 100;
            this.chkUpdateSummaryPart.Text = "Update Summary Partial";
            this.chkUpdateSummaryPart.UseVisualStyleBackColor = false;
            this.chkUpdateSummaryPart.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(785, 293);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 13);
            this.label1.TabIndex = 101;
            this.label1.Text = "yyMMdd";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(787, 211);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(60, 13);
            this.label2.TabIndex = 102;
            this.label2.Text = "DB Schema";
            // 
            // chkTest
            // 
            this.chkTest.AutoSize = true;
            this.chkTest.BackColor = System.Drawing.Color.Transparent;
            this.chkTest.Location = new System.Drawing.Point(726, 42);
            this.chkTest.Name = "chkTest";
            this.chkTest.Size = new System.Drawing.Size(76, 17);
            this.chkTest.TabIndex = 103;
            this.chkTest.Text = "Test Mode";
            this.chkTest.UseVisualStyleBackColor = false;
            this.chkTest.CheckedChanged += new System.EventHandler(this.chkTest_CheckedChanged);
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Image = global::StatementFile.Properties.Resources._264;
            this.btnSelectFile.Location = new System.Drawing.Point(945, 364);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(23, 20);
            this.btnSelectFile.TabIndex = 100;
            this.btnSelectFile.Visible = false;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // btnExit
            // 
            this.btnExit.BackgroundImage = global::StatementFile.Properties.Resources.Close;
            this.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnExit.Location = new System.Drawing.Point(884, 29);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(33, 31);
            this.btnExit.TabIndex = 32;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(128, 25);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 21);
            this.comboBox1.TabIndex = 104;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            this.label3.Location = new System.Drawing.Point(86, 30);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(40, 16);
            this.label3.TabIndex = 105;
            this.label3.Text = "Cycle";
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(315, 25);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(121, 21);
            this.comboBox2.TabIndex = 106;
            this.comboBox2.SelectedIndexChanged += new System.EventHandler(this.comboBox2_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            this.label4.Location = new System.Drawing.Point(264, 30);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(51, 16);
            this.label4.TabIndex = 107;
            this.label4.Text = "E-STMT";
            // 
            // txtPrefixRun
            // 
            this.txtPrefixRun.BackColor = System.Drawing.Color.White;
            this.txtPrefixRun.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPrefixRun.ForeColor = System.Drawing.Color.Black;
            this.txtPrefixRun.Location = new System.Drawing.Point(759, 288);
            this.txtPrefixRun.Name = "txtPrefixRun";
            this.txtPrefixRun.Size = new System.Drawing.Size(22, 21);
            this.txtPrefixRun.TabIndex = 28;
            this.txtPrefixRun.Leave += new System.EventHandler(this.txtPrefixRun_Leave);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(837, 293);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(79, 13);
            this.label5.TabIndex = 109;
            this.label5.Text = "Run# (2 digits)";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.utilitiesToolStripMenuItem});
        this.menuStrip1.Location = new System.Drawing.Point(0, 0);
        this.menuStrip1.Name = "menuStrip1";
        this.menuStrip1.Size = new System.Drawing.Size(926, 24);
        this.menuStrip1.TabIndex = 110;
        this.menuStrip1.Text = "menuStrip1";
        // 
        // utilitiesToolStripMenuItem
        // 
        this.utilitiesToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.changeTypeToolStripMenuItem});
        this.utilitiesToolStripMenuItem.Name = "utilitiesToolStripMenuItem";
        this.utilitiesToolStripMenuItem.Size = new System.Drawing.Size(58, 20);
        this.utilitiesToolStripMenuItem.Text = "&Utilities";
        // 
        // changeTypeToolStripMenuItem
        // 
        this.changeTypeToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cReditToolStripMenuItem,
            this.corporateToolStripMenuItem,
            this.debitToolStripMenuItem});
        this.changeTypeToolStripMenuItem.Name = "changeTypeToolStripMenuItem";
        this.changeTypeToolStripMenuItem.Size = new System.Drawing.Size(143, 22);
        this.changeTypeToolStripMenuItem.Text = "&Change Type";
        // 
        // cReditToolStripMenuItem
        // 
        this.cReditToolStripMenuItem.Name = "cReditToolStripMenuItem";
        this.cReditToolStripMenuItem.Size = new System.Drawing.Size(127, 22);
        this.cReditToolStripMenuItem.Text = "C&redit";
        this.cReditToolStripMenuItem.Click += new System.EventHandler(this.creditToolStripMenuItem_Click);
        // 
        // corporateToolStripMenuItem
        // 
        this.corporateToolStripMenuItem.Name = "corporateToolStripMenuItem";
        this.corporateToolStripMenuItem.Size = new System.Drawing.Size(127, 22);
        this.corporateToolStripMenuItem.Text = "C&orporate";
        this.corporateToolStripMenuItem.Click += new System.EventHandler(this.corporateToolStripMenuItem_Click);
        // 
        // debitToolStripMenuItem
        // 
        this.debitToolStripMenuItem.Name = "debitToolStripMenuItem";
        this.debitToolStripMenuItem.Size = new System.Drawing.Size(127, 22);
        this.debitToolStripMenuItem.Text = "De&bit";
        this.debitToolStripMenuItem.Click += new System.EventHandler(this.debitToolStripMenuItem_Click);
        // 
        // button1
        // 
        this.button1.BackColor = System.Drawing.Color.Red;
        this.button1.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
        this.button1.Location = new System.Drawing.Point(782, 182);
        this.button1.Name = "button1";
        this.button1.Size = new System.Drawing.Size(116, 20);
        this.button1.TabIndex = 111;
        this.button1.Text = "External Mode";
        this.button1.UseVisualStyleBackColor = false;
        this.button1.Visible = false;
        this.button1.Click += new System.EventHandler(this.button1_Click);
        // 
        // btnBanksMails
        // 
        this.btnBanksMails.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnBanksMails.BackgroundImage")));
        this.btnBanksMails.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
        this.btnBanksMails.Location = new System.Drawing.Point(849, 29);
        this.btnBanksMails.Name = "btnBanksMails";
        this.btnBanksMails.Size = new System.Drawing.Size(34, 30);
        this.btnBanksMails.TabIndex = 112;
        this.btnBanksMails.UseVisualStyleBackColor = true;
        this.btnBanksMails.Click += new System.EventHandler(this.btnBanksMails_Click);
        // 
        // CombineFiles
        // 
        this.CombineFiles.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("CombineFiles.BackgroundImage")));
        this.CombineFiles.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
        this.CombineFiles.Location = new System.Drawing.Point(849, 62);
        this.CombineFiles.Name = "CombineFiles";
        this.CombineFiles.Size = new System.Drawing.Size(34, 30);
        this.CombineFiles.TabIndex = 113;
        this.CombineFiles.UseVisualStyleBackColor = true;
        this.CombineFiles.Click += new System.EventHandler(this.CombineFiles_Click);
        // 
        // frmStatementFile
        // 
        this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
        this.ClientSize = new System.Drawing.Size(926, 411);
        this.Controls.Add(this.CombineFiles);
        this.Controls.Add(this.btnBanksMails);
        this.Controls.Add(this.button1);
        this.Controls.Add(this.label5);
        this.Controls.Add(this.txtPrefixRun);
        this.Controls.Add(this.chkDontPrompt);
        this.Controls.Add(this.chkUpdateSummary);
        this.Controls.Add(this.chkSaveStatement);
        this.Controls.Add(this.chkDevelopers);
        this.Controls.Add(this.chkSendBankMail);
        this.Controls.Add(this.label4);
        this.Controls.Add(this.comboBox2);
        this.Controls.Add(this.label3);
        this.Controls.Add(this.comboBox1);
        this.Controls.Add(this.chkTest);
        this.Controls.Add(this.label2);
        this.Controls.Add(this.label1);
        this.Controls.Add(this.chkUpdateSummaryPart);
        this.Controls.Add(this.chkTclientRestor);
        this.Controls.Add(this.panel1);
        this.Controls.Add(this.datGenerateAtTime);
        this.Controls.Add(this.chkMonitorFtp);
        this.Controls.Add(this.btnLoad);
        this.Controls.Add(this.btnUnSelectAll);
        this.Controls.Add(this.txtBankCode);
        this.Controls.Add(this.txtBank);
        this.Controls.Add(this.txtCurProgress);
        this.Controls.Add(this.txtTotProgress);
        this.Controls.Add(this.txtTblDetail);
        this.Controls.Add(this.txtTblMaster);
        this.Controls.Add(this.txtDbSchema);
        this.Controls.Add(this.chkImportData);
        this.Controls.Add(this.chkRenameTables);
        this.Controls.Add(this.chkExportData);
        this.Controls.Add(this.chkMaintain_Data);
        this.Controls.Add(this.btnAll);
        this.Controls.Add(this.chkLstProducts);
        this.Controls.Add(this.datStmntData);
        this.Controls.Add(this.txtFileName2);
        this.Controls.Add(this.btnUpdateData);
        this.Controls.Add(this.btnGetFileMD5);
        this.Controls.Add(this.txtRunResult);
        this.Controls.Add(this.txtFileMD5);
        this.Controls.Add(this.btnMaintainBank);
        this.Controls.Add(this.progBarStatus);
        this.Controls.Add(this.lblStatus);
        this.Controls.Add(this.btnSelectFile);
        this.Controls.Add(this.btnDumpName);
        this.Controls.Add(this.btnSaveFileName);
        this.Controls.Add(this.txtTablePrefix);
        this.Controls.Add(this.txtFileName);
        this.Controls.Add(this.btnExit);
        this.Controls.Add(this.lblBankCode);
        this.Controls.Add(this.lblBank);
        this.Controls.Add(this.lblStatementPath);
        this.Controls.Add(this.chkGenerateAtTime);
        this.Controls.Add(this.chkEndCycle);
        this.Controls.Add(this.chkValidateBankEmail);
        this.Controls.Add(this.chkStatGenAfterMonth);
        this.Controls.Add(this.chkDontFixBranchData);
        this.Controls.Add(this.chkExitAfterComplete);
        this.Controls.Add(this.lblVer);
        this.Controls.Add(this.chkSendOprMail);
        this.Controls.Add(this.chkStatPlanExec);
        this.Controls.Add(this.chkStartCycle);
        this.Controls.Add(this.menuStrip1);
        this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
        this.MainMenuStrip = this.menuStrip1;
        this.MaximizeBox = false;
        this.Name = "frmStatementFile";
        this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
        this.Text = "Statement File";
        this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmStatementFile_FormClosing);
        this.Load += new System.EventHandler(this.frmStatementFile_Load);
        this.Resize += new System.EventHandler(this.frmStatementFile_Resize);
        this.panel1.ResumeLayout(false);
        this.panel1.PerformLayout();
        this.menuStrip1.ResumeLayout(false);
        this.menuStrip1.PerformLayout();
        this.ResumeLayout(false);
        this.PerformLayout();

    }
    #endregion

    private void btnSaveFileName_Click(object sender, System.EventArgs e)
    {
        // Show the FolderBrowserDialog.
        //txtFileName.Text = clsBasFile.openDirDialog();


        //DialogResult result = fldSelectFolder.ShowDialog();
        //if( result == DialogResult.OK )
        //    txtFileName.Text = clsBasFile.makeStrAsPath(fldSelectFolder.SelectedPath); 
        //System.Windows.Forms.BrowseForFolderForm F = new System.Windows.Forms.BrowseForFolderForm(txtFileName.Text);

        //if (F.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        //{
        //    txtFileName.Text = System.Windows.Forms.BrowseForFolderForm.m_FolderPath;
        //}

        //txtFileName.Text = clsBasFile.makeStrAsPath(clsBasFile.openDirDialog("Open Path"));
        //txtFileName.Text = clsBasFile.openDirDialog1(this);
        //txtFileName.Text = clsBasFile.makeStrAsPath(clsBasFile.openDirDialog("Open Path"));
        string strPath = Directory.Exists(clsBasFile.getPathWithoutFile(txtFileName.Text)) ? clsBasFile.getPathWithoutFile(txtFileName.Text) : @"D:\";
        txtFileName.Text = clsBasFile.makeStrAsPath(clsBasFile.openDirDialog("Open Path", strPath));
    }

    private void frmStatementFile_Load(object sender, System.EventArgs e)
    {
        GeneratePrefix();
        strServer = clsDbCon.sServer;
        strUserName = clsDbCon.sUserName;
        inf = new AssemblyInfo();

        this.Text = "Statement File " + strUserName + "@" + clsBasNetwork.getMyComputerName() + "@" + clsBasNetwork.getMyComputerIP() + " >> " + strServer + " " + clsBasUserData.loginDate.ToString("dd/MM/yyyy hh:mm:ss") + " - " + Application.StartupPath + " - Version:" + inf.Version + " - " + clsBasUserData.sType;
        //b4FrmLoad();
        showToolTips();
        if (System.IO.File.Exists(Application.StartupPath + "\\frmBackground.jpg"))
        {
            this.BackgroundImage = Image.FromFile(Application.StartupPath + "\\frmBackground.jpg");
            this.menuStrip1.BackgroundImage = Image.FromFile(Application.StartupPath + "\\frmBackground.jpg");
            //grpBoxClientEmails.BackgroundImage = Image.FromFile(Application.StartupPath + "\\frmBackground.jpg");
        }
        DateTime cDate;
        cDate = datStmntData.Value;
        datGenerateAtTime.Value = DateTime.Now;
        Calendar myCals = CultureInfo.InvariantCulture.Calendar;
        //myCals.GetDaysInMonth(cDate);
        //if (myCals.GetDaysInMonth(cDate.Year, cDate.Month) - cDate.Day >= 0 && myCals.GetDaysInMonth(cDate.Year, cDate.Month) - cDate.Day < 4)
        //    {
        //    datStmntData.Value = datStmntData.Value.AddMonths(1);
        //    cDate = datStmntData.Value;
        //    datStmntData.Value = datStmntData.Value.AddDays(-(cDate.Day - 1));
        //    } // EDT-728
        txtUserName.Text = clsDbCon.sUserName;
        if (clsBasUserData.sExternal == true)
            txtDbSchema.Text = clsSessionValues.workingDbSchema;
        else
            txtDbSchema.Text = txtUserName.Text + ".";
        //System.DateTime newMonth = new System.DateTime(cDate.AddMonths(1).Year, cDate.AddMonths(1).Month, 1, 1, 1, 1, 11);
        //if (newMonth.DayOfYear - cDate.DayOfYear > -1 && newMonth.DayOfYear - cDate.DayOfYear < 4)
        //{
        //  datStmntData.Value = datStmntData.Value.AddMonths(1);
        //  datStmntData.Value = datStmntData.Value.AddDays( -(cDate.Day-1));
        //}
        //clsCheckEmail checkEmail = new clsCheckEmail("mmohammed", "Apach@nt58", "192.168.1.20");
        //checkEmail.startCheck(3000);
        //runTimer(6000);
        setStatusDelegate = new StatusDelegate(statusUpdate);
        setProgressDelegate = new SetProgressDelegate(setProgress);
        setMinMaxProgressDelegate = new SetMinMaxProgressDelegate(setMinMaxProgress);
        makeAppTray();
        ChangeType();
    }

    private void btnSaveStatement_Click(object sender, System.EventArgs e)
    {
        //Cursor.Current = Cursors.WaitCursor;
        lblStatus.ForeColor = Color.DarkRed;
        lblStatus.Text = "Please wait will Creating Statement File";
        btnSaveStatement.Enabled = false;
        curIndx = ((ListItem)chkLstProducts.Items[chkLstProducts.SelectedIndex]).ID;//chkLstProducts.SelectedIndex;
                                                                                    // Start the asynchronous operation.
                                                                                    //curCmbProdInx = curIndx; // cmbProducts.SelectedIndex
        backgroundWorker = new BackgroundWorker();
        backgroundWorker.WorkerSupportsCancellation = true;
        backgroundWorker.DoWork += new DoWorkEventHandler(backgroundWorker_DoWork);
        backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker_RunWorkerCompleted);
        //backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
        backgroundWorker.RunWorkerAsync();
        //btnCancelGeneration.Enabled = true;
        //Cursor.Current = Cursors.Default;
    }


    private void showToolTips()
    {
        ToolTip frmToolTip = new ToolTip();

        // Set up the delays for the ToolTip.
        frmToolTip.AutoPopDelay = 5000;
        frmToolTip.InitialDelay = 1000;
        frmToolTip.ReshowDelay = 500;
        // Force the ToolTip text to be displayed whether or not the form is active.
        frmToolTip.ShowAlways = true;

        // Set up the ToolTip text for the Button and Checkbox.
        frmToolTip.SetToolTip(this.btnExit, "Exit The Program");
        frmToolTip.SetToolTip(this.btnSaveStatement, "Extract Statement to Statement Path");
        frmToolTip.SetToolTip(this.btnSaveFileName, "Select Statement File Path");
        frmToolTip.SetToolTip(this.txtFileName, "Statement File Path");
        frmToolTip.SetToolTip(this.chkLstProducts, "Select Product of Statement File");
    }



    private void validateBankEmail()
    {
        setCurrentBank();
        BeginInvoke(setStatusDelegate, new object[] { bankCode, "Start Creat Statement for " + strStatementType });//strStatementType
        clsValidateBankEmail vldBankEmail = new clsValidateBankEmail();
        vldBankEmail.setFrm = this;
        vldBankEmail.Statement(txtFileName.Text, bankName, bankCode, strFileName, datStmntData.Value, stmntClientEmail, appendData);
        vldBankEmail = null;
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

    //private void btnCancelGeneration_Click(object sender, EventArgs e)
    //    {
    //    // Cancel the asynchronous operation.
    //    this.backgroundWorker.CancelAsync();

    //    // Disable the Cancel button.
    //    //btnCancelGeneration.Enabled = false;
    //    }

    private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
    {
        // Get the BackgroundWorker that raised this event.
        //BackgroundWorker worker = sender as BackgroundWorker;

        // Assign the result of the computation
        // to the Result property of the DoWorkEventArgs
        // object. This is will be available to the 
        // RunWorkerCompleted eventhandler.
        runStatement(curIndx);//curCmbProdInx
    }

    private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
        lblStatus.ForeColor = Color.Blue;
        lblStatus.Text = "Statement File Creation Done for " + strStatementType + " \r\n" + checkErrRslt;
        txtRunResult.Text = lblStatus.Text + "\r\n\r\n" + basText.Left(txtRunResult.Text, 2000);
        btnSaveStatement.Enabled = true;
        //btnCancelGeneration.Enabled = false;


        checkErrRslt = string.Empty;
    }

    private void btnGetFileMD5_Click(object sender, EventArgs e)
    {
        clsStatementSummary StatSummary = new clsStatementSummary();
        StatSummary.BankCode = 2;
        StatSummary.BankName = "SSB";
        StatSummary.StatementDate = DateTime.Now;
        StatSummary.CreationDate = DateTime.Now;
        StatSummary.StatementProduct = "Credit";
        StatSummary.StatementType = "pdf";
        StatSummary.NoOfStatements = 5;
        StatSummary.NoOfPages = 10;
        StatSummary.NoOfTransactions = 15;
        //StatSummary.InsertRecord();
        StatSummary = null;

        //clsBasFile.receiveFileFromFtp("merchantstatement.pdf",@"\_Receive\SSB",@"D:\TEMP\P20Files\Statement\_MerchantStatement");

        //System.Diagnostics.Process.Start(@"C:\Program Files\SSH Communications Security\SSH Tectia\SSH Tectia Client\" + "sftpg3", @" -B D:\TEMP\P20Files\Statement\SendFTP.txt --password=Mscc1234  mmohammed@192.168.3.10");
        //System.Diagnostics.Process.Start(@"C:\Program Files\SSH Communications Security\SSH Tectia\SSH Tectia Client\" + "sftpg3", " -B " + getPathWithoutFile(pFile) + @"\SendFTP.txt" + " --password=Mscc1234  mmohammed@192.168.3.10");

        //clsBasFile.sendFile2ftp(@"D:\TEMP\P20Files\Statement\_200810F\20081101\200810NSGB_Credit\NSGB_Credit_Statement_File_200810.zip");

        //FileStream fileSummary;
        //StreamWriter streamSummary;
        //fileSummary = new FileStream(@"D:\TEMP\P20Files\Statement\SendFTP.txt", FileMode.Create); //Create
        //streamSummary = new StreamWriter(fileSummary, System.Text.Encoding.Default);
        //streamSummary.WriteLine("sput \"C:\\CardHolderLayout.txt\" \"CardHolderLayout.txt\"");
        //streamSummary.Flush();
        //streamSummary.Close();
        //fileSummary.Close();
        //System.Diagnostics.Process.Start(@"C:\Program Files\SSH Communications Security\SSH Tectia\SSH Tectia Client\" + "sftpg3", " -B " + @"D:\TEMP\P20Files\Statement\SendFTP.txt" + " --password=Mscc1234  mmohammed@192.168.3.10 "); 

        //clsFillLst.fillLst("Select serno, loadtime, loaduser From Batches Where direction = 'T' And filesource = 'statements' And product = 'BM-V-CR-GOLD' Order by 2 desc", lvwBatches);

        //sendOutlookEmail();
        //clsOMR omr = new clsOMR();
        //txtFileMD5.Text = omr.Asterisk(Convert.ToInt32(txtFileName2.Text));

        //txtFileMD5.Text = clsBasFile.getFileMD5(txtFileName2.Text);
        //clsMaintainData maintainData = new clsMaintainData();
        //maintainData.mergeTrans(5);

    }





    private void btnSendBankMail_Click(object sender, EventArgs e)
    {
        SendBankMail();
    }

    private void SendBankMail()
    {
        //string strFileName;
        //DateTime vCurDate;
        clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + "start sending bank mail function", System.IO.FileMode.Append);

        setCurrentBank();
        if (!chkDontPrompt.Checked)
        {
            if (DialogResult.No == MessageBox.Show("Are you sure you want send bank mail for " + stmntMail + ".", "Send bank Mail", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2))
            {
                return;
            }
        }
        //add2StatExec();
        //return;
        clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + "Start Sent Email to the Bank for " + strStatementType, System.IO.FileMode.Append);

        BeginInvoke(setStatusDelegate, new object[] { bankCode, "Start Sent Email to the Bank for " + strStatementType });//strStatementType
        //Cursor.Current = Cursors.WaitCursor;
        //>lblStatus.ForeColor = Color.DarkRed;
        //>lblStatus.Text = "Sending Email";
        //    strFileName = clsBasFile.makeStrAsPath(txtFileName.Text);
        //if (bankName == "BIC" || bankName == "AUB")
        //{
        //  vCurDate = DateTime.Now;//.AddMonths(-1)
        //}
        //else
        //{
        //  vCurDate = DateTime.Now.AddMonths(-1);
        //}
        //clsBasFile.createDirectory(strFileName + vCurDate.ToString("yyyyMM") + bankName);
        //strFileName = clsBasFile.makeStrAsPath(txtFileName.Text) + vCurDate.ToString("yyyyMM") + bankName + "\\" + bankName + strFileName + vCurDate.ToString("yyyyMM") + ".zip";
        //strFileName = clsBasFile.makeStrAsPath(txtFileName.Text) + stmntDate.ToString("yyyyMM") + strStatementType + "\\" + strStatementType + strFileName + stmntDate.ToString("yyyyMM") + ".zip";
        //string strDay = "15";
        //if (stmntDate.Day >= 25 || stmntDate.Day < 10)
        //    strDay = "01";

        strFileName = clsBasFile.makeStrAsPath(txtFileName.Text) + "_WaitReply\\" + stmntDate.ToString("yyyyMM") + strStatementType + "\\" + strStatementType + strFileName + stmntDate.ToString("yyyyMM") + ".zip";
        //strFileName = clsBasFile.makeStrAsPath(txtFileName.Text) + "_" + stmntDate.ToString("yyyy") + stmntDate.ToString("MM") + "F\\" + (strDay == "01" ? stmntDate.AddMonths(1).ToString("yyyy") : stmntDate.ToString("yyyy")) + (strDay == "01" ? stmntDate.AddMonths(1).ToString("MM") : stmntDate.ToString("MM")) + strDay + "\\" + stmntDate.ToString("yyyy") + stmntDate.ToString("MM") + strStatementType + "\\" + strStatementType + strFileName + stmntDate.ToString("yyyyMM") + ".zip";
        //strFileName = clsBasFile.makeStrAsPath(txtFileName.Text) + "_WaitReply\\" + stmntDate.ToString("yyyy") + stmntDate.ToString("MM") + strStatementType + "\\" + strStatementType + strFileName + stmntDate.ToString("yyyyMM") + ".zip";
        //string strMD5Name;
        //strMD5Name = clsBasFile.makeStrAsPath(txtFileName.Text) + stmntDate.ToString("yyyyMM") + strStatementType + "\\" + strStatementType + strFileName + stmntDate.ToString("yyyyMM") + ".zip";

        clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + strFileName ?? "" + Environment.NewLine + "before send bank mail" + strStatementType, System.IO.FileMode.Append);

        clsMailStatement sendmail = new clsMailStatement();
        EventLog.WriteEntry("STMT", "before send bank mail", EventLogEntryType.Error);

        if (sendmail.sendBankMail(stmntMail, stmntDate, strFileName, strStatementType, txtFileName.Text))  //, strMD5Name  CstmntDate.ToString("MM/yyyy") :\TEMP\P20Files\Statement\200708NSGB\Old\NSGB_Credit_Statement_200708.zip
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + "Email Sent to the Bank for " + strStatementType, System.IO.FileMode.Append);
            EventLog.WriteEntry("STMT", "Email Sent to the Bank for " + strStatementType, EventLogEntryType.Error);

            BeginInvoke(setStatusDelegate, new object[] { bankCode, "Email Sent to the Bank for " + strStatementType });//strStatementType
        }
        else
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + strFileName ?? "" + Environment.NewLine + "Error on sending Bank Email for" + strStatementType, System.IO.FileMode.Append);
            EventLog.WriteEntry("STMT", "Error on sending Bank Email for " + strStatementType, EventLogEntryType.Error);

            BeginInvoke(setStatusDelegate, new object[] { bankCode, "Error on sending Bank Email for " + strStatementType });//strStatementType
        }
        Cursor.Current = Cursors.Default;
    }



    private void btnSelectFile_Click(object sender, EventArgs e)
    {
        string strFileName;
        strFileName = clsBasFile.openFileDialog("All Files (*.*)|*.*");
        if (strFileName.Trim().Length > 1)
            txtFileName2.Text = strFileName;

    }



    private void btnAll_Click(object sender, EventArgs e)
    {
        progBarStatus.Value = 0;
        clsBasFile.Filepath = txtFileName.Text;
        if (clsBasUserData.sType == "CR")
            clsBasFile.TableName = "CR_" + txtTablePrefix.Text + txtPrefixRun.Text;
        else if (clsBasUserData.sType == "CP")
            clsBasFile.TableName = "CP_" + txtTablePrefix.Text + txtPrefixRun.Text;
        else if (clsBasUserData.sType == "DB")
            clsBasFile.TableName = "DB_" + txtTablePrefix.Text + txtPrefixRun.Text;
        if (chkSaveStatement.Checked)
            if (txtPrefixRun.Text == "" && !txtDbSchema.Text.StartsWith("A4M"))
            {
                MessageBox.Show("Enter run# sequence!");
                return;
            }

        createAllStat();
    }

    private void createAllStat()
    {
        System.ComponentModel.BackgroundWorker generateAllSelected;

        //Cursor.Current = Cursors.WaitCursor;
        lblStatus.ForeColor = Color.DarkRed;
        //lblStatus.Text = "Please wait will Creating Statement File";
        //btnAll.Enabled = false;
        // Start the asynchronous operation.
        //curCmbProdInx = curIndx; // cmbProducts.SelectedIndex
        generateAllSelected = new BackgroundWorker();
        generateAllSelected.WorkerSupportsCancellation = true;
        generateAllSelected.DoWork += new DoWorkEventHandler(generateAllSelected_DoWork);
        generateAllSelected.RunWorkerCompleted += new RunWorkerCompletedEventHandler(generateAllSelected_RunWorkerCompleted);
        //generateAllSelected.ProgressChanged += new ProgressChangedEventHandler(generateAllSelected_ProgressChanged);
        generateAllSelected.RunWorkerAsync();
        //btnCancelGeneration.Enabled = true;
        //btnAll.Enabled = true;
        //chkLstProducts.SetItemChecked(curIndx, false); // = CheckState.Checked;
        //chkLstProducts.SelectedIndex = curIndx;

        GC.Collect();
        GC.WaitForPendingFinalizers();
        //Cursor.Current = Cursors.Default;
    }

    private void runAllSelected()
    {
        //if (chkStartCycle.Checked)
        //    startCycleProc();
        if (chkRenameTables.Checked)
            renameTables();
        //if (chkExportData.Checked)
        //    exportData();
        //if (chkMaintain_Data.Checked)
        //    maintainData();
        Cursor.Current = Cursors.WaitCursor;
        foreach (int indexChecked in chkLstProducts.CheckedIndices)
        {
            System.Diagnostics.Stopwatch stopWatch = new System.Diagnostics.Stopwatch();
            stopWatch.Start();
            //BeginInvoke(setStatusDelegate, new object[] { bankCode, "------------------------------" });//strStatementType
            //BeginInvoke(setStatusDelegate, new object[] { bankCode, "Start Date : " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") });//strStatementType
            curIndxVal = indexChecked;
            curIndx = ((ListItem)chkLstProducts.Items[indexChecked]).ID;//indexChecked;
            setCurrentBank();
            //if (clsBasStatement.getBasicData(bankCode))
            //    {
            //curIndx = ((ListItem)chkLstProducts.Items[indexChecked]).ID;//indexChecked;
            //if (chkSaveStatement.Checked)
            //    runStatement(curIndx);
            if (chkSaveStatement.Checked)
            {
                clsBasStatement clsBasStatement = new clsBasStatement();
                if (clsBasStatement.getBasicData(bankCode, StDate, strStatementType, clsBasFile.TableName))
                {
                    runStatement(curIndx);
                }
                //runStatement(curIndx);
            }
            if (chkDevelopers.Checked)
                SendConfEmailDev();
            //if (chkSendOprMail.Checked)
            //    SendOprMail();
            if (chkSendBankMail.Checked)
                SendBankMail();
            if (chkValidateBankEmail.Checked)
                validateBankEmail();

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;
            //txtRunResult.Text = lblStatus.Text + "\r\n\r\n" + basText.Left(String.Format("{0}", h.ElapsedTime), 2000);
            BeginInvoke(setStatusDelegate, new object[] { bankCode, "End Date : " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") });//strStatementType
            BeginInvoke(setStatusDelegate, new object[] { bankCode, "Elapsed Time : " + String.Format("{0:00}:{1:00}:{2:00}.{3:000} HH:MM:SS:Mil", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds) });//strStatementType

            //if (chkSendBankMail.Checked)
            //  SendBankMail();

            //lblStatus.Text = "Statement File Creation Done for " + strStatementType + " \r\n" + checkErrRslt;
            //txtRunResult.Text = lblStatus.Text + "\r\n\r\n" + basText.Left(txtRunResult.Text, 2000);

            //BeginInvoke(statusDelegate, new object[] { bankCode });
            //}
            BeginInvoke(setStatusDelegate, new object[] { bankCode, "------------------------------" });//strStatementType
                                                                                                        //if (chkEndCycle.Checked)
                                                                                                        //    endCycleProc();
                                                                                                        //if (chkTclientRestor.Checked)
                                                                                                        //    tClientRestor();
                                                                                                        //if (chkStatPlanExec.Checked)
                                                                                                        //    statPlanExec();
                                                                                                        //if (chkImportData.Checked)
                                                                                                        //    importData();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private void generateAllSelected_DoWork(object sender, DoWorkEventArgs e)
    {
        runAllSelected(); //runAllSelected(curIndx);//curCmbProdInx
    }

    private void generateAllSelected_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
        //lblStatus.ForeColor = Color.Blue;
        //lblStatus.Text = "Statement File Creation Done for " + strStatementType + " \r\n" + checkErrRslt;
        //txtRunResult.Text = lblStatus.Text + "\r\n\r\n" + basText.Left(txtRunResult.Text, 2000);
        //btnSaveStatement.Enabled = true;
        //btnCancelGeneration.Enabled = false;

        clsBasUserData.b4CloseProgram();
        Cursor.Current = Cursors.Default;

        if (chkExitAfterComplete.Checked)
            this.Close();

        //checkErrRslt = string.Empty;
    }



    private void SendConfEmailDev()
    {
        //string strFileName;
        //DateTime vCurDate;

        setCurrentBank();
        if (!chkDontPrompt.Checked)
            if (DialogResult.No == MessageBox.Show("Are you sure you want send Settlement mail for " + bankName + ".", "Send Settlement Mail",
               MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2))
                return;

        BeginInvoke(setStatusDelegate, new object[] { bankCode, "Start Send Email to Settlement for " + strStatementType });//strStatementType
                                                                                                                            //Cursor.Current = Cursors.WaitCursor;
                                                                                                                            //>lblStatus.ForeColor = Color.DarkRed;
                                                                                                                            //>lblStatus.Text = "Sending Email";
                                                                                                                            //    strFileName = clsBasFile.makeStrAsPath(txtFileName.Text);
                                                                                                                            //if (bankName == "BIC" || bankName == "AUB")
                                                                                                                            //{
                                                                                                                            //  vCurDate = DateTime.Now;//.AddMonths(-1)
                                                                                                                            //}
                                                                                                                            //else
                                                                                                                            //{
                                                                                                                            //  vCurDate = DateTime.Now.AddMonths(-1);
                                                                                                                            //}
                                                                                                                            //clsBasFile.createDirectory(strFileName + vCurDate.ToString("yyyyMM") + bankName);
                                                                                                                            //>strFileName = clsBasFile.makeStrAsPath(txtFileName.Text) + vCurDate.ToString("yyyyMM") + bankName + "\\" + bankName + strFileName + vCurDate.ToString("yyyyMM") + ".zip";
                                                                                                                            //strFileName = clsBasFile.makeStrAsPath(txtFileName.Text) + stmntDate.ToString("yyyyMM") + bankName + "\\" + bankName + strFileName + stmntDate.ToString("yyyyMM") + ".zip";
        strFileName = clsBasFile.makeStrAsPath(txtFileName.Text) + stmntDate.ToString("yyyyMM") + strStatementType + "\\" + strStatementType + strFileName + stmntDate.ToString("yyyyMM") + ".zip";



        clsMailStatement sendmail = new clsMailStatement();
        //if (sendmail.sendOprationMail(bankName, stmntDate.ToString("MM/yyyy"), strFileName))  //D:\TEMP\P20Files\Statement\200708NSGB\Old\NSGB_Credit_Statement_200708.zip
        //if (sendmail.sendDevelopersMail(bankCode.ToString(),strStatementType, stmntDate.ToString("MM/yyyy"), strFileName, txtFileName.Text))  // stmntDate.ToString("MM/yyyy")  D:\TEMP\P20Files\Statement\200708NSGB\Old\NSGB_Credit_Statement_200708.zip
        if (sendmail.sendDevelopersMail(bankCode.ToString(), strStatementType, stmntDate.ToString("MM/yyyy"), stmntDate, strFileName, txtFileName.Text, clsBasFile.TableName))  // stmntDate.ToString("MM/yyyy")  D:\TEMP\P20Files\Statement\200708NSGB\Old\NSGB_Credit_Statement_200708.zip
        {
            //>lblStatus.ForeColor = Color.Blue;
            //>lblStatus.Text = "Email Sent to Opration for " + strStatementType;
            //>txtRunResult.Text = lblStatus.Text + "\r\n\r\n" + basText.Left(txtRunResult.Text, 2000);
            BeginInvoke(setStatusDelegate, new object[] { bankCode, "Email Sent to Settlement for " + strStatementType });//strStatementType
        }
        else
        {
            //>lblStatus.ForeColor = Color.Red;
            //>lblStatus.Text = "Error on sending Email for " + strStatementType;
            //>txtRunResult.Text = lblStatus.Text + "\r\n\r\n" + basText.Left(txtRunResult.Text, 2000);
            BeginInvoke(setStatusDelegate, new object[] { bankCode, "Error on sending Settlement Email for " + strStatementType });//strStatementType
        }
        //Cursor.Current = Cursors.Default;
    }




    private void statusUpdate(int pBankCode, string pStat)
    {
        lblStatus.Text = pStat + " \r\n";//"Statement File Creation Done for " + pStat
        txtRunResult.Text = lblStatus.Text + basText.Left(txtRunResult.Text, 2000);// + "\r\n"  \r\n
                                                                                   //clsBasFile.SetAccessRule(clsBasFile.Filepath + @"StatementLog.txt");
        clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + pStat, System.IO.FileMode.Append);
    }

    private void btnUnSelectAll_Click(object sender, EventArgs e)
    {
        foreach (int indexChecked in chkLstProducts.CheckedIndices)
        {
            chkLstProducts.SetItemChecked(indexChecked, false);
        }
        chkRenameTables.Checked = false;
        chkTest.Checked = false;
        chkValidateBankEmail.Checked = false;
        chkSaveStatement.Checked = false;
        chkDevelopers.Checked = false;
        chkSendBankMail.Checked = false;
        chkLstProducts.Items.Clear();
    }

    private void OnCreateComplete(object sender, EventArgs e)
    {
    }

    private void btnUpdateDataCTL_Click(object sender, EventArgs e)
    {
        clsCheckNetwork.GetDNSAddressInfo("192.168.1.141");//mmohammed
    }



    private void makeAppTray()
    {
        this.components = new System.ComponentModel.Container();
        this.contextMenu1 = new System.Windows.Forms.ContextMenu();
        this.menuItem1 = new System.Windows.Forms.MenuItem();
        this.menuItem2 = new System.Windows.Forms.MenuItem();

        // Initialize contextMenu1
        this.contextMenu1.MenuItems.AddRange(
                    new System.Windows.Forms.MenuItem[] { this.menuItem1, this.menuItem2 });

        // Initialize menuItem1
        this.menuItem1.Index = 0;
        this.menuItem1.Text = "Show";
        this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);

        this.menuItem2.Index = 1;
        this.menuItem2.Text = "E&xit";
        this.menuItem2.Click += new System.EventHandler(this.menuItem2_Click);

        // Set up how the form should be displayed.
        //this.ClientSize = new System.Drawing.Size(441, 87);//292, 266

        // Create the NotifyIcon.
        this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);

        // The Icon property sets the icon that will appear
        // in the systray for this application.
        notifyIcon1.Icon = new Icon(this.Icon, 40, 40); //SystemIcons.Shield   //new Icon("FileMonitor.i
                                                        //);

        // The ContextMenu property sets the menu that will
        // appear when the systray icon is right clicked.
        notifyIcon1.ContextMenu = this.contextMenu1;

        // The Text property sets the text that will be displayed,
        // in a tooltip, when the mouse hovers over the systray icon.
        //notifyIcon1.Text = "Statement Program - " + Application.StartupPath;
        //notifyIcon1.Text = "Statement Program - " + Application.StartupPath + " " + Convert.ToString((int)(((float)progBarStatus.Value / (float)progBarStatus.Maximum) * 100.00)) + " %";
        inf = new AssemblyInfo();
        notifyIcon1.Text = "Statement Program - V:" + inf.Version;
        notifyIcon1.Visible = true;

        // Handle the DoubleClick event to activate the form.
        notifyIcon1.DoubleClick += new System.EventHandler(this.notifyIcon1_DoubleClick);

        this.Visible = false;
    }

    private void notifyIcon1_DoubleClick(object Sender, EventArgs e)
    {
        // Show the form when the user double clicks on the notify icon.

        // Set the WindowState to normal if the form is minimized.
        if (this.WindowState == FormWindowState.Minimized)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
        }
        else if (this.WindowState == FormWindowState.Normal)
        {
            this.Hide();
            this.WindowState = FormWindowState.Minimized;
        }

        // Activate the form.
        this.Activate();
    }

    private void menuItem1_Click(object Sender, EventArgs e)
    {
        // Close the form, which closes the application.
        this.Show();
        this.WindowState = FormWindowState.Normal;
        this.Activate();
    }

    private void menuItem2_Click(object Sender, EventArgs e)
    {
        // Close the form, which closes the application.
        if (MessageBox.Show("Are You Sure", "Attention", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly) == DialogResult.OK)
            this.Close();
    }

    private void frmStatementFile_Resize(object sender, EventArgs e)
    {
        if (this.WindowState == FormWindowState.Minimized)
            this.Hide();
    }

    private void renameTables()
    {
        //clsDbOracleLayer.doAction("rename tstatementmastertable to " + txtTblMaster.Text);
        //clsDbOracleLayer.doAction("rename tstatementdetailtable to " + txtTblDetail.Text);
        BeginInvoke(setStatusDelegate, new object[] { bankCode, "Start Rename Statement Tables" });//strStatementType
                                                                                                   //clsDbOracleLayer.doAction("create table  STMT." + txtTblMaster.Text + " as select * from " + MainSchema + "tstatementmastertable");
                                                                                                   //clsDbOracleLayer.doAction("create table  STMT." + txtTblDetail.Text + " as select * from " + MainSchema + "tstatementdetailtable");
                                                                                                   //// EMPA-8560
                                                                                                   //clsDbOracleLayer.doAction("ALTER TABLE STMT." + txtTblMaster.Text + " PARALLEL ( DEGREE 4 INSTANCES Default )");
                                                                                                   //clsDbOracleLayer.doAction("ALTER TABLE STMT." + txtTblDetail.Text + " PARALLEL ( DEGREE 4 INSTANCES Default )");
                                                                                                   //clsDbOracleLayer.doAction("CREATE UNIQUE INDEX STMT.MASTER" + txtTblMaster.Text.Substring(21, 9) + "_idx_1 ON STMT." + txtTblMaster.Text +
                                                                                                   //"(BRANCH, STATEMENTNO) " +
                                                                                                   //"LOGGING " +
                                                                                                   //"STORAGE    ( " +
                                                                                                   //"            BUFFER_POOL      DEFAULT " +
                                                                                                   //"            FLASH_CACHE      DEFAULT " +
                                                                                                   //"            CELL_FLASH_CACHE DEFAULT " +
                                                                                                   //"           ) " +
                                                                                                   //"NOPARALLEL " +
                                                                                                   //"ONLINE");
                                                                                                   //clsDbOracleLayer.doAction("CREATE INDEX STMT.MASTER" + txtTblMaster.Text.Substring(21, 9) + "_idx_2 ON STMT." + txtTblMaster.Text +
                                                                                                   //"(BRANCH) " +
                                                                                                   //"LOGGING " +
                                                                                                   //"STORAGE    ( " +
                                                                                                   //"            BUFFER_POOL      DEFAULT " +
                                                                                                   //"            FLASH_CACHE      DEFAULT " +
                                                                                                   //"            CELL_FLASH_CACHE DEFAULT " +
                                                                                                   //"           ) " +
                                                                                                   //"NOPARALLEL " +
                                                                                                   //"ONLINE");
                                                                                                   //clsDbOracleLayer.doAction("CREATE INDEX STMT.DETAIL" + txtTblDetail.Text.Substring(21, 9) + "_idx_1 ON STMT." + txtTblDetail.Text +
                                                                                                   //"(BRANCH, STATEMENTNO) " +
                                                                                                   //"LOGGING " +
                                                                                                   //"STORAGE    ( " +
                                                                                                   //"            BUFFER_POOL      DEFAULT " +
                                                                                                   //"            FLASH_CACHE      DEFAULT " +
                                                                                                   //"            CELL_FLASH_CACHE DEFAULT " +
                                                                                                   //"           ) " +
                                                                                                   //"NOPARALLEL " +
                                                                                                   //"ONLINE");
                                                                                                   //EDT-1224 => EMPA-443
                                                                                                   //clsDbOracleLayer.doActionProc("STMT.ZM_STMT_APP.RenameTable", txtTblMaster.Text.Substring(21, 9));
                                                                                                   //EDT-1456
        if (clsBasUserData.sType == "CR")
            clsDbOracleLayer.doActionRenameProc("STMT.ZM_STMT_APP.RenameTable", txtTblMaster.Text.Substring(16, 11), "CR");
        else if (clsBasUserData.sType == "CP")
            clsDbOracleLayer.doActionRenameProc("STMT.ZM_STMT_APP.RenameTable", txtTblMaster.Text.Substring(16, 11), "CP");
        else if (clsBasUserData.sType == "DB")
            clsDbOracleLayer.doActionRenameProc("STMT.ZM_STMT_APP.RenameTable", txtTblMaster.Text.Substring(16, 11), "DB");

        BeginInvoke(setStatusDelegate, new object[] { bankCode, "Rename Statement Tables Done" });//strStatementType
        if (!chkTest.Checked)
        {
            bool rtVal;
            //send the email
            string strMessage = string.Empty;
            clsEmail sndMail = new clsEmail();
            ArrayList pLstTo = new ArrayList(), pLstCC = new ArrayList(), pLstBCC = new ArrayList(), pLstAttachedFile = new ArrayList();
            pLstTo.Add("statement@emp-group.com");//"Developers@emp-group.com"
            strMessage = "Dear Team";
            strMessage += "\r\nRename Complete.. Please proceed.";
            rtVal = sndMail.sendEmailHTML("statement@emp-group.com", pLstTo, pLstCC, pLstBCC, pLstAttachedFile, "Rename tables - " + clsBasFile.TableName, strMessage, clsCnfg.readSetting("SmtpServer"), false);
            if (rtVal)
                this.BeginInvoke(this.setStatusDelegate, new object[] { 0, "\r\nSuccessfully Sent - Rename tables\r\n" });
            else
                this.BeginInvoke(this.setStatusDelegate, new object[] { 0, "\r\nError in Sending - Rename tables\r\n" });

            sndMail = null;
        }
    }
    private void chkStatGenAfterMonth_CheckedChanged(object sender, EventArgs e)
    {

    }

    private void btnExit_Click(object sender, System.EventArgs e)
    {
        if (MessageBox.Show("Are You Sure ?!", "Attention", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly) == DialogResult.OK)
            this.Close();
    }

    private void frmStatementFile_FormClosing(object sender, FormClosingEventArgs e)
    {
        if (e.CloseReason == CloseReason.TaskManagerClosing)
        {
            e.Cancel = true;
            this.notifyIcon1.Visible = false;
        }
    }

    private void txtTblMaster_TextChanged(object sender, EventArgs e)
    {
        txtTblMaster.CharacterCasing = CharacterCasing.Upper;
    }

    private void txtTblDetail_TextChanged(object sender, EventArgs e)
    {
        txtTblDetail.CharacterCasing = CharacterCasing.Upper;
    }

    private void txtDbSchema_TextChanged(object sender, EventArgs e)
    {
        txtDbSchema.CharacterCasing = CharacterCasing.Upper;
        if (txtDbSchema.Text.StartsWith("A4M"))
        {
            txtTablePrefix.Text = "";
            txtTablePrefix.ReadOnly = true;
            txtPrefixRun.Text = "";
            txtPrefixRun.ReadOnly = true;
        }
        else
        {
            txtPrefixRun.ReadOnly = false;
            GeneratePrefix();
        }
        if (clsBasUserData.sType == "CR")
            GenerateTableNameCR();
        else if (clsBasUserData.sType == "DB")
            GenerateTableNameDB();
        else if (clsBasUserData.sType == "CP")
            GenerateTableNameCP();
    }

    private void txtDbSchema_Leave(object sender, EventArgs e)
    {
        if (!txtDbSchema.Text.EndsWith("."))
            txtDbSchema.Text = txtDbSchema.Text + ".";
        if (clsBasUserData.sType == "CR")
            GenerateTableNameCR();
        else if (clsBasUserData.sType == "DB")
            GenerateTableNameDB();
        else if (clsBasUserData.sType == "CP")
            GenerateTableNameCP();
    }

    private void txtPrefixRun_Leave(object sender, EventArgs e)
    {
        //if (int.Parse(txtPrefixRun.Text) < 9)
        //    txtPrefixRun.Text = "0" + txtPrefixRun.Text;
        //else
        //    txtPrefixRun.Text = txtPrefixRun.Text;
        if (clsBasUserData.sType == "CR")
            GenerateTableNameCR();
        else if (clsBasUserData.sType == "DB")
            GenerateTableNameDB();
        else if (clsBasUserData.sType == "CP")
            GenerateTableNameCP();
    }

    //the cycle dropdown
    private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
    {
        comboBox2.Text = "";
        chkUpdateSummary.Visible = true;
        chkDevelopers.Visible = true;
        chkValidateBankEmail.Visible = false;
        if (clsBasUserData.sType == "CR")
        {
            if (comboBox1.SelectedIndex == 0) //5th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("13) BK Bank of Kigali >> Default Credit VIP Text 1/m", 328));//328
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire Classic >> Text 5/m", 49));//49
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire Classic >> PDF 5/m", 555));//49
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire Gold >> Text 5/m", 58));//58
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Standard >> Text 5/m", 138));//138
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Gold >> Text 5/m", 139));//139
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Installment >> Text 5/m", 173));//173
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Titanium >> Text 5/m", 273));//273
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard World Elite >> Text 5/m", 318));//318
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Platinum >> Text 5/m", 513));//513

                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire Classic STAFF >> Text 5/m", 413));//413
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire Gold STAFF >> Text 5/m", 414));//414
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Standard STAFF >> Text 5/m", 415));//415
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Gold STAFF >> Text 5/m", 416));//416
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Installment STAFF >> Text 5/m", 417));//417
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Titanium STAFF >> Text 5/m", 418));//418
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard World Elite STAFF >> Text 5/m", 419));//419
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Platinum STAFF >> Text 5/m", 514));//514
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Corporate >> Text 30/m", 515));//515
                chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Corporate STAFF >> Text 30/m", 516));//516

            }
            else if (comboBox1.SelectedIndex == 1) // 7th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 7th/m", 212));//212
                chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 7th/m", 267));//267
            }
            else if (comboBox1.SelectedIndex == 2) // 10th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("130) UNB Union National Bank >> Raw Text 10th/m", 431));//431
                chkLstProducts.Items.Add(new ListItem("130) UNB Union National Bank >> Credit Text 10th/m", 437));//437        
                       

            }
            else if (comboBox1.SelectedIndex == 3) // 12th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 12th/m", 213));//213
                chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 12th/m", 268));//268
            }
            else if (comboBox1.SelectedIndex == 4) // 15th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("4) BMSR Bank Misr Moga >> Text 16/m", 15));//15
                chkLstProducts.Items.Add(new ListItem("14) BIC Banco BIC, S.A. >> PDF 16/m", 9));//9
                chkLstProducts.Items.Add(new ListItem("21) ZEN Zenith Bank PLC >> MS Access MDB 16/m", 17));//17
                chkLstProducts.Items.Add(new ListItem("21) ZEN Zenith Bank PLC >> Default Text for Email 16/m", 66));//66
                chkLstProducts.Items.Add(new ListItem("22) BPC BANCO DE POUPANCA E CREDITO, SARL Credit >> Credit PDF 16/m", 14));//14
                chkLstProducts.Items.Add(new ListItem("29) UBA United Bank for Africa Plc Nigeria >> PDF 15/m", 436));//436
                chkLstProducts.Items.Add(new ListItem("31) NBS NBS Bank Limited >> PDF 16/m", 40));//40
                chkLstProducts.Items.Add(new ListItem("33) ZENG Zenith Bank (Ghana) Limited >> Visa Classic Credit PDF 16/m", 42));//42
                chkLstProducts.Items.Add(new ListItem("33) ZENG Zenith Bank (Ghana) Limited >> Visa Classic Credit (Staff) PDF 16/m", 281));//281
                chkLstProducts.Items.Add(new ListItem("33) ZENG Zenith Bank (Ghana) Limited >> MasterCard Credit PDF 16/m", 282));//282
                chkLstProducts.Items.Add(new ListItem("33) ZENG Zenith Bank (Ghana) Limited >> MasterCard Credit (Staff) PDF 16/m", 283));//283
                chkLstProducts.Items.Add(new ListItem("36) DBN Diamond Bank Nigeria >> Reward Default Text 16/m", 46));//46
                chkLstProducts.Items.Add(new ListItem("36) DBN Diamond Bank Nigeria - VIP >> Reward Default Text 16/m", 97));//97
                chkLstProducts.Items.Add(new ListItem("36) DBN Diamond Bank Nigeria >> VISA Platinum - ParkNShop Co-Brand Text 16/m", 111));//111
                chkLstProducts.Items.Add(new ListItem("36) DBN Diamond Bank Nigeria >> VISA Platinum - EXCO_VIP-Brand Text 16/m", 149));//149
                chkLstProducts.Items.Add(new ListItem("36) DBN Diamond Bank Nigeria >> MasterCard Credit Text 16/m", 220));//220
                chkLstProducts.Items.Add(new ListItem("37) KBL Keystone Bank >> Default Text 1/m", 57));//57
                chkLstProducts.Items.Add(new ListItem("38) GTBN GUARANTY TRUST BANK PLC NIGERIA  >> Credit Default PDF 16/m", 161));//161
                chkLstProducts.Items.Add(new ListItem("47) SBN Sterling Bank Nigeria >> Default Credit Text 15/m", 315));//315
                chkLstProducts.Items.Add(new ListItem("55) FBP Fidelity Bank PLC >> Default Text 16/m", 95));//95
                chkLstProducts.Items.Add(new ListItem("58) SBP Polaris BANK >> Default Text 16/m", 79));//79
                chkLstProducts.Items.Add(new ListItem("69) FBN First Bank of Nigeria  >> Credit Default Text 16/m", 123));//123
                chkLstProducts.Items.Add(new ListItem("77) I&M Bank Rwanda Limited >> Credit PDF 15/m", 201));//201
                chkLstProducts.Items.Add(new ListItem("85) UTBG UT Bank Limited Ghana >> Default Credit PDF 1/m", 167));//167
                chkLstProducts.Items.Add(new ListItem("87) WEMA BANK PLC NIGERIA >> Credit Default 15/m", 179));//179
                chkLstProducts.Items.Add(new ListItem("98) GTBK Guaranty trust bank Kenya  >> Credit PDF 15/m", 228));//228

                //chkLstProducts.Items.Add(new ListItem("106) DSBJ EAST AFRICA BANK Djibouti  >> Credit Islamic Text 15/m", 234));//234

                chkLstProducts.Items.Add(new ListItem("115) UNBN UNION BANK NIGERIA  >> Credit Text 15/m", 276));//276
                chkLstProducts.Items.Add(new ListItem("154) UBP Unity Bank PlC >> Credit Text 15/m", 441));//441        
            }
            else if (comboBox1.SelectedIndex == 5) // 17th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 17th/m", 214));//214
                chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 17th/m", 269));//269
            }
            else if (comboBox1.SelectedIndex == 6) // 20th
            {
                chkLstProducts.Items.Clear();
                //chkLstProducts.Items.Add(new ListItem("42) ENG Echo Bank Nigeria >> Default Credit Text 20/m", 206));//206
                chkLstProducts.Items.Add(new ListItem("51) TMB Trust Merchant Bank >> Raw Credit VISA Text 20/m", 75));//75
                chkLstProducts.Items.Add(new ListItem("51) TMB Trust Merchant Bank >> Raw Credit MasterCard Text 20/m", 204));//204
                chkLstProducts.Items.Add(new ListItem("51) TMB Trust Merchant Bank >> Credit PDF 20/m", 108));//108
                                                                                                              //chkLstProducts.Items.Add(new ListItem("51) TMB Trust Merchant Bank >> Corporate Text 20/m", 105));//105
                chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 20th/m", 133));//133
                chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 20th/m", 270));//270
                //chkLstProducts.Items.Add(new ListItem("115) UNBN UNION BANK NIGERIA  >> Credit Text 20/m", 277));//277
            }
            else if (comboBox1.SelectedIndex == 7) // 23rd
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 23rd/m", 215));//215
                chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 23rd/m", 271));//271
                chkLstProducts.Items.Add(new ListItem("158) RBGH Republic Bank (Ghana) Limited >> Default Credit Text 23rd/m", 459)); //459
            }
            else if (comboBox1.SelectedIndex == 8) // 27th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default Visa Credit Text 27th/m", 216));//216
                chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default MasterCard Credit Text 27th/m", 272));//272
                //chkLstProducts.Items.Add(new ListItem("115) UNBN UNION BANK NIGERIA  >> Credit Text 27/m", 278));//278
            }
            else if (comboBox1.SelectedIndex == 9) // EOM
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("1) QNB ALAHLI Credit >> Text Splitted by product 1/m", 1));//1

                //chkLstProducts.Items.Add(new ListItem("1) QNB ALAHLI Business >> Text 1/m", 2));//2       
                //chkLstProducts.Items.Add(new ListItem("1) QNB ALAHLI Business - SME >> Text 1/m", 78));//78   
                //chkLstProducts.Items.Add(new ListItem("1) QNB ALAHLI Mastercard Business - SME >> Text 1/m", 148));//148

                chkLstProducts.Items.Add(new ListItem("3) NBK VISA Classic >> Text 1/m", 3));//3
                chkLstProducts.Items.Add(new ListItem("3) NBK VISA Gold >> Text 1/m", 60));//60
                chkLstProducts.Items.Add(new ListItem("4) BMSR Bank Misr Credit Classic >> Text 1/m", 16));//16
                chkLstProducts.Items.Add(new ListItem("4) BMSR Bank Misr Credit Gold >> Text 1/m", 88));//88
                chkLstProducts.Items.Add(new ListItem("4) BMSR Bank Misr Credit Mobinet >> Text 1/m", 89));//89
                chkLstProducts.Items.Add(new ListItem("5) AIB Raw Data VISA USD>> Credit Text 1/m", 4));//4
                chkLstProducts.Items.Add(new ListItem("5) AIB Raw Data VISA EUR>> Credit Text 1/m", 64));//64
                chkLstProducts.Items.Add(new ListItem("5) AIB Raw Data VISA EGP>> Credit Text 1/m", 187));//187
                chkLstProducts.Items.Add(new ListItem("5) AIB Raw Data MasterCard USD >> Credit Text 1/m", 237));//237
                chkLstProducts.Items.Add(new ListItem("5) AIB PDF >> Credit PDF 1/m", 107));//107
                chkLstProducts.Items.Add(new ListItem("6) AAIB Arab African International Bank >> Raw Text 1/m", 205));//205
                chkLstProducts.Items.Add(new ListItem("6) AAIB Arab African International Bank >> Text Spool 1/m", 217));//217

                //chkLstProducts.Items.Add(new ListItem("7) BAI Products >> PDF 1/m", 7));//7

                chkLstProducts.Items.Add(new ListItem("7) BAI Visa Classic >> PDF 1/m", 464));//464
                chkLstProducts.Items.Add(new ListItem("7) BAI Visa Gold >> PDF 1/m", 465));//465
                chkLstProducts.Items.Add(new ListItem("7) BAI Visa Platinum >> PDF 1/m", 466));//466
                //chkLstProducts.Items.Add(new ListItem("7) BAI Corporate >> PDF 1/m", 140));//140
                //chkLstProducts.Items.Add(new ListItem("10) BNP Credit Reward >> Text 1/m", 37));//37
                //chkLstProducts.Items.Add(new ListItem("10) BNP Corporate >> Text 1/m", 36));//36

                chkLstProducts.Items.Add(new ListItem("11) GTB Bank Prestigo >> Credit Text 1/m", 400));//400
                                                                                                        //
                                                                                                        //chkLstProducts.Items.Add(new ListItem("11) GTB Bank Prestigo >> Credit MDB Access file 1/m", 430));//430        

                chkLstProducts.Items.Add(new ListItem("13) BK Bank of Kigali >> Default Credit Text 1/m", 154));//154
                chkLstProducts.Items.Add(new ListItem("16) ABP Access Bank Plc with Label Credit >> Text 1/m", 10));//10
                chkLstProducts.Items.Add(new ListItem("16) ABP Access Bank Plc with Label Dual Currency >> Text 1/m", 35));//35
                chkLstProducts.Items.Add(new ListItem("16) ABP Access Bank Plc with Label Infinite Dual Currency >> Text 1/m", 61));//61
                chkLstProducts.Items.Add(new ListItem("16) ABP Access Bank Plc with Label AMEX Dual Currency >> Text 1/m", 471));//471

                chkLstProducts.Items.Add(new ListItem("20) BICV Banco Intertlantico >> Credit Classic >> PDF 1/m", 23));//23
                chkLstProducts.Items.Add(new ListItem("20) BICV Banco Intertlantico >> Credit Classic >> Text 1/m", 165));//23

                //chkLstProducts.Items.Add(new ListItem("22) BPC BANCO DE POUPANCA E CREDITO, SARL Corporate >> Default PDF 1/m", 52));//52

                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) Credit Gold Standard>> PDF 1/m", 26));//26
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) Credit Gold Standard >> Text 1/m", 249));//249
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) Credit Classic Standard>> PDF 1/m", 27));//27
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) Credit Classic Standard>> Text 1/m", 260));//260
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) MC Credit Titanium Standard>> PDF 1/m", 196));//196
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) MC Credit Titanium Standard>> Text 1/m", 261));//261
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) MC Credit Classic Standard>> PDF 1/m", 197));//197
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) MC Credit Classic Standard>> Text 1/m", 262));//262
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) Credit Gold Staff>> PDF 1/m", 242));//242
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) Credit Gold Staff>> Text 1/m", 263));//263
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) Credit Classic Staff>> PDF 1/m", 243));//243
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) Credit Classic Staff>> Text 1/m", 264));//264
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) MC Credit Titanium Staff>> PDF 1/m", 244));//244
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) MC Credit Titanium Staff>> Text 1/m", 265));//265
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) MC Credit Classic Staff>> PDF 1/m", 245));//245
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) MC Credit Classic Staff>> Text 1/m", 266));//266

                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) MC Credit Platnium Standard>> PDF 1/m", 426));//426
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) MC Credit Platnium Standard>> Text 1/m", 427));//427
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) MC Credit Platnium Staff>> PDF 1/m", 428));//428
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) MC Credit Platnium Staff>> Text 1/m", 429));//429

                //chkLstProducts.Items.Add(new ListItem("24) BOAL Bank of Africa Mali >> Credit >> French PDF 1/m", 168));//168
                chkLstProducts.Items.Add(new ListItem("25) AUB Ahli United Bank >> Text 1/m", 51));// 51
                                                                                                   //chkLstProducts.Items.Add(new ListItem("26) BOAS Bank of Africa Senegal >> Credit >> French PDF 1/m", 199));//199
                chkLstProducts.Items.Add(new ListItem("27) BCA BANCO COMERCIAL DO ATLANTICO >> Credit >> PDF 1/m", 20));//20
                chkLstProducts.Items.Add(new ListItem("27) BCA BANCO COMERCIAL DO ATLANTICO >> Credit Text 30/m", 300));//300 iatta  
                                                                                                                        //chkLstProducts.Items.Add(new ListItem("27) BCA BANCO COMERCIAL DO ATLANTICO >> Corporate >> PDF 1/m", 70));//70
                chkLstProducts.Items.Add(new ListItem("28) CECV CAIXA ECONOMICA DE CABO VERDE >> Credit Classic >> PDF 1/m", 24));//24
                chkLstProducts.Items.Add(new ListItem("28) CECV CAIXA ECONOMICA DE CABO VERDE >> Credit Classic >> Text 1/m", 163));//24

                //chkLstProducts.Items.Add(new ListItem("29) UBA United Bank for Africa Plc Nigeria >> PDF 1/m", 8));//8
                //chkLstProducts.Items.Add(new ListItem("35) BOAC Bank of Africa Cote D'Ivoire >> Credit >> French PDF 1/m", 198));//198
                chkLstProducts.Items.Add(new ListItem("36) DBN Diamond Bank Nigeria - VIP 1,2,5 >> Reward Default Text 1/m", 83));//83


                //chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Corporate >> Text 30/m", 515));//515
                //chkLstProducts.Items.Add(new ListItem("41) BDCA Banque Du Caire MasterCard Corporate STAFF >> Text 30/m", 516));//516

                //chkLstProducts.Items.Add(new ListItem("47) SBN Sterling Bank Nigeria >> Default Credit Text 1/m", 222));//222

                chkLstProducts.Items.Add(new ListItem("50) ICBG ICBG Access Bank Ghana Limited Credit >> PDF 1/m", 99));//99
                                                                                                                        //chkLstProducts.Items.Add(new ListItem("50) ICBG Intercontinental Bank Ghana Limited Corporate >> PDF 1/m", 106));//106
                chkLstProducts.Items.Add(new ListItem("55) FBP Fidelity Bank PLC >> Default Text 1/m", 218));//218
                                                                                                             //chkLstProducts.Items.Add(new ListItem("56) NCB National Commercial Bank >> Credit >> PDF 1/m", 77));//77
                chkLstProducts.Items.Add(new ListItem("58) SBP Polaris BANK >> Default Credit Text 1/m", 92));//92
                //chkLstProducts.Items.Add(new ListItem("58) SBP SKYE BANK PLC >> Default Corporate Text 1/m", 125));//125
                //chkLstProducts.Items.Add(new ListItem("69) FBN First Bank of Nigeria  >> Corporate Default Text 1/m", 116));//116
                chkLstProducts.Items.Add(new ListItem("72) ABPR ACCESS BANK RWANDA SA >> Default Credit Text 1/m", 131));//131
                chkLstProducts.Items.Add(new ListItem("74) BLME Blom Bank Egypt VISA>> Credit Text 1/m", 143));//143
                chkLstProducts.Items.Add(new ListItem("74) BLME Blom Bank Egypt VISA>> Credit PDF 1/m", 135));//135
                chkLstProducts.Items.Add(new ListItem("74) BLME Blom Bank Egypt MasterCard>> Credit Text 1/m", 141));//141
                chkLstProducts.Items.Add(new ListItem("74) BLME Blom Bank Egypt MasterCard>> Credit PDF 1/m", 178));//178
                chkLstProducts.Items.Add(new ListItem("76) EDBE EXPORT DEVELOPMENT BANK OF EGYPT >> Default Credit Text 1/m", 146));//146


                chkLstProducts.Items.Add(new ListItem("77) I&M Bank Rwanda Limited >> Credit PDF 1/m", 201));//201

                chkLstProducts.Items.Add(new ListItem("81) SIBN STANBIC IBTC BANK NIGERIA  >> Credit Default Text 1/m", 156));//156
                chkLstProducts.Items.Add(new ListItem("94) FABG First Atlantic Bank Ghana >> Credit Text 1/m", 439));//439
                chkLstProducts.Items.Add(new ListItem("94) FABG First Atlantic Bank Ghana >> Credit PDF 1/m", 467));//467    
                chkLstProducts.Items.Add(new ListItem("109) CGBK COMPAGNIE GENERALE DE BANQUE LTD >> MasterCard Credit Text 1/m", 238));//238


                //chkLstProducts.Items.Add(new ListItem("114) IDBE Industrial Development & Workers Bank of Egypt >> Credit Text 1/m", 258));//258
                //chkLstProducts.Items.Add(new ListItem("114) IDBE Industrial Development & Workers Bank of Egypt >> Credit XML 1/m", 280));//280


                //chkLstProducts.Items.Add(new ListItem("115) UNBN UNION BANK NIGERIA  >> Credit Text 30/m", 279));//279

                chkLstProducts.Items.Add(new ListItem("122) ALXB ALEXBANK >> Credit Raw 30/m", 298));//298
                chkLstProducts.Items.Add(new ListItem("122) ALXB ALEXBANK >> Credit MF Raw 30/m", 344));//344
                chkLstProducts.Items.Add(new ListItem("122) ALXB ALEXBANK  >> Credit Text 30/m", 299));//299
                chkLstProducts.Items.Add(new ListItem("122) ALXB ALEXBANK >> Credit MF Text 30/m", 343));//343
                chkLstProducts.Items.Add(new ListItem("127) AIBK Arab Investment Bank of Egypt  >> Credit PDF 30/m", 292));//292
                chkLstProducts.Items.Add(new ListItem("127) AIBK Arab Investment Bank of Egypt  >> Credit Customer PDF 30/m", 293));//293
                chkLstProducts.Items.Add(new ListItem("127) AIBK Arab Investment Bank of Egypt  >> Credit Staff 30/m", 306));//306
                                                                                                                             //chkLstProducts.Items.Add(new ListItem("127) AIBK Arab Investment Bank of Egypt  >> Credit Raw 30/m", 412));//412
                                                                                                                             //chkLstProducts.Items.Add(new ListItem("127) AIBK Arab Investment Bank of Egypt  >> Installment PDF 30/m", 307));//307
                                                                                                                             //chkLstProducts.Items.Add(new ListItem("127) AIBK Arab Investment Bank of Egypt  >> Installment Customer PDF 30/m", 308));//308
                                                                                                                             //chkLstProducts.Items.Add(new ListItem("127) AIBK Arab Investment Bank of Egypt  >> Installment Staff PDF 30/m", 309));//309
                chkLstProducts.Items.Add(new ListItem("128) EGB Egyptian Gulf Bank of Egypt  >> Credit OMR PDF 30/m", 295));//295        
                                                                                                                            //chkLstProducts.Items.Add(new ListItem("128) EGB Egyptian Gulf Bank of Egypt >> Credit Raw 30/m", 412));//412
                chkLstProducts.Items.Add(new ListItem("136) BPG Bank Prestigo >> Credit Test 1/m", 319));//319        
                                                                                                         //                                                                                         //bsayed barka 210805
                chkLstProducts.Items.Add(new ListItem("153) BRKA AL Baraka Bank of Egypt >> Credit PDF 30/m", 470));//469
                chkLstProducts.Items.Add(new ListItem("153) BRKA AL Baraka Bank of Egypt >> Raw Text 30/m", 469));//469





            }
        }
        if (clsBasUserData.sType == "CP")
        {
            if (comboBox1.SelectedIndex == 0) //5th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("13) BK Bank of Kigali >> Default Corporate VIP Text 5/m", 329));//329
            }
            else if (comboBox1.SelectedIndex == 1) //15th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("38) GTBN GUARANTY TRUST BANK PLC NIGERIA  >> Corprate Default Text 16/m", 326));//326
                
            }
            else if (comboBox1.SelectedIndex == 2) //20th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("51) TMB Trust Merchant Bank >> Corporate Text 20/m", 105));//105
            }
            else if (comboBox1.SelectedIndex == 3) // EOM
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("1) QNB ALAHLI Business >> Text 1/m", 2));//2
                chkLstProducts.Items.Add(new ListItem("1) QNB ALAHLI Business - SME >> Text 1/m", 78));//78
                chkLstProducts.Items.Add(new ListItem("1) QNB ALAHLI Mastercard Business - SME >> Text 1/m", 148));//148
                chkLstProducts.Items.Add(new ListItem("1) QNB ALAHLI FEDCOC >> Text 1/m", 432));//432
                //chkLstProducts.Items.Add(new ListItem("1) QNB ALAHLI Cardholder Corporate - B2B >> Text 1/m", 433));//433
                chkLstProducts.Items.Add(new ListItem("1) QNB ALAHLI Corporate Contract - B2B >> Text 1/m", 434));//434
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Visa Platinum MVSE >> Text 1/m", 4000));//317
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI VISA Platinum B2B >> Text 1/m", 4003));//317
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Visa Platinum Single Limit in EGB >> Text 1/m", 4001));//317
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Visa Platinum SME >> Text 1/m", 4002));//317
                chkLstProducts.Items.Add(new ListItem("5) AIB Corporate >> Raw Data File 1/m", 296));//296
                chkLstProducts.Items.Add(new ListItem("5) AIB Corporate >> Default Text 1/m", 297));//297
                chkLstProducts.Items.Add(new ListItem("7) BAI Corporate >> PDF 1/m", 140));//140
                                                                                           //chkLstProducts.Items.Add(new ListItem("10) BNP Corporate >> Text 1/m", 36));//36
                chkLstProducts.Items.Add(new ListItem("13) BK Bank of Kigali >> Default Corporate Text 1/m", 330));//330
                chkLstProducts.Items.Add(new ListItem("16) ABP Access Bank Plc with Label Corporate >> Text 1/m", 43));//43
                chkLstProducts.Items.Add(new ListItem("22) BPC BANCO DE POUPANCA E CREDITO, SARL Corporate >> Default PDF 1/m", 52));//52
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) >> Corporate Text 1/m", 440));//440
                chkLstProducts.Items.Add(new ListItem("23) Suez Canal Bank(SCB) >> Corporate PDF 1/m", 7100));//440
                chkLstProducts.Items.Add(new ListItem("27) BCA BANCO COMERCIAL DO ATLANTICO >> Corporate >> PDF 1/m", 70));//70
                chkLstProducts.Items.Add(new ListItem("27) BCA BANCO COMERCIAL DO ATLANTICO >> Corporate >> Default Text 1/m", 341));//341
                //chkLstProducts.Items.Add(new ListItem("50) ICBG Access Bank Ghana Limited Corporate >> PDF 1/m", 106));//106
                chkLstProducts.Items.Add(new ListItem("58) SBP Polaris BANK >> Default Corporate Text 1/m", 125));//125
                chkLstProducts.Items.Add(new ListItem("69) FBN First Bank of Nigeria  >> Corporate Default Text 1/m", 116));//116
                chkLstProducts.Items.Add(new ListItem("81) SIBN STANBIC IBTC BANK NIGERIA  >> Corporate Default Text 1/m", 158));//158
                                                                                                                                 //chkLstProducts.Items.Add(new ListItem("112) GTBR Guaranty trust bank Rwanda  >> Corporate Text 1/m", 252));//252
                chkLstProducts.Items.Add(new ListItem("122) ALXB ALEXBANK >> Corporate Raw 30/m", 304));//304
                chkLstProducts.Items.Add(new ListItem("122) ALXB ALEXBANK  >> Corporate Text 30/m", 305));//305
                chkLstProducts.Items.Add(new ListItem("130) ADBC Corporate >> Text 10th/m", 1999));//437 

            }
        }
        else if (clsBasUserData.sType == "DB")
        {
            if (comboBox1.SelectedIndex == 0) // 15th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("33) ZENG Zenith Bank (Ghana) Limited >> MasterCard Prepaid text 16/m", 275));//275
                chkLstProducts.Items.Add(new ListItem("33) ZENG Zenith Bank (Ghana) Limited >> MasterCard Prepaid PDF 16/m", 442));//442
                chkLstProducts.Items.Add(new ListItem("55) FBP Fidelity Bank PLC >> Default Text Debit 16/m", 151));//151
                                                                                                                    //chkLstProducts.Items.Add(new ListItem("69) FBN First Bank of Nigeria  >> Prepaid Default Text 16/m", 115));//115
                chkLstProducts.Items.Add(new ListItem("77) I&M Bank Rwanda Limited >> Prepaid PDF 15/m", 202));//202
                chkLstProducts.Items.Add(new ListItem("87) WEMA BANK PLC NIGERIA >> Prepaid Default 15/m", 182));//182
                //chkLstProducts.Items.Add(new ListItem("98) GTBK Guaranty trust bank Kenya  >> Prepaid PDF 15/m", 256));//256
                //chkLstProducts.Items.Add(new ListItem("106) DSBJ EAST AFRICA BANK Djibouti  >> Debit Text 15/m", 233));//233
            }
            else if (comboBox1.SelectedIndex == 1) // 20th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("73) FCMB First City Monument Bank Nigeria >> Default Debit Text 20/m", 134));//134
            }
            else if (comboBox1.SelectedIndex == 2) // EOM
            {
                chkLstProducts.Items.Clear();
                //NSGB-15360
                //chkLstProducts.Items.Add(new ListItem("1) QNB ALAHLI MasterCard Salary Prepaid >> PDF 1/m", 114));//114

                chkLstProducts.Items.Add(new ListItem("5) AIB Electron EGP >> Debit Text 1/m", 188));//188
                chkLstProducts.Items.Add(new ListItem("5) AIB Prepaid EGP >> Prepaid Raw 1/m", 303));//303
                // removed by abdelsalam daif untile the bank confirmation
                //chkLstProducts.Items.Add(new ListItem("174) ABPK Prepaid >> Prepaid Raw 1/m", 1009));//1009
                chkLstProducts.Items.Add(new ListItem("7) BAI Prepaid >> PDF 1/m", 127));//127
                chkLstProducts.Items.Add(new ListItem("12) BSIC Alwaha Bank Libya Inc (Oasis Bank) >> PDF 1/m", 50));//50
                chkLstProducts.Items.Add(new ListItem("12) BSIC Alwaha Bank Libya Inc (Oasis Bank) >> Default Text 1/m", 339));//339
                chkLstProducts.Items.Add(new ListItem("13) BK Bank of Kigali >> Default Prepaid Text 1/m", 100));//100
                chkLstProducts.Items.Add(new ListItem("15) SSB SG-SSB Ltd Debit >> PDF 1/m", 74));//74
                chkLstProducts.Items.Add(new ListItem("20) BICV Banco Intertlantico >> Debit >> PDF 1/m", 25));//25
                chkLstProducts.Items.Add(new ListItem("20) BICV Banco Intertlantico >> Debit >> Text 1/m", 166));//25
                                                                                                                 //chkLstProducts.Items.Add(new ListItem("22) BPC BANCO DE POUPANCA E CREDITO, SARL Debit >> Debit PDF 1/m", 191));//191
                chkLstProducts.Items.Add(new ListItem("24) BOAL Bank of Africa Mali >> Debit >> Default PDF 1/m", 76));//76
                chkLstProducts.Items.Add(new ListItem("27) BCA BANCO COMERCIAL DO ATLANTICO >> Debit >> Default Text 1/m", 69));//69   
                chkLstProducts.Items.Add(new ListItem("28) CECV CAIXA ECONOMICA DE CABO VERDE >> Debit - Prepaid >> PDF 1/m", 91));//91
                chkLstProducts.Items.Add(new ListItem("28) CECV CAIXA ECONOMICA DE CABO VERDE >> Debit - Prepaid >> Text 1/m", 164));//91
                                                                                                                                     //chkLstProducts.Items.Add(new ListItem("32) UBAG UNITED BANK FOR AFRICA GHANA LIMITED >> Debit PDF 1/m", 193));//193
                                                                                                                                     //chkLstProducts.Items.Add(new ListItem("42) OBI Oceanic Bank International >> Default Debit Text 1/m", 119));//119
                                                                                                                                     //chkLstProducts.Items.Add(new ListItem("44) BOAK Bank of Africa Kenya >> Default Debit Text 1/m", 248));//248
                chkLstProducts.Items.Add(new ListItem("45) BOCD Bank Of Commerce and Development Debit >> PDF 1/m", 63));//63
                                                                                                                         //chkLstProducts.Items.Add(new ListItem("50) ICBG Intercontinental Bank Ghana Limited Debit >> Text 1/m", 87));//87
                                                                                                                         //chkLstProducts.Items.Add(new ListItem("50) ICBG Intercontinental Bank Ghana Limited Prepaid >> PDF 1/m", 113));//113
                                                                                                                         //chkLstProducts.Items.Add(new ListItem("56) NCB National Commercial Bank >> Prepaid >> PDF 1/m", 226));//226
                chkLstProducts.Items.Add(new ListItem("58) SBP Polaris BANK >> Default Prepaid Text 1/m", 175));//175
                                                                                                                //chkLstProducts.Items.Add(new ListItem("60) GUM GUMHOURIA Bank  >> Prepaid PDF 1/m", 207));//207
                                                                                                                //chkLstProducts.Items.Add(new ListItem("62) WHDA WHDA Bank  >> Prepaid PDF 1/m", 200));//200
                chkLstProducts.Items.Add(new ListItem("65) SBPG SKYE BANK PLC Gambia >> Default Debit Text 1/m", 121));//121
                chkLstProducts.Items.Add(new ListItem("66) ABPG Access Bank Plc Gambia >> Default Debit Text 1/m", 129));//129
                                                                                                                         //chkLstProducts.Items.Add(new ListItem("71) BOAD BANK OF AFRICA DEMOCRATIC REPUBLIC OF CONGO >> Default Debit Text 1/m", 136));//136
                                                                                                                         //chkLstProducts.Items.Add(new ListItem("72) ABPR ACCESS BANK RWANDA SA >> Default Debit Text 1/m", 132));//132
                chkLstProducts.Items.Add(new ListItem("75) UMB UNIVERSAL MERCHANT BANK GHANA >> Default Prepaid Text 1/m", 137));//137
                                                                                                                                 //chkLstProducts.Items.Add(new ListItem("90) UBG UniBank Ghana Limited  >> Debit PDF 1/m", 195));//195
                                                                                                                                 //chkLstProducts.Items.Add(new ListItem("95) GTBL Guaranty trust bank Liberia  >> Debit Text 1/m", 229));//229
                chkLstProducts.Items.Add(new ListItem("98) GTBK Guaranty trust bank Kenya  >> Prepaid PDF 15/m", 256));//256
                                                                                                                       //chkLstProducts.Items.Add(new ListItem("104) DSBS Dahabshil Bank International Somalia  >> Debit Text 1/m", 232));//232
                                                                                                                       //chkLstProducts.Items.Add(new ListItem("109) CGBK COMPAGNIE GENERALE DE BANQUE LTD >> MasterCard Prepaid Text 1/m", 239));//239
                                                                                                                       //chkLstProducts.Items.Add(new ListItem("110) HBLN Heritage Banking Company Nigeria >> Prepaid Text 1/m", 246));//246
                                                                                                                       //chkLstProducts.Items.Add(new ListItem("112) GTBR Guaranty trust bank Rwanda  >> Debit Text 1/m", 251));//251
                chkLstProducts.Items.Add(new ListItem("113) GTBU Guaranty trust bank Uganda  >> Debit Text 1/m", 254));//254
                chkLstProducts.Items.Add(new ListItem("127) AIBK Arab Investment Bank of Egypt  >> Prepaid Customer PDF 30/m", 310));//310
                chkLstProducts.Items.Add(new ListItem("127) AIBK Arab Investment Bank of Egypt  >> Prepaid Staff PDF 30/m", 311));//311
                chkLstProducts.Items.Add(new ListItem("129) SBL SAHARA BANK BNP PARIBAS  >> Prepaid Text 1/m", 294));//294

                chkLstProducts.Items.Add(new ListItem("146) Fidelity Bank Ghana Limited  >> PrePaid text 1/m", 502)); //FBPG
                chkLstProducts.Items.Add(new ListItem("[146] Fidelity Bank Ghana Limited  >> Credit text 1/m", 503)); //FBPG
            }
        }
        //else if (comboBox1.SelectedIndex == 9) // All
        //    {
        //    chkLstProducts.Items.Clear();
        //    b4FrmLoad();
        //    }
    }

    //the e-stmt dropdown
    private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
    {
        comboBox1.Text = "";
        chkUpdateSummary.Visible = false;
        chkDevelopers.Visible = false;
        chkValidateBankEmail.Visible = true;
        if (clsBasUserData.sType == "CR")
        {
            if (comboBox2.SelectedIndex == 0) //5th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("[13] BK Bank of Kigali >> VISA RWF Credit VIP Emails 5/m", 331));//331
                chkLstProducts.Items.Add(new ListItem("[13] BK Bank of Kigali >> VISA USD Credit VIP Emails 5/m", 332));//332
                chkLstProducts.Items.Add(new ListItem("[13] BK Bank of Kigali >> MasterCard RWF Credit VIP Emails 5/m", 333));//333
                chkLstProducts.Items.Add(new ListItem("[13] BK Bank of Kigali >> MasterCard USD Credit VIP Emails 5/m", 334));//334

                //// BDCA New E-Statement
                chkLstProducts.Items.Add(new ListItem("[41] BDCA Banque Du Caire Classic  >> Credit email 5/m", 405));//405
                chkLstProducts.Items.Add(new ListItem("[41] BDCA Banque Du Caire Gold  >> Credit email 5/m", 406));//406
                chkLstProducts.Items.Add(new ListItem("[41] BDCA Banque Du Caire MasterCard Standard  >> Credit email 5/m", 407));//407
                chkLstProducts.Items.Add(new ListItem("[41] BDCA Banque Du Caire MasterCard Gold  >> Credit email 5/m", 408));//408
                chkLstProducts.Items.Add(new ListItem("[41] BDCA Banque Du Caire MasterCard Installment  >> Credit email 5/m", 409));//409
                chkLstProducts.Items.Add(new ListItem("[41] BDCA Banque Du Caire MasterCard Titanium  >> Credit email 5/m", 410));//410
                chkLstProducts.Items.Add(new ListItem("[41] BDCA Banque Du Caire MasterCard World Elite  >> Credit email 5/m", 411));//411
                chkLstProducts.Items.Add(new ListItem("[41] BDCA Banque Du Caire MasterCard Platinum  >> Credit email 5/m", 511));//511
                chkLstProducts.Items.Add(new ListItem("[41] BDCA Banque Du Caire MasterCard Corporate  >> Credit email 1/m", 512));//512


            }
            else if (comboBox2.SelectedIndex == 1) //7th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("[73] FCMB First City Monument Bank Nigeria >> Visa Credit Emails 7/m", 447));//447
                chkLstProducts.Items.Add(new ListItem("[73] FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 7/m", 448));//448
            }
            else if (comboBox2.SelectedIndex == 2) //12th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("[73] FCMB First City Monument Bank Nigeria >> Visa Credit Emails 12/m", 449));//449
                chkLstProducts.Items.Add(new ListItem("[73] FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 12/m", 450));//450
            }
            else if (comboBox2.SelectedIndex == 3) //15th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("[21] ZEN Zenith Bank PLC >> Email 16/m", 41));//41
                chkLstProducts.Items.Add(new ListItem("[22] BPC BANCO DE POUPANCA E CREDITO, SARL Credit >> Email 16/m", 144));//144
                chkLstProducts.Items.Add(new ListItem("[29] UBA United Bank for Africa Plc Nigeria >> Emails 15/m", 435));//435
                chkLstProducts.Items.Add(new ListItem("[31] NBS NBS Bank Limited >> Email 16/m", 145));//145

                ////chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> Email 16/m", 48));//48
                //chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> Email Classic 16/m", 284));
                //chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> Email Gold 16/m", 285));
                //chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> Email Platinum 16/m", 286));
                //chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> Email Platinum-USD 16/m", 340));
                //chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria - VIP >> Email 16/m", 98));//98
                //chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> VISA Platinum - ParkNShop Co-Brand Email 16/m", 112));//112
                //chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> VISA SPECIAL EXCO VIP - Privilege Email 16/m", 150));//150 
                //chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> MasterCard Credit Email 16/m", 221));//221

                chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> Email Classic 16/m", 284));
                chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> Email Gold 16/m", 285));
                chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> Email Platinum 16/m", 286));
                chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> Email Platinum-USD 16/m", 340));
                chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria - VIP >> Email 16/m", 98));//98
                chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> VISA Platinum - ParkNShop Co-Brand Email 16/m", 112));//112
                chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> VISA SPECIAL EXCO VIP - Privilege Email 16/m", 150));//150 
                chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria >> MasterCard Credit Email 16/m", 221));//221

                chkLstProducts.Items.Add(new ListItem("[37] KBL Keystone Bank >> Emails 1/m", 62));//62
                chkLstProducts.Items.Add(new ListItem("[38] GTBN GUARANTY TRUST BANK PLC NIGERIA  >> Credit Emails 16/m", 162));//162
                chkLstProducts.Items.Add(new ListItem("[47] SBN Sterling Bank Nigeria >> Classic, Gold, and platinum Credit Email 15/m", 223));//223
                chkLstProducts.Items.Add(new ListItem("[47] SBN Sterling Bank Nigeria >> INFINITE Credit Email 15/m", 316));//316
                chkLstProducts.Items.Add(new ListItem("[47] SBN Sterling Bank Nigeria >> SIGNATURE Credit Email 15/m", 401));//401
                //chkLstProducts.Items.Add(new ListItem("[47] SBN Sterling Bank Nigeria >> Default Credit Email 15/m", 223));//223
                //chkLstProducts.Items.Add(new ListItem("[47] SBN Sterling Bank Nigeria >> Default Credit Email 15/m", 316));//316

                chkLstProducts.Items.Add(new ListItem("[55] FBP Fidelity Bank PLC >> Emails 16/m", 96));//96
                chkLstProducts.Items.Add(new ListItem("[58] SBP Polaris BANK >> Emails 16/m", 80));//80
                chkLstProducts.Items.Add(new ListItem("[58] SBP Polaris BANK  >> MasterCard Platinum Credit Emails 16/m", 253));//253
                                                                                                                               //chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Credit Emails 16/m", 124));//124
                chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Emails Classic 16/m", 321));//321
                chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Emails Gold 16/m", 322));//322
                chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Emails Infinite 16/m", 323));//323
                chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Emails Classic Platinum 16/m", 324));//324
                chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Supplementary Credit Emails 16/m", 152));//152
                chkLstProducts.Items.Add(new ListItem("[77] I&M Bank Rwanda Limited >> Credit Emails 15/m", 320));//320
                chkLstProducts.Items.Add(new ListItem("[87] WEMA BANK PLC NIGERIA >> Credit Emails 15/m", 180));//180
                chkLstProducts.Items.Add(new ListItem("[98] GTBK Guaranty trust bank Kenya  >> Credit Emails 15/m", 230));//230
                chkLstProducts.Items.Add(new ListItem("[115] UNBN UNION BANK NIGERIA  >> Credit Emails 15/m", 288));//288

            }
            else if (comboBox2.SelectedIndex == 4) //17th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("[73] FCMB First City Monument Bank Nigeria >> Visa Credit Emails 17/m", 451));//451
                chkLstProducts.Items.Add(new ListItem("[73] FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 17/m", 452));//452
            }
            else if (comboBox2.SelectedIndex == 5) //20th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("[73] FCMB First City Monument Bank Nigeria >> Visa Credit Emails 20/m", 453));//453
                chkLstProducts.Items.Add(new ListItem("[73] FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 20/m", 454));//454
                //chkLstProducts.Items.Add(new ListItem("[115] UNBN UNION BANK NIGERIA  >> Credit Emails 20/m", 289));//289
            }
            else if (comboBox2.SelectedIndex == 6) //23th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("[73] FCMB First City Monument Bank Nigeria >> Visa Credit Emails 23/m", 455));//455
                chkLstProducts.Items.Add(new ListItem("[73] FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 23/m", 456));//456
                chkLstProducts.Items.Add(new ListItem("[158] RBGH Republic Bank (Ghana) Limited >> Default Credit Emails 23rd/m", 468)); //468
            }
            else if (comboBox2.SelectedIndex == 7) //27th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("[73] FCMB First City Monument Bank Nigeria >> Visa Credit Emails 27/m", 457));//457
                chkLstProducts.Items.Add(new ListItem("[73] FCMB First City Monument Bank Nigeria >> MasterCard Credit Emails 27/m", 458));//458
                //chkLstProducts.Items.Add(new ListItem("[115] UNBN UNION BANK NIGERIA  >> Credit Emails 27/m", 290));//290
            }
            else if (comboBox2.SelectedIndex == 8) // EOM
            {
                chkLstProducts.Items.Clear();
                //chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> Classic Email 1/m", 169));//169
                //chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> Gold Email 1/m", 170));//170
                //chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> Platinum Email 1/m", 171));//171
                //chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> Infinite Email 1/m", 172));//172
                //chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> Business Individual Email 1/m", 56));//56
                //chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> MasterCard Standard Email 1/m", 174));//174
                //chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> MasterCard Titanium Email 1/m", 403));//403

                //Orabi

                //chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> Platinum PDF Email  1/m", 404));//404
                //chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> Classic PDF Email  1/m", 420));//420
                //chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> Gold PDF Email  1/m", 421));//421
                //chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> Infinite PDF Email  1/m", 422));//422
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> Business Individual PDF Email  1/m", 423));//423
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> MasterCard Standard PDF Email  1/m", 424));//424
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> MasterCard Titanium PDF Email  1/m", 425));//425
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> MasterCard World Elite PDF Email  1/m", 446));//446
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> VISA Signature PDF Email 1/m", 472));//472
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Credit >> bebasata Gold Credit PDF Email 1/m", 473));//473


                //chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Mastercard Business - SME >> Email 1/m", 236));//236
                //chkLstProducts.Items.Add(new ListItem("[7] BAI Products >> Email 1/m", 85));//85
                chkLstProducts.Items.Add(new ListItem("[7] BAI Visa Classic >> Email 1/m", 460));//460
                chkLstProducts.Items.Add(new ListItem("[7] BAI Visa Gold >> Email 1/m", 461));//461
                chkLstProducts.Items.Add(new ListItem("[7] BAI Visa Platinum >> Email 1/m", 462));//462
                                                                                                  //chkLstProducts.Items.Add(new ListItem("[7] BAI Corporate >> Email 1/m", 181));//181
                chkLstProducts.Items.Add(new ListItem("[13] BK Bank of Kigali >> VISA RWF Credit Emails 1/m", 155));//155
                chkLstProducts.Items.Add(new ListItem("[13] BK Bank of Kigali >> VISA USD Credit Emails 1/m", 287));//287
                chkLstProducts.Items.Add(new ListItem("[13] BK Bank of Kigali >> MasterCard RWF Credit Emails 1/m", 313));//313
                chkLstProducts.Items.Add(new ListItem("[13] BK Bank of Kigali >> MasterCard USD Credit Emails 1/m", 314));//314
                chkLstProducts.Items.Add(new ListItem("[16] ABP Access Bank Plc  >> Credit Emails 1/m", 211));//211

                //chkLstProducts.Items.Add(new ListItem("[16] ABP Access Bank Plc  >> Corporate Cardholder Emails 1/m", 225));//225
                //chkLstProducts.Items.Add(new ListItem("[22] BPC BANCO DE POUPANCA E CREDITO, SARL Credit >> Cardholder Email 16/m", 186));//186

                //lajin
                //chkLstProducts.Items.Add(new ListItem("[29] UBA United Bank for Africa Plc Nigeria >> Emails 1/m", 45));//45

                chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria - VIP 1,2,5 >> Email 1/m", 84));//84

                //chkLstProducts.Items.Add(new ListItem("[36] DBN Diamond Bank Nigeria - VIP 1,2,5 Supplementary>> Email 1/m", 102));//102

                

                //chkLstProducts.Items.Add(new ListItem("[36] ABP (DBN Diamond Bank Nigeria) - VIP 1,2,5 >> Email 1/m", 84));//84
                //chkLstProducts.Items.Add(new ListItem("[36] ABP (DBN Diamond Bank Nigeria) - VIP 1,2,5 Supplementary>> Email 1/m", 102));//102

                //chkLstProducts.Items.Add(new ListItem("[47] SBN Sterling Bank Nigeria >> Default Credit Email 1/m", 223));//223
                chkLstProducts.Items.Add(new ListItem("[50] ICBG Access Bank Ghana Limited Credit >> Email 1/m", 274));//274
                chkLstProducts.Items.Add(new ListItem("[55] FBP Fidelity Bank PLC >> Emails 1/m", 219));//219
                                                                                                        //chkLstProducts.Items.Add(new ListItem("[56] NCB National Commercial Bank >> Credit >> Email 1/m", 224));//224
                chkLstProducts.Items.Add(new ListItem("[58] SBP Polaris BANK >> Emails 1/m", 93));//93
                                                                                                  //chkLstProducts.Items.Add(new ListItem("[58] SBP SKYE BANK PLC >> Corporate Emails 1/m", 126));//126
                                                                                                  //chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Corporate Cardholder Emails 1/m", 122));//122

                //lajin
                //chkLstProducts.Items.Add(new ListItem("[77] I&M Bank Rwanda Limited >> Credit Emails 1/m", 320));//320


                chkLstProducts.Items.Add(new ListItem("[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Platinum Emails 1/m", 208));//208
                chkLstProducts.Items.Add(new ListItem("[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Gold Emails 1/m", 209));//209
                chkLstProducts.Items.Add(new ListItem("[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Infinite Emails 1/m", 210));//210
                chkLstProducts.Items.Add(new ListItem("[81] SIBN STANBIC IBTC BANK NIGERIA  >> Credit Silver Emails 1/m", 259));//259

                //chkLstProducts.Items.Add(new ListItem("[81] STANBIC IBTC BANK NIGERIA >> Corporate Cardholder Emails 1/m", 159));//159
                //chkLstProducts.Items.Add(new ListItem("[110] HBLN Heritage Banking Company Nigeria >> MasterCard Credit Emails 1/m", 241));//241
                //chkLstProducts.Items.Add(new ListItem("[115] UNBN UNION BANK NIGERIA  >> Credit Emails 30/m", 291));//291

                chkLstProducts.Items.Add(new ListItem("[122] ALXB ALEXBANK  >> Credit email 30/m", 301));//301
                chkLstProducts.Items.Add(new ListItem("[122] ALXB ALEXBANK  >> Credit MF email 30/m", 342));//342
                chkLstProducts.Items.Add(new ListItem("[127] AIBK Arab Investment Bank of Egypt  >> Credit email 30/m", 302));//302
                chkLstProducts.Items.Add(new ListItem("[127] AIBK Arab Investment Bank of Egypt  >> Valu Credit email 30/m", 3097));//3097
                chkLstProducts.Items.Add(new ListItem("[146] Fidelity Bank Ghana Limited  >> Credit email 1/m", 501)); //FBPG

                //chkLstProducts.Items.Add(new ListItem("[139] CMB Coronation Merchant Bank  >> Credit email 30/15/m", 402));//402

                //chkLstProducts.Items.Add(new ListItem("[127] AIBK Arab Investment Bank of Egypt  >> Installment email 30/m", 312));//312


            }
        }
        if (clsBasUserData.sType == "CP")
        {
            if (comboBox2.SelectedIndex == 0) // 5th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("[13] BK Bank of Kigali >> MasterCard Corporate Cardholder VIP Credit Emails 5/m", 335));//335
                chkLstProducts.Items.Add(new ListItem("[13] BK Bank of Kigali >> MasterCard Corporate Cardholder VIP Credit Emails 5/m", 336));//336
            }
            else if (comboBox2.SelectedIndex == 1) // 15th
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("[38] GTBN GUARANTY TRUST BANK PLC NIGERIA  >> Corporate Cardholder Emails 16/m", 327));//327
            }
            else if (comboBox2.SelectedIndex == 2) // EOM
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI VISA Business Corporate >> Email 1/m", 317));//317
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Mastercard Business - SME >> Email 1/m", 236));//236
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Visa Business B2B >> Email 1/m", 443));//317
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Visa Business FEDCOC >> Email 1/m", 444));//236
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Visa Platinum MVSE >> Email 1/m", 3000));//317
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI VISA Platinum B2B >> Email 1/m", 3001));//317
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Visa Platinum Single Limit in EGB >> Email 1/m", 3002));//317
                chkLstProducts.Items.Add(new ListItem("[1] QNB ALAHLI Visa Platinum SME >> Email 1/m", 3003));//317

                //chkLstProducts.Items.Add(new ListItem("[7] BAI Corporate >> Email 1/m", 181));//181
                chkLstProducts.Items.Add(new ListItem("[7] BAI Visa Business >> Email 1/m", 463));//463
                chkLstProducts.Items.Add(new ListItem("[13] BK Bank of Kigali >> MasterCard RWF Corporate Cardholder Emails 1/m", 337));//337
                chkLstProducts.Items.Add(new ListItem("[13] BK Bank of Kigali >> MasterCard USD Corporate Cardholder Emails 1/m", 338));//338
                //chkLstProducts.Items.Add(new ListItem("[16] ABP Access Bank Plc  >> Corporate Cardholder Emails 1/m", 225));//225
                chkLstProducts.Items.Add(new ListItem("[22] BPC BANCO DE POUPANCA E CREDITO, SARL Credit >> Cardholder Email 1/m", 186));//186
                chkLstProducts.Items.Add(new ListItem("[58] SBP Polaris BANK >> Corporate Emails 1/m", 126));//126
                chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Corporate Cardholder Emails 1/m", 122));//122
                chkLstProducts.Items.Add(new ListItem("[81] STANBIC IBTC BANK NIGERIA >> Corporate Cardholder Emails 1/m", 159));//159
                chkLstProducts.Items.Add(new ListItem("[122] ALXB ALEXBANK  >> Corporate email 30/m", 438));//438

            }
        }
        else if (clsBasUserData.sType == "DB")
        {
            if (comboBox2.SelectedIndex == 0) //15th
            {
                chkLstProducts.Items.Clear();
                //chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> Prepaid Emails 16/m", 117));//117
                //chkLstProducts.Items.Add(new ListItem("[69] FBN First Bank of Nigeria  >> MasterCard Prepaid Emails 16/m", 203));//203
                chkLstProducts.Items.Add(new ListItem("[77] I&M Bank Rwanda Limited >> Prepaid Emails 15/m", 325));//325
                chkLstProducts.Items.Add(new ListItem("[87] WEMA BANK PLC NIGERIA >> Prepaid Emails 15/m", 183));//183
                                                                                                                 //chkLstProducts.Items.Add(new ListItem("[98] GTBK Guaranty trust bank Kenya  >> Prepaid Emails 15/m", 257));//257
            }
            else if (comboBox2.SelectedIndex == 1) // EOM
            {
                chkLstProducts.Items.Clear();
                chkLstProducts.Items.Add(new ListItem("[7] BAI Prepaid >> Email 1/m", 128));//128
                chkLstProducts.Items.Add(new ListItem("[13] BK Bank fo Kigali >> Prepaid Emails 1/m", 101));//101
                                                                                                            //chkLstProducts.Items.Add(new ListItem("[22] BPC BANCO DE POUPANCA E CREDITO, SARL Debit >> Email 1/m", 192));//192
                                                                                                            //chkLstProducts.Items.Add(new ListItem("[32] UBAG UNITED BANK FOR AFRICA GHANA LIMITED >> Debit Email 1/m", 194));//194
                                                                                                            //chkLstProducts.Items.Add(new ListItem("[56] NCB National Commercial Bank >> Prepaid >> Email 1/m", 227));//227
                chkLstProducts.Items.Add(new ListItem("[58] SBP Polaris BANK >> Prepaid VISA Emails 1/m", 176));//176
                chkLstProducts.Items.Add(new ListItem("[58] SBP Polaris BANK >> Prepaid VISA-NTDC Emails 1/m", 184));//184
                chkLstProducts.Items.Add(new ListItem("[58] SBP Polaris BANK >> Prepaid MasterCard Emails 1/m", 177));//177
                chkLstProducts.Items.Add(new ListItem("[98] GTBK Guaranty trust bank Kenya  >> Prepaid Emails 15/m", 257));//257
                                                                                                                           //chkLstProducts.Items.Add(new ListItem("[110] HBLN Heritage Banking Company Nigeria >> MasterCard Prepaid Emails 1/m", 247));//247
                                                                                                                           //chkLstProducts.Items.Add(new ListItem("[113] GTBU Guaranty trust bank Uganada >> Debit Emails 1/m", 255));//255

                chkLstProducts.Items.Add(new ListItem("[146] Fidelity Bank Ghana Limited  >> PrePaid email 1/m", 500));//FBP
                //chkLstProducts.Items.Add(new ListItem("[146] Fidelity Bank Ghana Limited  >> Credit email 1/m", 501)); //FBPG
            }
        }
    }

    private void datStmntData_ValueChanged(object sender, EventArgs e)
    {
        GeneratePrefix();
        StDate = datStmntData.Value;
        if (txtDbSchema.Text.StartsWith("A4M"))
        {
            txtTablePrefix.Text = "";
            txtPrefixRun.Text = "";
        }
        if (chkRenameTables.Checked == true)
        {
            //txtTablePrefix.ReadOnly = true;
            txtPrefixRun.ReadOnly = true;
            //if (!chkTest.Checked)
            //    {
            if (clsBasUserData.sType == "CR")
                GenerateRunCR();
            else if (clsBasUserData.sType == "DB")
                GenerateRunDB();
            else if (clsBasUserData.sType == "CP")
                GenerateRunCP();
            //txtPrefixRun.ReadOnly = false;
            //}
        }
        else
        {
            txtPrefixRun.Text = "";
            txtPrefixRun.ReadOnly = false;
        }
        if (clsBasUserData.sType == "CR")
            GenerateTableNameCR();
        else if (clsBasUserData.sType == "DB")
            GenerateTableNameDB();
        else if (clsBasUserData.sType == "CP")
            GenerateTableNameCP();
    }

    private void GeneratePrefix()
    {
        txtTablePrefix.Text = datStmntData.Value.ToString("yyMMdd");
    }

    private void chkTest_CheckedChanged(object sender, EventArgs e)
    {
        //if (chkTest.Checked == true)
        //    {
        //    txtTablePrefix.ReadOnly = false;
        //    txtTablePrefix.Text = "";
        //    txtPrefixRun.Text = "";
        //    }
        //else
        //    {
        //    txtTablePrefix.ReadOnly = true;
        //    GeneratePrefix();
        //    }
        //GenerateTableName();
        if (chkTest.Checked == true)
        {
            button1.Show();
        }
        else
        {
            Internal = false;
            InternalMailCC.Clear();
            InternalMailTo.Clear();
            InternalMailBCC.Clear();
            InternalMailFrom = "";
            InternalMailFromName = "";

            button1.BackColor = Color.Red;
            button1.Text = "External Mode";
            button1.Hide();
        }
    }

    private void GenerateRunCR()
    {
        txtPrefixRun.Text = "";
        int count = 0;
        OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
        OracleCommand cmd = new OracleCommand("select count(table_name) count from all_tables where table_name like '%MASTERCR%" + txtTablePrefix.Text + "%' and owner = '" + txtDbSchema.Text.Replace(".", "") + "'", conn);
        try
        {
            conn.Open();
            OracleDataReader dr;
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    count = int.Parse(dr.GetValue(0).ToString());
                }
            }
            dr.Close();
            conn.Close();
            if (count < 9)
            {
                txtPrefixRun.Text += "0" + (count + 1).ToString();
            }
            else
            {
                txtPrefixRun.Text = (count + 1).ToString();
            }
            clsBasFile.TableName = "CR_" + txtTablePrefix.Text + txtPrefixRun.Text;
        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
    }

    private void GenerateRunDB()
    {
        txtPrefixRun.Text = "";
        int count = 0;
        OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
        OracleCommand cmd = new OracleCommand("select count(table_name) count from all_tables where table_name like '%MASTERDB%" + txtTablePrefix.Text + "%' and owner = '" + txtDbSchema.Text.Replace(".", "") + "'", conn);
        try
        {
            conn.Open();
            OracleDataReader dr;
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    count = int.Parse(dr.GetValue(0).ToString());
                }
            }
            dr.Close();
            conn.Close();
            if (count < 9)
            {
                txtPrefixRun.Text += "0" + (count + 1).ToString();
            }
            else
            {
                txtPrefixRun.Text = (count + 1).ToString();
            }
            clsBasFile.TableName = "DB_" + txtTablePrefix.Text + txtPrefixRun.Text;
        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
    }

    private void GenerateRunCP()
    {
        txtPrefixRun.Text = "";
        int count = 0;
        OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
        OracleCommand cmd = new OracleCommand("select count(table_name) count from all_tables where table_name like '%MASTERCP%" + txtTablePrefix.Text + "%' and owner = '" + txtDbSchema.Text.Replace(".", "") + "'", conn);
        try
        {
            conn.Open();
            OracleDataReader dr;
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    count = int.Parse(dr.GetValue(0).ToString());
                }
            }
            dr.Close();
            conn.Close();
            if (count < 9)
            {
                txtPrefixRun.Text += "0" + (count + 1).ToString();
            }
            else
            {
                txtPrefixRun.Text = (count + 1).ToString();
            }
            clsBasFile.TableName = "CP_" + txtTablePrefix.Text + txtPrefixRun.Text;
        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
    }

    private void txtTablePrefix_TextChanged(object sender, EventArgs e)
    {
        if (clsBasUserData.sType == "CR")
            GenerateTableNameCR();
        else if (clsBasUserData.sType == "DB")
            GenerateTableNameDB();
        else if (clsBasUserData.sType == "CP")
            GenerateTableNameCP();
    }

    private void txtTablePrefix_Leave(object sender, EventArgs e)
    {
        if (clsBasUserData.sType == "CR")
            GenerateTableNameCR();
        else if (clsBasUserData.sType == "DB")
            GenerateTableNameDB();
        else if (clsBasUserData.sType == "CP")
            GenerateTableNameCP();
    }

    private void btnBanksMails_Click(object sender, EventArgs e)
    {
        new frmBanksMails(this).Show();
        this.Enabled = false;
    }

    private void CombineFiles_Click(object sender, EventArgs e)
    {

        string path = @"D:\Temp\P20Files\Statement\Billing\";
        //this.Enabled = false;

        try
        {
            var date = txtTblDetail
                            .Text
                            .Split('_')[1]
                            .Substring(0, 4);

            if (!Directory.Exists($"{path}20{date}"))
            {
                MessageBox.Show("There is no Folder with this Date");
                return;
            }


            string[] inputFiles = Directory.GetFiles($"{path}20{date}", "*.txt",
                                                   SearchOption.TopDirectoryOnly);

            if (File.Exists($@"{path}20{date}\{biliingFileFinalName}{date}.txt"))
            {
                DialogResult result = MessageBox.Show("There is exist file, do you want to overide it?", "Override Message", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                    File.Delete($@"{path}20{date}\{biliingFileFinalName}{date}.txt");
                else
                    return;

            }

            const int chunkSize = 2 * 1024; // 2KB
            if (File.Exists($@"{path}\20{date}\20{date}.txt"))
                File.Delete($@"{path}20{date}\20{date}.txt");

            using (var output = File.Create($@"{path}20{date}\{biliingFileFinalName}{date}.txt"))
            {
                foreach (var file in inputFiles)
                {
                    if (file == $@"{path}20{date}\{biliingFileFinalName}{date}.txt")
                    {
                        DialogResult SucessMsgBox = MessageBox.Show($"File {biliingFileFinalName}{date}.txt : has been generated Successfully", "Billing File", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    using (var input = File.OpenRead(file))
                    {
                        var buffer = new byte[chunkSize];
                        int bytesRead;
                        while ((bytesRead = input.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            output.Write(buffer, 0, bytesRead);
                        }
                    }
                }
                DialogResult EndSucessMsgBox = MessageBox.Show($"File {biliingFileFinalName}{date}.txt : has been generated Successfully", "Billing File", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error while Generating Files {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

    }

    private void chkRenameTables_CheckedChanged(object sender, EventArgs e)
    {
        if (chkRenameTables.Checked == true)
        {
            //txtTablePrefix.ReadOnly = true;
            txtPrefixRun.ReadOnly = true;
            //if (!chkTest.Checked)
            //    {
            if (clsBasUserData.sType == "CR")
                GenerateRunCR();
            else if (clsBasUserData.sType == "DB")
                GenerateRunDB();
            else if (clsBasUserData.sType == "CP")
                GenerateRunCP();
            //txtPrefixRun.ReadOnly = false;
            //}
        }
        else
        {
            txtPrefixRun.Text = "";
            txtPrefixRun.ReadOnly = false;
        }
        if (clsBasUserData.sType == "CR")
            GenerateTableNameCR();
        else if (clsBasUserData.sType == "DB")
            GenerateTableNameDB();
        else if (clsBasUserData.sType == "CP")
            GenerateTableNameCP();
    }
    private void GenerateTableNameCR()
    {
        txtTblMaster.Text = mstrCR;
        txtTblDetail.Text = dtlCR;
        if (txtTablePrefix.Text != "")//|| txtPrefixRun.Text != "")
        {
            txtTblMaster.Text = txtTblMaster.Text + "_" + txtTablePrefix.Text + txtPrefixRun.Text;
            txtTblDetail.Text = txtTblDetail.Text + "_" + txtTablePrefix.Text + txtPrefixRun.Text;
            if (txtTblMaster.Text.Length > 30 || txtTblDetail.Text.Length > 30)
            {
                txtTablePrefix.Text = "";
                txtPrefixRun.Text = "";
                txtTblMaster.Text = mstrCR;
                txtTblDetail.Text = dtlCR;
                MessageBox.Show("Table name must not exceed 30 Char(s)");
            }
        }
    }

    private void GenerateTableNameDB()
    {
        txtTblMaster.Text = mstrDB;
        txtTblDetail.Text = dtlDB;
        if (txtTablePrefix.Text != "")//|| txtPrefixRun.Text != "")
        {
            txtTblMaster.Text = txtTblMaster.Text + "_" + txtTablePrefix.Text + txtPrefixRun.Text;
            txtTblDetail.Text = txtTblDetail.Text + "_" + txtTablePrefix.Text + txtPrefixRun.Text;
            if (txtTblMaster.Text.Length > 30 || txtTblDetail.Text.Length > 30)
            {
                txtTablePrefix.Text = "";
                txtPrefixRun.Text = "";
                txtTblMaster.Text = mstrDB;
                txtTblDetail.Text = dtlDB;
                MessageBox.Show("Table name must not exceed 30 Char(s)");
            }
        }
    }

    private void GenerateTableNameCP()
    {
        txtTblMaster.Text = mstrCP;
        txtTblDetail.Text = dtlCP;
        if (txtTablePrefix.Text != "")//|| txtPrefixRun.Text != "")
        {
            txtTblMaster.Text = txtTblMaster.Text + "_" + txtTablePrefix.Text + txtPrefixRun.Text;
            txtTblDetail.Text = txtTblDetail.Text + "_" + txtTablePrefix.Text + txtPrefixRun.Text;
            if (txtTblMaster.Text.Length > 30 || txtTblDetail.Text.Length > 30)
            {
                txtTablePrefix.Text = "";
                txtPrefixRun.Text = "";
                txtTblMaster.Text = mstrCP;
                txtTblDetail.Text = dtlCP;
                MessageBox.Show("Table name must not exceed 30 Char(s)");
            }
        }
    }

    private void ChangeType()
    {
        if (clsBasUserData.sType == "CR")
        {
            comboBox1.Items.Clear();
            comboBox1.Items.AddRange(CyclesCR);
            comboBox2.Items.Clear();
            comboBox2.Items.AddRange(EStmtCR);
            this.Text = "Statement File " + strUserName + "@" + clsBasNetwork.getMyComputerName() + "@" + clsBasNetwork.getMyComputerIP() + " >> " + strServer + " " + clsBasUserData.loginDate.ToString("dd/MM/yyyy hh:mm:ss") + " - " + Application.StartupPath + " - Version:" + inf.Version + " - " + clsBasUserData.sType;
            GenerateTableNameCR();
        }

        else if (clsBasUserData.sType == "DB")
        {
            comboBox1.Items.Clear();
            comboBox1.Items.AddRange(CyclesDB);
            comboBox2.Items.Clear();
            comboBox2.Items.AddRange(EStmtDB);
            this.Text = "Statement File " + strUserName + "@" + clsBasNetwork.getMyComputerName() + "@" + clsBasNetwork.getMyComputerIP() + " >> " + strServer + " " + clsBasUserData.loginDate.ToString("dd/MM/yyyy hh:mm:ss") + " - " + Application.StartupPath + " - Version:" + inf.Version + " - " + clsBasUserData.sType;
            GenerateTableNameDB();
        }

        else if (clsBasUserData.sType == "CP")
        {
            comboBox1.Items.Clear();
            comboBox1.Items.AddRange(CyclesCP);
            comboBox2.Items.Clear();
            comboBox2.Items.AddRange(EStmtCP);
            this.Text = "Statement File " + strUserName + "@" + clsBasNetwork.getMyComputerName() + "@" + clsBasNetwork.getMyComputerIP() + " >> " + strServer + " " + clsBasUserData.loginDate.ToString("dd/MM/yyyy hh:mm:ss") + " - " + Application.StartupPath + " - Version:" + inf.Version + " - " + clsBasUserData.sType;
            GenerateTableNameCP();
        }
    }

    private void creditToolStripMenuItem_Click(object sender, EventArgs e)
    {
        clsBasUserData.sType = "CR";
        ChangeType();
    }

    private void corporateToolStripMenuItem_Click(object sender, EventArgs e)
    {
        clsBasUserData.sType = "CP";
        ChangeType();
    }

    private void debitToolStripMenuItem_Click(object sender, EventArgs e)
    {
        clsBasUserData.sType = "DB";
        ChangeType();
    }

    private void button1_Click(object sender, EventArgs e)
    {
        if (button1.BackColor == Color.Red)
        {
            StatementConfig st = new StatementConfig();
            st.Owner = this;
            st.Show();
        }
        else
        {
            if (MessageBox.Show("Are you sure that you want to send statement to the customers directly ?", "Submit", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Internal = false;
                InternalMailCC.Clear();
                InternalMailTo.Clear();
                InternalMailBCC.Clear();
                InternalMailFrom = "";
                InternalMailFromName = "";
                internalAccNo = "";
                MailCount = 0;
                button1.BackColor = Color.Red;
                button1.Text = "External Mode";
            }
        }
    }

}
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.OracleClient;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;


	// Summary description for frmRptReportManager.
public class frmRptReportManager : System.Windows.Forms.Form
{
	private System.Windows.Forms.TreeView trvReopts;
	private System.Windows.Forms.Button btnShowReport;
	private System.Windows.Forms.Button btnExit;
	private System.Windows.Forms.TextBox txtRunResult;
	private System.Windows.Forms.ImageList imgListReportsTree;
	private System.ComponentModel.IContainer components;
	private System.Windows.Forms.Label lblInstitution;
	private System.Windows.Forms.ComboBox cmbInstitution;
	private System.Windows.Forms.Label lblBranch;
	private System.Windows.Forms.ComboBox cmbBranch;
	private System.Windows.Forms.Label lblProduct;
	private System.Windows.Forms.ComboBox cmbProduct;
	private System.Windows.Forms.Label lblReportDate;
	private System.Windows.Forms.Button btnSaveFileName;
	private System.Windows.Forms.TextBox txtFileName;
	private System.Windows.Forms.Label label1;
	private System.Windows.Forms.Button btnExport;

	private ArrayList customerArray = new ArrayList();

	/*/---------
	private const string SUCCESS = "The action was successful.";
	private const string FAILURE = "The action was not successful: ";
	private const string NOT_ALLOWED = "You are not allowed to do this action.";
	private const string FORMAT_NOT_SUPPORTED = "That format is not supported.";
	private const string NO_MATCHES_FOUND = "No matches were found for the value submitted.";

	private ReportDocument hierarchicalGroupingReport;
	private string exportPath;
	private DiskFileDestinationOptions diskFileDestinationOptions;
	private ExportOptions exportOptions;
	private bool selectedNoFormat = false;
	/--*/

	private int reportID=0;
	private string reportName="";
	private int bankID=0, branchID=0, productID=0;
	private string strServer=string.Empty, strUserName=string.Empty
		, strPassword=string.Empty, strCon=string.Empty;

	private System.Windows.Forms.Label lblVer;
	private System.Windows.Forms.ComboBox cmbDatabase;
	private System.Windows.Forms.Label lblDatabase;
	private System.Windows.Forms.ComboBox cmbExportTypesList;
	private System.Windows.Forms.Button btnSaveData;
	private System.Windows.Forms.TextBox txtDelimit;
	private System.Windows.Forms.Button btnPDF;
	private System.Windows.Forms.TextBox txtReportName;
	private System.Windows.Forms.TextBox txtUserName;
	private System.Windows.Forms.TextBox txtPassword;
	private System.Windows.Forms.Button btnOK;
	private System.Windows.Forms.TextBox txtTest;
	private System.Windows.Forms.DateTimePicker datDate;

	public frmRptReportManager()
	{
		//
		// Required for Windows Form Designer support
		//
		InitializeComponent();

		//
		// TODO: Add any constructor code after InitializeComponent call
		//
	}

	// Clean up any resources being used.
	protected override void Dispose( bool disposing )
	{
		if( disposing )
		{
			if(components != null)
			{
				components.Dispose();
			}
		}
		base.Dispose( disposing );
		Application.Exit();

	}

	#region Windows Form Designer generated code
	/// <summary>
	/// Required method for Designer support - do not modify
	/// the contents of this method with the code editor.
	/// </summary>
	private void InitializeComponent()
	{
		this.components = new System.ComponentModel.Container();
		System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRptReportManager));
		this.trvReopts = new System.Windows.Forms.TreeView();
		this.imgListReportsTree = new System.Windows.Forms.ImageList(this.components);
		this.btnShowReport = new System.Windows.Forms.Button();
		this.btnExit = new System.Windows.Forms.Button();
		this.txtRunResult = new System.Windows.Forms.TextBox();
		this.lblInstitution = new System.Windows.Forms.Label();
		this.cmbInstitution = new System.Windows.Forms.ComboBox();
		this.lblBranch = new System.Windows.Forms.Label();
		this.cmbBranch = new System.Windows.Forms.ComboBox();
		this.lblProduct = new System.Windows.Forms.Label();
		this.cmbProduct = new System.Windows.Forms.ComboBox();
		this.datDate = new System.Windows.Forms.DateTimePicker();
		this.lblReportDate = new System.Windows.Forms.Label();
		this.btnSaveFileName = new System.Windows.Forms.Button();
		this.txtFileName = new System.Windows.Forms.TextBox();
		this.label1 = new System.Windows.Forms.Label();
		this.cmbExportTypesList = new System.Windows.Forms.ComboBox();
		this.btnExport = new System.Windows.Forms.Button();
		this.lblVer = new System.Windows.Forms.Label();
		this.cmbDatabase = new System.Windows.Forms.ComboBox();
		this.lblDatabase = new System.Windows.Forms.Label();
		this.btnSaveData = new System.Windows.Forms.Button();
		this.txtDelimit = new System.Windows.Forms.TextBox();
		this.txtReportName = new System.Windows.Forms.TextBox();
		this.btnPDF = new System.Windows.Forms.Button();
		this.txtUserName = new System.Windows.Forms.TextBox();
		this.txtPassword = new System.Windows.Forms.TextBox();
		this.btnOK = new System.Windows.Forms.Button();
		this.txtTest = new System.Windows.Forms.TextBox();
		this.SuspendLayout();
		// 
		// trvReopts
		// 
		this.trvReopts.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
			| System.Windows.Forms.AnchorStyles.Left) 
			| System.Windows.Forms.AnchorStyles.Right)));
		this.trvReopts.ImageList = this.imgListReportsTree;
		this.trvReopts.Location = new System.Drawing.Point(8, 56);
		this.trvReopts.Name = "trvReopts";
		this.trvReopts.Size = new System.Drawing.Size(260, 192);
		this.trvReopts.TabIndex = 1;
		this.trvReopts.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.trvReopts_AfterSelect);
		// 
		// imgListReportsTree
		// 
		this.imgListReportsTree.ImageSize = new System.Drawing.Size(16, 16);
		this.imgListReportsTree.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imgListReportsTree.ImageStream")));
		this.imgListReportsTree.TransparentColor = System.Drawing.Color.Transparent;
		// 
		// btnShowReport
		// 
		this.btnShowReport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
		this.btnShowReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.btnShowReport.Location = new System.Drawing.Point(55, 336);
		this.btnShowReport.Name = "btnShowReport";
		this.btnShowReport.Size = new System.Drawing.Size(85, 20);
		this.btnShowReport.TabIndex = 10;
		this.btnShowReport.Text = "&Show Report";
		this.btnShowReport.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.btnShowReport.Click += new System.EventHandler(this.btnShowReport_Click);
		// 
		// btnExit
		// 
		this.btnExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
		this.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		this.btnExit.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.btnExit.Location = new System.Drawing.Point(9, 336);
		this.btnExit.Name = "btnExit";
		this.btnExit.Size = new System.Drawing.Size(39, 20);
		this.btnExit.TabIndex = 11;
		this.btnExit.Text = "&Exit";
		this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
		// 
		// txtRunResult
		// 
		this.txtRunResult.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
			| System.Windows.Forms.AnchorStyles.Right)));
		this.txtRunResult.Location = new System.Drawing.Point(272, 252);
		this.txtRunResult.Multiline = true;
		this.txtRunResult.Name = "txtRunResult";
		this.txtRunResult.Size = new System.Drawing.Size(436, 72);
		this.txtRunResult.TabIndex = 4;
		this.txtRunResult.Text = "";
		// 
		// lblInstitution
		// 
		this.lblInstitution.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
		this.lblInstitution.BackColor = System.Drawing.Color.Transparent;
		this.lblInstitution.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.lblInstitution.Location = new System.Drawing.Point(272, 29);
		this.lblInstitution.Name = "lblInstitution";
		this.lblInstitution.Size = new System.Drawing.Size(60, 20);
		this.lblInstitution.TabIndex = 5;
		this.lblInstitution.Text = "Institution";
		// 
		// cmbInstitution
		// 
		this.cmbInstitution.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
		this.cmbInstitution.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.cmbInstitution.Location = new System.Drawing.Point(330, 28);
		this.cmbInstitution.Name = "cmbInstitution";
		this.cmbInstitution.Size = new System.Drawing.Size(182, 23);
		this.cmbInstitution.TabIndex = 2;
		this.cmbInstitution.SelectedIndexChanged += new System.EventHandler(this.cmbInstitution_SelectedIndexChanged);
		// 
		// lblBranch
		// 
		this.lblBranch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
		this.lblBranch.BackColor = System.Drawing.Color.Transparent;
		this.lblBranch.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.lblBranch.Location = new System.Drawing.Point(517, 29);
		this.lblBranch.Name = "lblBranch";
		this.lblBranch.Size = new System.Drawing.Size(47, 20);
		this.lblBranch.TabIndex = 5;
		this.lblBranch.Text = "Branch";
		// 
		// cmbBranch
		// 
		this.cmbBranch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
		this.cmbBranch.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.cmbBranch.Location = new System.Drawing.Point(560, 28);
		this.cmbBranch.Name = "cmbBranch";
		this.cmbBranch.Size = new System.Drawing.Size(144, 23);
		this.cmbBranch.TabIndex = 3;
		this.cmbBranch.SelectedIndexChanged += new System.EventHandler(this.cmbBranch_SelectedIndexChanged);
		// 
		// lblProduct
		// 
		this.lblProduct.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
		this.lblProduct.BackColor = System.Drawing.Color.Transparent;
		this.lblProduct.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.lblProduct.Location = new System.Drawing.Point(272, 56);
		this.lblProduct.Name = "lblProduct";
		this.lblProduct.Size = new System.Drawing.Size(60, 23);
		this.lblProduct.TabIndex = 5;
		this.lblProduct.Text = "Product";
		// 
		// cmbProduct
		// 
		this.cmbProduct.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
		this.cmbProduct.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.cmbProduct.Location = new System.Drawing.Point(330, 56);
		this.cmbProduct.Name = "cmbProduct";
		this.cmbProduct.Size = new System.Drawing.Size(182, 23);
		this.cmbProduct.TabIndex = 4;
		this.cmbProduct.SelectedIndexChanged += new System.EventHandler(this.cmbProduct_SelectedIndexChanged);
		// 
		// datDate
		// 
		this.datDate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
		this.datDate.CustomFormat = "yyyy-MM-dd hh:mm:ss";
		this.datDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.datDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
		this.datDate.Location = new System.Drawing.Point(560, 56);
		this.datDate.Name = "datDate";
		this.datDate.ShowUpDown = true;
		this.datDate.Size = new System.Drawing.Size(144, 23);
		this.datDate.TabIndex = 5;
		// 
		// lblReportDate
		// 
		this.lblReportDate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
		this.lblReportDate.BackColor = System.Drawing.Color.Transparent;
		this.lblReportDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.lblReportDate.Location = new System.Drawing.Point(517, 56);
		this.lblReportDate.Name = "lblReportDate";
		this.lblReportDate.Size = new System.Drawing.Size(48, 23);
		this.lblReportDate.TabIndex = 5;
		this.lblReportDate.Text = "Date";
		// 
		// btnSaveFileName
		// 
		this.btnSaveFileName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
		this.btnSaveFileName.Image = ((System.Drawing.Image)(resources.GetObject("btnSaveFileName.Image")));
		this.btnSaveFileName.Location = new System.Drawing.Point(239, 280);
		this.btnSaveFileName.Name = "btnSaveFileName";
		this.btnSaveFileName.Size = new System.Drawing.Size(28, 20);
		this.btnSaveFileName.TabIndex = 7;
		this.btnSaveFileName.Click += new System.EventHandler(this.btnSaveFileName_Click);
		// 
		// txtFileName
		// 
		this.txtFileName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
		this.txtFileName.Location = new System.Drawing.Point(36, 280);
		this.txtFileName.Name = "txtFileName";
		this.txtFileName.Size = new System.Drawing.Size(200, 20);
		this.txtFileName.TabIndex = 6;
		this.txtFileName.Text = "C:\\TEMP\\P20Files\\Statement\\";
		// 
		// label1
		// 
		this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
		this.label1.BackColor = System.Drawing.Color.Transparent;
		this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.label1.Location = new System.Drawing.Point(8, 280);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(40, 20);
		this.label1.TabIndex = 16;
		this.label1.Text = "Path";
		// 
		// cmbExportTypesList
		// 
		this.cmbExportTypesList.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
		this.cmbExportTypesList.Location = new System.Drawing.Point(8, 305);
		this.cmbExportTypesList.Name = "cmbExportTypesList";
		this.cmbExportTypesList.Size = new System.Drawing.Size(208, 21);
		this.cmbExportTypesList.TabIndex = 8;
		// 
		// btnExport
		// 
		this.btnExport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
		this.btnExport.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.btnExport.Location = new System.Drawing.Point(220, 305);
		this.btnExport.Name = "btnExport";
		this.btnExport.Size = new System.Drawing.Size(48, 20);
		this.btnExport.TabIndex = 9;
		this.btnExport.Text = "E&xport";
		this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
		// 
		// lblVer
		// 
		this.lblVer.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
		this.lblVer.BackColor = System.Drawing.Color.Transparent;
		this.lblVer.ForeColor = System.Drawing.Color.Blue;
		this.lblVer.Location = new System.Drawing.Point(604, 0);
		this.lblVer.Name = "lblVer";
		this.lblVer.Size = new System.Drawing.Size(104, 16);
		this.lblVer.TabIndex = 18;
		this.lblVer.Text = "Ver 1.010";
		this.lblVer.TextAlign = System.Drawing.ContentAlignment.TopRight;
		this.lblVer.DoubleClick += new System.EventHandler(this.lblVer_DoubleClick);
		// 
		// cmbDatabase
		// 
		this.cmbDatabase.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.cmbDatabase.Location = new System.Drawing.Point(71, 4);
		this.cmbDatabase.Name = "cmbDatabase";
		this.cmbDatabase.Size = new System.Drawing.Size(149, 23);
		this.cmbDatabase.TabIndex = 0;
		this.cmbDatabase.SelectedIndexChanged += new System.EventHandler(this.cmbDatabase_SelectedIndexChanged);
		// 
		// lblDatabase
		// 
		this.lblDatabase.BackColor = System.Drawing.Color.Transparent;
		this.lblDatabase.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.lblDatabase.Location = new System.Drawing.Point(11, 5);
		this.lblDatabase.Name = "lblDatabase";
		this.lblDatabase.Size = new System.Drawing.Size(60, 20);
		this.lblDatabase.TabIndex = 19;
		this.lblDatabase.Text = "Database";
		// 
		// btnSaveData
		// 
		this.btnSaveData.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
		this.btnSaveData.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.btnSaveData.Location = new System.Drawing.Point(185, 336);
		this.btnSaveData.Name = "btnSaveData";
		this.btnSaveData.Size = new System.Drawing.Size(68, 20);
		this.btnSaveData.TabIndex = 22;
		this.btnSaveData.Text = "Save &Data";
		this.btnSaveData.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.btnSaveData.Click += new System.EventHandler(this.btnSaveData_Click);
		// 
		// txtDelimit
		// 
		this.txtDelimit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
		this.txtDelimit.Location = new System.Drawing.Point(248, 336);
		this.txtDelimit.Name = "txtDelimit";
		this.txtDelimit.Size = new System.Drawing.Size(16, 20);
		this.txtDelimit.TabIndex = 23;
		this.txtDelimit.Text = "|";
		// 
		// txtReportName
		// 
		this.txtReportName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
		this.txtReportName.Location = new System.Drawing.Point(8, 253);
		this.txtReportName.Name = "txtReportName";
		this.txtReportName.Size = new System.Drawing.Size(220, 20);
		this.txtReportName.TabIndex = 24;
		this.txtReportName.Text = "";
		// 
		// btnPDF
		// 
		this.btnPDF.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
		this.btnPDF.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.btnPDF.Location = new System.Drawing.Point(230, 253);
		this.btnPDF.Name = "btnPDF";
		this.btnPDF.Size = new System.Drawing.Size(38, 20);
		this.btnPDF.TabIndex = 25;
		this.btnPDF.Text = "&PDF";
		this.btnPDF.Click += new System.EventHandler(this.btnPDF_Click);
		// 
		// txtUserName
		// 
		this.txtUserName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.txtUserName.Location = new System.Drawing.Point(8, 30);
		this.txtUserName.Name = "txtUserName";
		this.txtUserName.Size = new System.Drawing.Size(124, 22);
		this.txtUserName.TabIndex = 26;
		this.txtUserName.Text = "";
		// 
		// txtPassword
		// 
		this.txtPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.txtPassword.Location = new System.Drawing.Point(136, 30);
		this.txtPassword.Name = "txtPassword";
		this.txtPassword.PasswordChar = '*';
		this.txtPassword.Size = new System.Drawing.Size(132, 22);
		this.txtPassword.TabIndex = 27;
		this.txtPassword.Text = "";
		// 
		// btnOK
		// 
		this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
		this.btnOK.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
		this.btnOK.Location = new System.Drawing.Point(224, 5);
		this.btnOK.Name = "btnOK";
		this.btnOK.Size = new System.Drawing.Size(44, 20);
		this.btnOK.TabIndex = 28;
		this.btnOK.Text = "&Login";
		this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
		// 
		// txtTest
		// 
		this.txtTest.Location = new System.Drawing.Point(272, 228);
		this.txtTest.Name = "txtTest";
		this.txtTest.Size = new System.Drawing.Size(80, 20);
		this.txtTest.TabIndex = 29;
		this.txtTest.Text = "";
		this.txtTest.Visible = false;
		// 
		// frmRptReportManager
		// 
		this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
		this.BackColor = System.Drawing.SystemColors.Window;
		this.ClientSize = new System.Drawing.Size(712, 365);
		this.Controls.Add(this.txtTest);
		this.Controls.Add(this.btnOK);
		this.Controls.Add(this.txtPassword);
		this.Controls.Add(this.txtUserName);
		this.Controls.Add(this.txtReportName);
		this.Controls.Add(this.txtDelimit);
		this.Controls.Add(this.txtFileName);
		this.Controls.Add(this.txtRunResult);
		this.Controls.Add(this.cmbProduct);
		this.Controls.Add(this.cmbBranch);
		this.Controls.Add(this.btnPDF);
		this.Controls.Add(this.btnSaveData);
		this.Controls.Add(this.cmbDatabase);
		this.Controls.Add(this.lblDatabase);
		this.Controls.Add(this.lblVer);
		this.Controls.Add(this.cmbExportTypesList);
		this.Controls.Add(this.label1);
		this.Controls.Add(this.btnSaveFileName);
		this.Controls.Add(this.datDate);
		this.Controls.Add(this.cmbInstitution);
		this.Controls.Add(this.lblInstitution);
		this.Controls.Add(this.btnExit);
		this.Controls.Add(this.btnShowReport);
		this.Controls.Add(this.trvReopts);
		this.Controls.Add(this.lblBranch);
		this.Controls.Add(this.lblProduct);
		this.Controls.Add(this.lblReportDate);
		this.Controls.Add(this.btnExport);
		this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
		this.Name = "frmRptReportManager";
		this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
		this.Text = "Report Data";
		this.Load += new System.EventHandler(this.frmRptReportManager_Load);
		this.ResumeLayout(false);

	}
	#endregion

	private void btnExit_Click(object sender, System.EventArgs e)
	{
		this.Dispose(true);  //this.Close();
	}

	private void frmRptReportManager_Load(object sender, System.EventArgs e)
	{
		AssemblyInfo ainfo = new AssemblyInfo();
		lblVer.Text = "Ver " + ainfo.Version;

		Cursor.Current = Cursors.Default;
		strServer = clsDbCon.sServer;
		strUserName = clsDbCon.sUserName;
		txtUserName.Text = strUserName;
		strPassword = clsDbCon.sPassword;
		strCon ="SERVER=" + strServer + ";UID=" + strUserName + ";" + "PASSWORD=" + strPassword + ";";

		if (System.IO.File.Exists(Application.StartupPath + "\\frmBackground.jpg"))
		{
			this.BackgroundImage = Image.FromFile(Application.StartupPath + "\\frmBackground.jpg");
		}


		string[] strBanksName;
		strBanksName = basText.spliteStrAt(clsBasUserData.readKey("Settings","ServerName"),",");
		foreach ( string strBank  in strBanksName )
		{
			cmbDatabase.Items.Add (strBank);
		}

		if (cmbDatabase.Items.Count > 0)
			cmbDatabase.SelectedIndex =0;

		fillTreeReports();
		fillCombo();
	}

	private void fillTreeReports()
	{
		//ImageList il = new ImageList();
		//il.Images.Add(new Icon("KEY04.ICO"));
		//il.Images.Add(new Icon("ARW06LT.ICO"));
		//treeView1.ImageList = il ; 

		TreeNode masterRootNode, detailRootNode;
		trvReopts.BeginUpdate();
		DataSet groupDataSet = new DataSet();
		DataSet reportsDataSet = new DataSet();
		groupDataSet =clsDbOracleLayer.getDataset("Select ReportGroupCode, reportgroupname From a4m.r_reportgroup where reportgroupmaincode = 1","QueryResault", strCon);
		foreach (DataRow masterRow in groupDataSet.Tables["QueryResault"].Rows)
		{
			masterRootNode = new TreeNode(masterRow[1].ToString(), 0, 0);
			masterRootNode.Tag = "G" + masterRow[0].ToString();		//Group
			trvReopts.Nodes.Add(masterRootNode);
			masterRootNode.Expand();
			reportsDataSet =clsDbOracleLayer.getDataset("Select REPORTDESCCODE, REPORTDESCNAME From a4m.R_REPORTDESC where reportgroupcode =" + masterRow[0].ToString(),"QueryResault", strCon);
			foreach (DataRow detailRow in reportsDataSet.Tables["QueryResault"].Rows)
			{
				detailRootNode = new TreeNode(detailRow[1].ToString(), 1, 1);
				detailRootNode.Tag = "R" + detailRow[0].ToString();		//Report
				masterRootNode.Nodes.Add(detailRootNode);
				detailRootNode.Expand();
			}

		}
		trvReopts.EndUpdate();
	}

	
	private void fillCombo()
	{
		OracleDataReader drSQL;
		ListItem objListItem;

		// fill Institution
		cmbInstitution.Items.Clear();
		//drSQL =clsDbOracleLayer.getDataReader("Select BRANCH , NAME || ' >> ' || FIID From TREFERENCEFI where ACCOUNTGROUPCODE IS NULL order by NAME");
		drSQL =clsDbOracleLayer.getDataReader("Select BankCode, bankName From a4m.v_r_Banks", strCon);
		while (drSQL.Read())
		{
			objListItem = new ListItem(drSQL[1].ToString(),Convert.ToInt32(drSQL[0]));
			cmbInstitution.Items.Add(objListItem);
		}

		if(cmbInstitution.Items.Count > 0)
			cmbInstitution.SelectedIndex = 0;
			
		// fill export types
		cmbExportTypesList.DataSource = System.Enum.GetValues(typeof(ExportFormatType));
	}

	private void cmbInstitution_SelectedIndexChanged(object sender, System.EventArgs e)
	{
		OracleDataReader drSQL;
		ListItem objListItem;

		bankID =((ListItem) cmbInstitution.Items[cmbInstitution.SelectedIndex]).ID;

		// fill Branch
		cmbBranch.Items.Clear();
		//drSQL =clsDbOracleLayer.getDataReader("Select BRANCHPART , NAME From TBRANCHPART  where BRANCH = " + bankID + " order by NAME");
		drSQL =clsDbOracleLayer.getDataReader("Select branchCode , branchName From a4m.v_r_Branches where BankCode = " + bankID + " order by branchName", strCon);

		// all Branches
		objListItem = new ListItem("All Branches".ToString(),0);
		cmbBranch.Items.Add(objListItem);
		while (drSQL.Read())
		{
			objListItem = new ListItem(drSQL[1].ToString(),Convert.ToInt32(drSQL[0]));
			cmbBranch.Items.Add(objListItem);
		}

		if (cmbBranch.Items.Count > 0)
			cmbBranch.SelectedIndex =0;


		// fill Product
		cmbProduct.Items.Clear();
		//drSQL =clsDbOracleLayer.getDataReader("Select CODE, NAME From TREFERENCECARDPRODUCT where BRANCH = " + bankID + " order by NAME");
		drSQL =clsDbOracleLayer.getDataReader("Select productCode, productName From a4m.v_r_Products where BankCode = " + bankID + " order by productName", strCon);
		
		// All Products
		objListItem = new ListItem("All Products",0);
		cmbProduct.Items.Add(objListItem);
		while (drSQL.Read())
		{
			objListItem = new ListItem(drSQL[1].ToString(),Convert.ToInt32(drSQL[0]));
			cmbProduct.Items.Add(objListItem);
		}

		if (cmbProduct.Items.Count > 0)
			cmbProduct.SelectedIndex =0; 
	}

	private void trvReopts_AfterSelect(object sender, System.Windows.Forms.TreeViewEventArgs e)
	{
		if (e.Action != TreeViewAction.Unknown)
		{
			if (e.Node.Tag.ToString().Substring(0,1) == "R") 
			{
				reportID =Convert.ToInt32(e.Node.Tag.ToString().Substring(1));
				reportName = e.Node.Text;
				txtReportName.Text = e.Node.Text;
				txtRunResult.Text = reportName + "\r\n\r\n" + basText.Left(txtRunResult.Text,2000);

			}
		}
	}


	private void cmbBranch_SelectedIndexChanged(object sender, System.EventArgs e)
	{
		branchID =((ListItem) cmbBranch.Items[cmbBranch.SelectedIndex]).ID;
	}

	private void cmbProduct_SelectedIndexChanged(object sender, System.EventArgs e)
	{
		productID =((ListItem) cmbProduct.Items[cmbProduct.SelectedIndex]).ID;
	}


	private void btnShowReport_Click(object sender, System.EventArgs e)
	{
		if(reportID < 1) return;
		string strSQL = string.Empty ;
		Cursor.Current = Cursors.WaitCursor;
		frmBasRptShow frmRptShowTest = new frmBasRptShow();

		OracleDataReader drSQL;
		drSQL =clsDbOracleLayer.getDataReader("Select r.reportdescname, r.reportrptname,r.datasourcenam From a4m.r_reportdesc r where r.reportdesccode = " + reportID, strCon);
		while (drSQL.Read())
		{
			frmRptShowTest.sReortName = clsBasUserData.appPath() + @"Reports\" + drSQL["reportrptname"].ToString();//Cards Statistics.rpt
			frmRptShowTest.sCaption = drSQL["reportdescname"].ToString();   //"Rpt Preview";
			strSQL = drSQL["datasourcenam"].ToString();
			if(strSQL.Trim() != string.Empty)
				strSQL = ReplaceParm(strSQL);//drSQL["datasourcenam"].ToString()
			frmRptShowTest.sRptSqlQuery = strSQL; //"SELECT * from V1 where BRANCH = 3";
		}
		txtRunResult.Text = strSQL + "\r\n\r\n" + basText.Left(txtRunResult.Text,2000);

		frmRptShowTest.Show();
		//testPackage();

	}


	public string ReplaceParm(string pSqlStr)
	{
		string rtrnSql=pSqlStr, strSelect, strWhere = string.Empty ;
		int strPos = -1,startPos = -1, endPos = -1,mStartPos = -1, mEndPos = -1 ;
		string frmtStr ="", rplcStr ="", mRplcStr ="";
		int lastIndx = -1,startIndx = -1;

		pSqlStr = pSqlStr.ToLower();
		startIndx = rtrnSql.IndexOf(" where ");
		if(startIndx > 0)
		{
			strSelect = rtrnSql.Substring(0,startIndx + " where ".Length -1 );
			strWhere = rtrnSql.Substring(startIndx + " where ".Length);

			if(strWhere.IndexOf("<Institution")> 0)
			{
				mStartPos = strWhere.IndexOf("<Institution");
				mEndPos = strWhere.IndexOf(">",mStartPos);
				mRplcStr = strWhere.Substring(mStartPos,mEndPos-mStartPos+1); 

				strWhere = basText.replaceStr(strWhere,mRplcStr,bankID.ToString());
			}



			if(strWhere.IndexOf("<Product")> 0)
			{
				mStartPos = strWhere.IndexOf("<Product");
				mEndPos = strWhere.IndexOf(">",mStartPos);
				mRplcStr = strWhere.Substring(mStartPos,mEndPos-mStartPos+1); 

				if(cmbProduct.SelectedIndex !=0)
					strWhere = basText.replaceStr(strWhere,mRplcStr,productID.ToString());
				else
				{
					strWhere = excludeCond(strWhere,mRplcStr);
				}
			}


			if(strWhere.IndexOf("<Branch>")> 0)
			{
				mStartPos = strWhere.IndexOf("<Product");
				mEndPos = strWhere.IndexOf(">",mStartPos);
				mRplcStr = strWhere.Substring(mStartPos,mEndPos-mStartPos+1); 

				if(cmbBranch.SelectedIndex !=0)
					strWhere = basText.replaceStr(strWhere,mRplcStr,branchID.ToString());
				else
				{
					strWhere = excludeCond(strWhere,mRplcStr);
				}
			}

			while(( strPos = strWhere.LastIndexOf("<Date")) > 0) //if
			{
				strWhere = strWhere.Replace('m','M'); //for month in date
				mStartPos = strWhere.IndexOf("<Date");
				mEndPos = strWhere.IndexOf(">",mStartPos);
				mRplcStr = strWhere.Substring(mStartPos,mEndPos-mStartPos+1); 

				startPos = strWhere.IndexOf("'",strPos)+1;
				endPos = strWhere.IndexOf("'",startPos +1 );
				frmtStr = strWhere.Substring(startPos,endPos-startPos); 
				rplcStr = "'" + basText.formatDate(datDate.Value.ToLongDateString(),frmtStr) + "'";
				//rplcStr = "'" + String.Format("{0:"+frmtStr+"}" , datDate.Value) + "'";
				strWhere =  basText.Replace(strWhere,mRplcStr,rplcStr);
			}
		}
		else
		{
			strSelect = rtrnSql;
		}

		strWhere = strWhere.Trim();
		if(strWhere.Substring(0,4) == "and ")
			strWhere = strWhere.Substring(4);
		if(strWhere.Substring(0,3) == "or ")
			strWhere = strWhere.Substring(3);
		if(strWhere.Substring(0,4) == "not ")
			strWhere = strWhere.Substring(4);



		
		rtrnSql =strSelect + " " + strWhere;
		rtrnSql = rtrnSql.Trim();
		if(rtrnSql.Substring(rtrnSql.Length - 5,5)== "WHERE")
			rtrnSql = rtrnSql.Substring(0,rtrnSql.Length - 5);


		return rtrnSql;
	}


	private void btnExport_Click(object sender, System.EventArgs e)
	{
		if(reportID < 1) return;
		string strSQL = string.Empty, strReportName = string.Empty;
		Cursor.Current = Cursors.WaitCursor;

		OracleDataReader drSQL;
		drSQL =clsDbOracleLayer.getDataReader("Select r.reportdescname, r.reportrptname,r.datasourcenam From a4m.r_reportdesc r where r.reportdesccode = " + reportID, strCon);
		while (drSQL.Read())
		{
			strReportName = clsBasUserData.appPath() + @"Reports\" + drSQL["reportrptname"].ToString();//Cards Statistics.rpt
			strSQL = drSQL["datasourcenam"].ToString();
			if(strSQL.Trim() != string.Empty)
				strSQL = ReplaceParm(strSQL);//drSQL["datasourcenam"].ToString()
		}
		txtRunResult.Text = strSQL + "\r\n\r\n" + txtRunResult.Text;

		clsExportReport expReport = new clsExportReport();
		expReport.ExportSetup(txtFileName.Text ,txtReportName.Text,strReportName,strSQL);
		expReport.ExportSelection((ExportFormatType)cmbExportTypesList.SelectedIndex);
		txtRunResult.Text = expReport.ExportCompletion() + "\r\n\r\n" + basText.Left(txtRunResult.Text,2000);
		Cursor.Current = Cursors.Default;
	}




	private void cmbDatabase_SelectedIndexChanged(object sender, System.EventArgs e)
	{
		clsDbCon.sServer = cmbDatabase.Text;
	}


	private string excludeCond(string pStr,string pCond)
	{
		string rtrnSql;
		int lastIndx = -1,startIndx = -1;
		rtrnSql = pStr;
		lastIndx = rtrnSql.IndexOf(pCond) + pCond.Length; 
		startIndx = rtrnSql.LastIndexOf(" and ", lastIndx);
		if(startIndx < 0)
			startIndx = rtrnSql.LastIndexOf(" or ", lastIndx);
		if(startIndx < 0)
			startIndx = rtrnSql.LastIndexOf(" not ", lastIndx);
		if(startIndx < 0)
		{
			startIndx = rtrnSql.LastIndexOf(" where ", lastIndx);
			if(startIndx > 0)
				startIndx = startIndx + " where ".Length;
		}

		if(lastIndx > 0 && startIndx < 0)
			startIndx = 0;

		rtrnSql = rtrnSql.Remove(startIndx,lastIndx-startIndx);
		return rtrnSql;
	}


	private void lblVer_DoubleClick(object sender, System.EventArgs e)
	{
		if(txtPassword.Text == "IOP")
		{
			Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
			new frmQuery().Show();
		}
	}

	private void btnSaveFileName_Click(object sender, System.EventArgs e)
	{
		txtFileName.Text = clsBasFile.makeStrAsPath(clsBasFile.openDirDialog("Open Path"));
	}

	private void btnSaveData_Click(object sender, System.EventArgs e)
	{
		if(reportID < 1) return;
		string strSQL = string.Empty ;
		string strDelimit = txtDelimit.Text.Trim();
		Cursor.Current = Cursors.WaitCursor;

		OracleDataReader drSQL;
		drSQL =clsDbOracleLayer.getDataReader("Select r.reportdescname, r.reportrptname,r.datasourcenam From a4m.r_reportdesc r where r.reportdesccode = " + reportID, strCon);
		while (drSQL.Read())
		{
			strSQL = drSQL["datasourcenam"].ToString();
			if(strSQL.Trim() != string.Empty)
				strSQL = ReplaceParm(strSQL);
			clsExportData exportSqlReport = new clsExportData();
			exportSqlReport.sql2File(strSQL,txtFileName.Text.Trim()+ txtReportName.Text + ".TXT",strDelimit);//drSQL["reportdescname"].ToString()
		}
		txtRunResult.Text = strSQL + "\r\n\r\n" + basText.Left(txtRunResult.Text,2000);

		Cursor.Current = Cursors.Default;
	}

	private void btnPDF_Click(object sender, System.EventArgs e)
	{
		if(reportID < 1) return;
		string strSQL = string.Empty, strReportName = string.Empty;
		Cursor.Current = Cursors.WaitCursor;

		OracleDataReader drSQL;
		drSQL =clsDbOracleLayer.getDataReader("Select r.reportdescname, r.reportrptname,r.datasourcenam From a4m.r_reportdesc r where r.reportdesccode = " + reportID, strCon);
		while (drSQL.Read())
		{
			strReportName = clsBasUserData.appPath() + @"Reports\" + drSQL["reportrptname"].ToString();//Cards Statistics.rpt
			strSQL = drSQL["datasourcenam"].ToString();
			if(strSQL.Trim() != string.Empty)
				strSQL = ReplaceParm(strSQL);//drSQL["datasourcenam"].ToString()
		}
		txtRunResult.Text = strSQL + "\r\n\r\n" + txtRunResult.Text;

		clsExportReport expReport = new clsExportReport();
		expReport.ExportSetup(txtFileName.Text ,txtReportName.Text,strReportName,strSQL);
		expReport.ExportSelection(ExportFormatType.PortableDocFormat);  //(ExportFormatType)cmbExportTypesList.SelectedIndex
		txtRunResult.Text = expReport.ExportCompletion() + "\r\n\r\n" + basText.Left(txtRunResult.Text,2000);
		Cursor.Current = Cursors.Default;

/*		if(reportID < 1) return;
		string strSQL = string.Empty, strReportName = string.Empty;
		Cursor.Current = Cursors.WaitCursor;

		OracleDataReader drSQL;
		drSQL =clsDbOracleLayer.getDataReader("Select r.reportdescname, r.reportrptname,r.datasourcenam From a4m.r_reportdesc r where r.reportdesccode = " + reportID);
		while (drSQL.Read())
		{
			strReportName = clsBasUserData.appPath() + @"Reports\" + drSQL["reportrptname"].ToString();//Cards Statistics.rpt
			strSQL = drSQL["datasourcenam"].ToString();
			if(strSQL.Trim() != string.Empty)
				strSQL = ReplaceParm(strSQL);
			clsExportData exportSqlReport = new clsExportData();
			exportSqlReport.sql2File(strSQL,txtFileName.Text.Trim()+ txtReportName.Text + ".TXT","");//drSQL["reportdescname"].ToString()
		}
		clsExportReport expReport = new clsExportReport();
		expReport.exportPDF(txtFileName.Text,txtReportName.Text,strReportName,strSQL);

		txtRunResult.Text = strSQL + "\r\n\r\n" + basText.Left(txtRunResult.Text,2000);

		Cursor.Current = Cursors.Default;
		*/
	}

	private void btnOK_Click(object sender, System.EventArgs e)
	{
		clsDbCon.sServer = cmbDatabase.Text ; //   "tw1tst";
		clsDbCon.sUserName = txtUserName.Text ; // "a4m";
		clsDbCon.sPassword =txtPassword.Text ; //"a4m";
	}


	void testPackage()
	{
		if(reportID < 1) return;
		string strSQL = string.Empty ;
		Cursor.Current = Cursors.WaitCursor;
		frmBasRptShow frmRptShowTest = new frmBasRptShow();

		OracleDataReader drSQL;
			frmRptShowTest.sReortName = clsBasUserData.appPath() + @"Reports\TestPackage.rpt" ;
			frmRptShowTest.sCaption = "TestPackage";   //"Rpt Preview";
			if(strSQL.Trim() != string.Empty)
				strSQL = ReplaceParm(strSQL);//drSQL["datasourcenam"].ToString()
			frmRptShowTest.sRptSqlQuery = "BEGIN A4M.TESTPACKAGE.TEST_REPORT(:TEST_TYPE_CUR, 48); END ;"; //"SELECT * from V1 where BRANCH = 3";

		txtRunResult.Text = strSQL + "\r\n\r\n" + basText.Left(txtRunResult.Text,2000);
		frmRptShowTest.Show();
	}




}


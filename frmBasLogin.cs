using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Configuration;


//
public class frmBasLogin : System.Windows.Forms.Form
    {
    private System.Windows.Forms.Button btnOK;
    private System.Windows.Forms.Button btnCancel;
    private System.Windows.Forms.TextBox txtUserName;
    private System.Windows.Forms.Label lblUserName;
    private System.Windows.Forms.Label lblPassword;
    private System.Windows.Forms.TextBox txtPassword;
    private System.Windows.Forms.GroupBox grpLogin;
    private System.Windows.Forms.Label lblVersion;
    private System.Windows.Forms.Label lblDatabase;
    private TextBox txtServerName;
    private ComboBox cmbDatabase;
    private Label lblDatabaseName;
    private CheckBox chkExternal;
    private ComboBox cmbType;
    private Label lblType;
    private IContainer components;
    // 
    public frmBasLogin()
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
        }

    #region Windows Form Designer generated code
    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
        {
        this.btnOK = new System.Windows.Forms.Button();
        this.btnCancel = new System.Windows.Forms.Button();
        this.grpLogin = new System.Windows.Forms.GroupBox();
        this.cmbType = new System.Windows.Forms.ComboBox();
        this.lblType = new System.Windows.Forms.Label();
        this.chkExternal = new System.Windows.Forms.CheckBox();
        this.cmbDatabase = new System.Windows.Forms.ComboBox();
        this.txtUserName = new System.Windows.Forms.TextBox();
        this.lblDatabaseName = new System.Windows.Forms.Label();
        this.lblUserName = new System.Windows.Forms.Label();
        this.lblPassword = new System.Windows.Forms.Label();
        this.txtPassword = new System.Windows.Forms.TextBox();
        this.lblVersion = new System.Windows.Forms.Label();
        this.lblDatabase = new System.Windows.Forms.Label();
        this.txtServerName = new System.Windows.Forms.TextBox();
        this.grpLogin.SuspendLayout();
        this.SuspendLayout();
        // 
        // btnOK
        // 
        this.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Left;
        this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
        this.btnOK.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
        this.btnOK.Location = new System.Drawing.Point(105, 85);
        this.btnOK.Name = "btnOK";
        this.btnOK.Size = new System.Drawing.Size(64, 24);
        this.btnOK.TabIndex = 3;
        this.btnOK.Text = "&OK";
        this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
        // 
        // btnCancel
        // 
        this.btnCancel.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
        this.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Left;
        this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
        this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
        this.btnCancel.Location = new System.Drawing.Point(181, 85);
        this.btnCancel.Name = "btnCancel";
        this.btnCancel.Size = new System.Drawing.Size(64, 24);
        this.btnCancel.TabIndex = 4;
        this.btnCancel.Text = "&Cancel";
        this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
        // 
        // grpLogin
        // 
        this.grpLogin.Controls.Add(this.cmbType);
        this.grpLogin.Controls.Add(this.lblType);
        this.grpLogin.Controls.Add(this.chkExternal);
        this.grpLogin.Controls.Add(this.cmbDatabase);
        this.grpLogin.Controls.Add(this.txtUserName);
        this.grpLogin.Controls.Add(this.lblDatabaseName);
        this.grpLogin.Controls.Add(this.lblUserName);
        this.grpLogin.Controls.Add(this.lblPassword);
        this.grpLogin.Controls.Add(this.txtPassword);
        this.grpLogin.Location = new System.Drawing.Point(6, 6);
        this.grpLogin.Name = "grpLogin";
        this.grpLogin.Size = new System.Drawing.Size(280, 73);
        this.grpLogin.TabIndex = 1;
        this.grpLogin.TabStop = false;
        // 
        // cmbType
        // 
        this.cmbType.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        this.cmbType.FormattingEnabled = true;
        this.cmbType.Items.AddRange(new object[] {
            "Credit",
            "Corporate",
            "Debit"});
        this.cmbType.Location = new System.Drawing.Point(99, 44);
        this.cmbType.Name = "cmbType";
        this.cmbType.Size = new System.Drawing.Size(172, 21);
        this.cmbType.TabIndex = 22;
        this.cmbType.SelectedIndexChanged += new System.EventHandler(this.cmbType_SelectedIndexChanged);
        // 
        // lblType
        // 
        this.lblType.BackColor = System.Drawing.Color.Transparent;
        this.lblType.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
        this.lblType.Location = new System.Drawing.Point(8, 46);
        this.lblType.Name = "lblType";
        this.lblType.Size = new System.Drawing.Size(92, 16);
        this.lblType.TabIndex = 21;
        this.lblType.Text = "Type:";
        // 
        // chkExternal
        // 
        this.chkExternal.BackColor = System.Drawing.Color.Transparent;
        this.chkExternal.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        this.chkExternal.ForeColor = System.Drawing.Color.DarkSlateGray;
        this.chkExternal.Location = new System.Drawing.Point(5, 83);
        this.chkExternal.Name = "chkExternal";
        this.chkExternal.Size = new System.Drawing.Size(133, 20);
        this.chkExternal.TabIndex = 20;
        this.chkExternal.Text = "External";
        this.chkExternal.UseVisualStyleBackColor = false;
        this.chkExternal.Visible = false;
        // 
        // cmbDatabase
        // 
        this.cmbDatabase.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        this.cmbDatabase.FormattingEnabled = true;
        this.cmbDatabase.Location = new System.Drawing.Point(99, 13);
        this.cmbDatabase.Name = "cmbDatabase";
        this.cmbDatabase.Size = new System.Drawing.Size(172, 21);
        this.cmbDatabase.TabIndex = 4;
        this.cmbDatabase.SelectedIndexChanged += new System.EventHandler(this.cmbDatabase_SelectedIndexChanged);
        // 
        // txtUserName
        // 
        this.txtUserName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
        this.txtUserName.Location = new System.Drawing.Point(99, 74);
        this.txtUserName.Name = "txtUserName";
        this.txtUserName.Size = new System.Drawing.Size(172, 22);
        this.txtUserName.TabIndex = 6;
        this.txtUserName.Visible = false;
        // 
        // lblDatabaseName
        // 
        this.lblDatabaseName.BackColor = System.Drawing.Color.Transparent;
        this.lblDatabaseName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
        this.lblDatabaseName.Location = new System.Drawing.Point(8, 15);
        this.lblDatabaseName.Name = "lblDatabaseName";
        this.lblDatabaseName.Size = new System.Drawing.Size(92, 16);
        this.lblDatabaseName.TabIndex = 3;
        this.lblDatabaseName.Text = "Database:";
        // 
        // lblUserName
        // 
        this.lblUserName.BackColor = System.Drawing.Color.Transparent;
        this.lblUserName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
        this.lblUserName.Location = new System.Drawing.Point(8, 72);
        this.lblUserName.Name = "lblUserName";
        this.lblUserName.Size = new System.Drawing.Size(92, 16);
        this.lblUserName.TabIndex = 5;
        this.lblUserName.Text = "User Name:";
        this.lblUserName.Visible = false;
        // 
        // lblPassword
        // 
        this.lblPassword.BackColor = System.Drawing.Color.Transparent;
        this.lblPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
        this.lblPassword.Location = new System.Drawing.Point(8, 73);
        this.lblPassword.Name = "lblPassword";
        this.lblPassword.Size = new System.Drawing.Size(92, 16);
        this.lblPassword.TabIndex = 7;
        this.lblPassword.Text = "Password:";
        this.lblPassword.Visible = false;
        // 
        // txtPassword
        // 
        this.txtPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
        this.txtPassword.Location = new System.Drawing.Point(99, 74);
        this.txtPassword.Name = "txtPassword";
        this.txtPassword.PasswordChar = '*';
        this.txtPassword.Size = new System.Drawing.Size(172, 22);
        this.txtPassword.TabIndex = 0;
        this.txtPassword.Visible = false;
        // 
        // lblVersion
        // 
        this.lblVersion.BackColor = System.Drawing.Color.Transparent;
        this.lblVersion.ForeColor = System.Drawing.SystemColors.HotTrack;
        this.lblVersion.Location = new System.Drawing.Point(170, 152);
        this.lblVersion.Name = "lblVersion";
        this.lblVersion.Size = new System.Drawing.Size(112, 20);
        this.lblVersion.TabIndex = 5;
        this.lblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
        // 
        // lblDatabase
        // 
        this.lblDatabase.BackColor = System.Drawing.Color.Transparent;
        this.lblDatabase.ForeColor = System.Drawing.SystemColors.HotTrack;
        this.lblDatabase.Location = new System.Drawing.Point(8, 0);
        this.lblDatabase.Name = "lblDatabase";
        this.lblDatabase.Size = new System.Drawing.Size(274, 16);
        this.lblDatabase.TabIndex = 0;
        this.lblDatabase.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
        // 
        // txtServerName
        // 
        this.txtServerName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
        this.txtServerName.Location = new System.Drawing.Point(228, 208);
        this.txtServerName.Name = "txtServerName";
        this.txtServerName.Size = new System.Drawing.Size(172, 22);
        this.txtServerName.TabIndex = 16;
        this.txtServerName.Visible = false;
        // 
        // frmBasLogin
        // 
        this.AcceptButton = this.btnOK;
        this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
        this.CancelButton = this.btnCancel;
        this.ClientSize = new System.Drawing.Size(311, 144);
        this.Controls.Add(this.lblVersion);
        this.Controls.Add(this.btnOK);
        this.Controls.Add(this.txtServerName);
        this.Controls.Add(this.btnCancel);
        this.Controls.Add(this.grpLogin);
        this.Controls.Add(this.lblDatabase);
        this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
        this.Name = "frmBasLogin";
        this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
        this.Text = "Database Login";
        this.Load += new System.EventHandler(this.frmLogin_Load);
        this.grpLogin.ResumeLayout(false);
        this.grpLogin.PerformLayout();
        this.ResumeLayout(false);
        this.PerformLayout();

        }
    #endregion



    private void btnCancel_Click(object sender, System.EventArgs e)
        {
        this.Close();
        //this.Dispose(true);
        }

    private void frmLogin_Load(object sender, System.EventArgs e)
        {
        //clsXmlReadWrite.readConfigration();//readINI();
        //clsDbCon.sPackage = clsCnfg.readSetting("Package");
        clsDbCon.sServer = clsCnfg.readSetting("ServerName");
        //clsDbCon.sUserName = clsCnfg.readSetting("UserName");
        //clsDbCon.sPassword = clsCnfg.readSetting("Password");
        //clsXmlReadWrite.writeConfigration();
        string[] strArray = clsDbCon.sPackage.Split(';');
        strArray = clsDbCon.sServer.Split(';');
        foreach (string str in strArray)
            cmbDatabase.Items.Add(str);
        cmbDatabase.SelectedIndex = 0;
        txtServerName.Text = clsDbCon.sServer;
        clsBasUserData.userName = Environment.UserName;
        clsBasUserData.sExternal = true;
        //if (SystemInformation.UserName == "mmohammed") // 
        //{
        txtUserName.Text = clsDbCon.sUserName;
        txtUserName.Tag = clsDbCon.sUserName;
        txtPassword.Text = clsDbCon.sPassword;
        txtPassword.Tag = clsDbCon.sPassword;
        //}
        //if (SystemInformation.ComputerName != "MMOHAMED") // 
        //{
        //this.Close();
        //}
        if (System.IO.File.Exists(Application.StartupPath + "\\frmBackground.jpg"))
            {
            this.BackgroundImage = Image.FromFile(Application.StartupPath + "\\frmBackground.jpg");
            grpLogin.BackgroundImage = Image.FromFile(Application.StartupPath + "\\frmBackground.jpg");
            }

        AssemblyInfo ainfo = new AssemblyInfo();
        lblVersion.Text = ainfo.Version;

        showToolTips();

        }

    private void btnOK_Click(object sender, System.EventArgs e)
        {
        // create OracleConnection object
        //clsDbCon.sPackage = cmbPackage.SelectedItem.ToString(); //   "tw1tst";
        clsDbCon.sServer = cmbDatabase.SelectedItem.ToString(); //   "tw1tst";
        clsDbCon.sUserName = txtUserName.Tag.ToString(); // "a4m";
        clsDbCon.sPassword = txtPassword.Text; //"a4m";

        //OracleConnection Conn = new OracleConnection("SERVER=tw1tst;" + "UID=a4m;" + "PASSWORD=a4m;");
        if (cmbType.SelectedIndex == -1)
            {
            MessageBox.Show("Please select type!", "Logon",
               MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
            }
        OracleConnection Conn = new OracleConnection(clsDbCon.sConOracle);
        try
            {
            //open connection
            Conn.Open();
            //open was successful   
            //MessageBox.Show("Connection Properties:\n" +
            //"Connection String: " + Conn.ConnectionString + "\n" +
            //"DataSource: " + Conn.DataSource + "\n" +
            //"ServerVersion: " + Conn.ServerVersion + "\n" +
            //"State:" + Conn.State
            //, "Oracle Connection Successfully Opened!",
            //MessageBoxButtons.OK, MessageBoxIcon.Information);
            MessageBox.Show("Oracle Connection Opened Successfully!", "Information",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            clsDbCon.IsRightCon = true;
            clsBasUserData.loginDate = DateTime.Now;
            clsBasUserData.ComputerIP = clsBasNetwork.getMyComputerIP();
            //clsBasLocalization.cultureName = "ar-EG";  //"ru-RU";
            //clsBasLocalization.setCulture();
            this.Dispose(false);  //false this.Close();
            //			new frmRptReportManager().Show();
            new frmStatementFile().Show();

            }
        catch (OracleException ex)
            {
            clsBasErrors.catchError(ex);
            clsDbCon.IsRightCon = false;
            //Application.Exit();
            }
        catch (Exception ex)   //DivideByZeroException ex)
            {
            clsBasErrors.catchError(ex);
            }
        finally
            {
            // close connection
            Conn.Close();
            }
        }

    private void readINI()
        {
        //clsAppConfig config = new clsAppConfig();
        //string str = (string)( config.GetValue( "DbName", typeof( string ) ) );
        //bool 	 bln 	= (bool)( config.GetValue( "Boolean", typeof( bool ) ) );
        //DateTime date	= (DateTime)( config.GetValue( "DateTime", typeof( DateTime ) ) ); 	
        //txtServerName.Text = str;

        IniReader ini = new IniReader(@Environment.CurrentDirectory + @"\DataManagers.INI");//D:\pC#\exe\
        txtServerName.Text = ini.ReadString("Settings", "ServerName", ""); //"tw1tst";
        txtUserName.Text = ini.ReadString("Settings", "UserName", ""); //"a4m";
        txtPassword.Text = ini.ReadString("Settings", "Password", ""); //"a4m";
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
        frmToolTip.SetToolTip(this.btnCancel, "Cancel Login and Exit The Program");
        frmToolTip.SetToolTip(this.btnOK, "Login to Program");
        }

    private void cmbDatabase_SelectedIndexChanged(object sender, EventArgs e)
        {
        //if ((cmbDatabase.SelectedIndex == 0))
        //{
        //if (chkExternal.Checked == true)
        //    {
        txtUserName.Tag = "/";
        txtPassword.Tag = "";
        //    }
        //else if (chkExternal.Checked == false)
        //    {
        //    txtUserName.Enabled = true;
        //    clsDbCon.sUserName = clsCnfg.readSetting("UserName");
        //    clsDbCon.sPassword = clsCnfg.readSetting("Password");
        //    txtUserName.Text = clsDbCon.sUserName;
        //    txtUserName.Tag = clsDbCon.sUserName;
        //    txtPassword.Text = clsDbCon.sPassword;
        //    txtPassword.Tag = clsDbCon.sPassword;
        //    }
        //    }
        //else if (cmbDatabase.SelectedIndex == 1)
        //    {
        //if (chkExternal.Checked == true)
        //    {
        //txtUserName.Tag = "/";
        //txtPassword.Tag = "";
        //}
        //else if (chkExternal.Checked == false)
        //    {
        //    txtUserName.Enabled = true;
        //    clsDbCon.sUserName = clsCnfg.readSetting("TestUserName1");
        //    clsDbCon.sPassword = clsCnfg.readSetting("TestPassword1");
        //    txtUserName.Text = clsDbCon.sUserName;
        //    txtUserName.Tag = clsDbCon.sUserName;
        //    txtPassword.Text = clsDbCon.sPassword;
        //    txtPassword.Tag = clsDbCon.sPassword;
        //    }
        //    }
        //else if (cmbDatabase.SelectedIndex == 2)
        //    {
        //if (chkExternal.Checked == true)
        //    {
        //txtUserName.Tag = "/";
        //txtPassword.Tag = "";
        //}
        //else if (chkExternal.Checked == false)
        //    {
        //    txtUserName.Enabled = true;
        //    clsDbCon.sUserName = clsCnfg.readSetting("TestUserName2");
        //    clsDbCon.sPassword = clsCnfg.readSetting("TestPassword2");
        //    txtUserName.Text = clsDbCon.sUserName;
        //    txtUserName.Tag = clsDbCon.sUserName;
        //    txtPassword.Text = clsDbCon.sPassword;
        //    txtPassword.Tag = clsDbCon.sPassword;
        //    }
        //}
        //else if (cmbDatabase.SelectedIndex == 3)
        //{
        //if (chkExternal.Checked == true)
        //{
        //txtUserName.Tag = "/";
        //txtPassword.Tag = "";
        //    }
        //else if (chkExternal.Checked == false)
        //    {
        //    txtUserName.Enabled = true;
        //    clsDbCon.sUserName = clsCnfg.readSetting("TestUserName3");
        //    clsDbCon.sPassword = clsCnfg.readSetting("TestPassword3");
        //    txtUserName.Text = clsDbCon.sUserName;
        //    txtUserName.Tag = clsDbCon.sUserName;
        //    txtPassword.Text = clsDbCon.sPassword;
        //    txtPassword.Tag = clsDbCon.sPassword;
        //    }
        //}
        }

    //private void chkExternal_CheckedChanged(object sender, EventArgs e)
    //    {
    //    if (chkExternal.Checked == true)
    //        {
    //        txtUserName.Text = "";
    //        txtUserName.Tag = "/";
    //        txtUserName.Enabled = false;
    //        txtPassword.Text = "";
    //        txtPassword.Enabled = false;
    //        clsBasUserData.sExternal = true;
    //        }
    //    else if (chkExternal.Checked == false)
    //        {
    //        txtUserName.Text = clsDbCon.sUserName;
    //        txtUserName.Tag = clsDbCon.sUserName;
    //        txtUserName.Enabled = true;
    //        txtPassword.Text = clsDbCon.sPassword;
    //        txtPassword.Enabled = true;
    //        clsBasUserData.sExternal = false;
    //        }
    //    }

    private void cmbType_SelectedIndexChanged(object sender, EventArgs e)
        {
        if (cmbType.SelectedIndex == 0)
            clsBasUserData.sType = "CR";
        else if (cmbType.SelectedIndex == 1)
            clsBasUserData.sType = "CP";
        else if (cmbType.SelectedIndex == 2)
            clsBasUserData.sType = "DB";
        }
    }

using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

//
public class frmValidation : System.Windows.Forms.Form
{
  public System.Windows.Forms.ProgressBar progBarStatus;
  private System.Windows.Forms.Label lblStatus;
  private System.Windows.Forms.Button btnSaveStatement;
  private System.Windows.Forms.Button btnSaveFileName;
  private System.Windows.Forms.TextBox txtFileName;
  private System.Windows.Forms.Button btnExit;
  private System.Windows.Forms.Label lblStatementPath;
  private System.Windows.Forms.Label lblPageSize;
  private System.Windows.Forms.TextBox txtPageSize;
  private System.Windows.Forms.Label lblCharacter;
  private System.Windows.Forms.TextBox txtCharacter;
  /// <summary>
  /// Required designer variable.
  /// </summary>
  private System.ComponentModel.Container components = null;

  public frmValidation()
  {
    //
    // Required for Windows Form Designer support
    //
    InitializeComponent();

    //
    // TODO: Add any constructor code after InitializeComponent call
    //
  }

  /// <summary>
  /// Clean up any resources being used.
  /// </summary>
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
      System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmValidation));
      this.progBarStatus = new System.Windows.Forms.ProgressBar();
      this.lblStatus = new System.Windows.Forms.Label();
      this.btnSaveStatement = new System.Windows.Forms.Button();
      this.btnSaveFileName = new System.Windows.Forms.Button();
      this.txtFileName = new System.Windows.Forms.TextBox();
      this.btnExit = new System.Windows.Forms.Button();
      this.lblStatementPath = new System.Windows.Forms.Label();
      this.txtPageSize = new System.Windows.Forms.TextBox();
      this.lblPageSize = new System.Windows.Forms.Label();
      this.txtCharacter = new System.Windows.Forms.TextBox();
      this.lblCharacter = new System.Windows.Forms.Label();
      this.SuspendLayout();
      // 
      // progBarStatus
      // 
      this.progBarStatus.Location = new System.Drawing.Point(8, 120);
      this.progBarStatus.Name = "progBarStatus";
      this.progBarStatus.Size = new System.Drawing.Size(424, 20);
      this.progBarStatus.TabIndex = 28;
      // 
      // lblStatus
      // 
      this.lblStatus.BackColor = System.Drawing.Color.Transparent;
      this.lblStatus.ForeColor = System.Drawing.Color.Brown;
      this.lblStatus.Location = new System.Drawing.Point(96, 72);
      this.lblStatus.Name = "lblStatus";
      this.lblStatus.Size = new System.Drawing.Size(472, 40);
      this.lblStatus.TabIndex = 27;
      // 
      // btnSaveStatement
      // 
      this.btnSaveStatement.Image = ((System.Drawing.Image)(resources.GetObject("btnSaveStatement.Image")));
      this.btnSaveStatement.Location = new System.Drawing.Point(52, 84);
      this.btnSaveStatement.Name = "btnSaveStatement";
      this.btnSaveStatement.Size = new System.Drawing.Size(32, 24);
      this.btnSaveStatement.TabIndex = 26;
      this.btnSaveStatement.Click += new System.EventHandler(this.btnSaveStatement_Click);
      // 
      // btnSaveFileName
      // 
      this.btnSaveFileName.Image = ((System.Drawing.Image)(resources.GetObject("btnSaveFileName.Image")));
      this.btnSaveFileName.Location = new System.Drawing.Point(308, 48);
      this.btnSaveFileName.Name = "btnSaveFileName";
      this.btnSaveFileName.Size = new System.Drawing.Size(28, 20);
      this.btnSaveFileName.TabIndex = 25;
      this.btnSaveFileName.Click += new System.EventHandler(this.btnSaveFileName_Click);
      // 
      // txtFileName
      // 
      this.txtFileName.Location = new System.Drawing.Point(12, 48);
      this.txtFileName.Name = "txtFileName";
      this.txtFileName.Size = new System.Drawing.Size(296, 20);
      this.txtFileName.TabIndex = 24;
      this.txtFileName.Text = "C:\\TEMP\\P20Files\\Statement\\";
      // 
      // btnExit
      // 
      this.btnExit.Image = ((System.Drawing.Image)(resources.GetObject("btnExit.Image")));
      this.btnExit.Location = new System.Drawing.Point(12, 84);
      this.btnExit.Name = "btnExit";
      this.btnExit.Size = new System.Drawing.Size(32, 24);
      this.btnExit.TabIndex = 23;
      this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
      // 
      // lblStatementPath
      // 
      this.lblStatementPath.BackColor = System.Drawing.Color.Transparent;
      this.lblStatementPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
      this.lblStatementPath.Location = new System.Drawing.Point(12, 24);
      this.lblStatementPath.Name = "lblStatementPath";
      this.lblStatementPath.Size = new System.Drawing.Size(232, 20);
      this.lblStatementPath.TabIndex = 22;
      this.lblStatementPath.Text = "StatementPath";
      // 
      // txtPageSize
      // 
      this.txtPageSize.Location = new System.Drawing.Point(364, 48);
      this.txtPageSize.Name = "txtPageSize";
      this.txtPageSize.Size = new System.Drawing.Size(44, 20);
      this.txtPageSize.TabIndex = 30;
      this.txtPageSize.Text = "48";
      // 
      // lblPageSize
      // 
      this.lblPageSize.BackColor = System.Drawing.Color.Transparent;
      this.lblPageSize.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
      this.lblPageSize.Location = new System.Drawing.Point(364, 24);
      this.lblPageSize.Name = "lblPageSize";
      this.lblPageSize.Size = new System.Drawing.Size(88, 20);
      this.lblPageSize.TabIndex = 29;
      this.lblPageSize.Text = "Page Size";
      // 
      // txtCharacter
      // 
      this.txtCharacter.Location = new System.Drawing.Point(468, 48);
      this.txtCharacter.Name = "txtCharacter";
      this.txtCharacter.Size = new System.Drawing.Size(84, 20);
      this.txtCharacter.TabIndex = 30;
      // 
      // lblCharacter
      // 
      this.lblCharacter.BackColor = System.Drawing.Color.Transparent;
      this.lblCharacter.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
      this.lblCharacter.Location = new System.Drawing.Point(468, 24);
      this.lblCharacter.Name = "lblCharacter";
      this.lblCharacter.Size = new System.Drawing.Size(88, 20);
      this.lblCharacter.TabIndex = 29;
      this.lblCharacter.Text = "Character";
      // 
      // frmValidation
      // 
      this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
      this.ClientSize = new System.Drawing.Size(572, 153);
      this.Controls.Add(this.txtPageSize);
      this.Controls.Add(this.txtFileName);
      this.Controls.Add(this.txtCharacter);
      this.Controls.Add(this.lblPageSize);
      this.Controls.Add(this.progBarStatus);
      this.Controls.Add(this.lblStatus);
      this.Controls.Add(this.btnSaveStatement);
      this.Controls.Add(this.btnSaveFileName);
      this.Controls.Add(this.btnExit);
      this.Controls.Add(this.lblStatementPath);
      this.Controls.Add(this.lblCharacter);
      this.Name = "frmValidation";
      this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
      this.Text = "Validation";
      this.Load += new System.EventHandler(this.frmValidation_Load);
      this.ResumeLayout(false);
      this.PerformLayout();

  }
  #endregion

  private void btnSaveFileName_Click(object sender, System.EventArgs e)
  {
    txtFileName.Text = clsBasFile.openFileDialog("All Files (*.*)|*.*");
  }

  private void btnExit_Click(object sender, System.EventArgs e)
  {
    this.Close();
  }

  private void btnSaveStatement_Click(object sender, System.EventArgs e)
  {
    clsValidatePageSize ValidatePageSize = new clsValidatePageSize();
    lblStatus.Text = ValidatePageSize.outMessage + ValidatePageSize.ValidatePageSize(txtFileName.Text, Convert.ToInt32(txtPageSize.Text), txtCharacter.Text).ToString() + " Errors";
  }

  private void frmValidation_Load(object sender, System.EventArgs e)
  {
    txtCharacter.Text = "\u000C";
  }
}

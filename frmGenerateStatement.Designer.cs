
partial class frmGenerateStatement
{
  /// <summary>
  /// Required designer variable.
  /// </summary>
  private System.ComponentModel.IContainer components = null;

  /// <summary>
  /// Clean up any resources being used.
  /// </summary>
  /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
  protected override void Dispose(bool disposing)
  {
    if (disposing && (components != null))
    {
      components.Dispose();
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
    System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmGenerateStatement));
    this.txtCurProgress = new System.Windows.Forms.TextBox();
    this.txtTotProgress = new System.Windows.Forms.TextBox();
    this.txtRunResult = new System.Windows.Forms.TextBox();
    this.progBarStatus = new System.Windows.Forms.ProgressBar();
    this.lblStatus = new System.Windows.Forms.Label();
    this.btnExit = new System.Windows.Forms.Button();
    this.btnCancel = new System.Windows.Forms.Button();
    this.SuspendLayout();
    // 
    // txtCurProgress
    // 
    this.txtCurProgress.Font = new System.Drawing.Font("MS Reference Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
    this.txtCurProgress.Location = new System.Drawing.Point(3, 126);
    this.txtCurProgress.Name = "txtCurProgress";
    this.txtCurProgress.Size = new System.Drawing.Size(93, 18);
    this.txtCurProgress.TabIndex = 69;
    // 
    // txtTotProgress
    // 
    this.txtTotProgress.Font = new System.Drawing.Font("MS Reference Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
    this.txtTotProgress.Location = new System.Drawing.Point(399, 126);
    this.txtTotProgress.Name = "txtTotProgress";
    this.txtTotProgress.Size = new System.Drawing.Size(75, 18);
    this.txtTotProgress.TabIndex = 68;
    // 
    // txtRunResult
    // 
    this.txtRunResult.Location = new System.Drawing.Point(3, 38);
    this.txtRunResult.Multiline = true;
    this.txtRunResult.Name = "txtRunResult";
    this.txtRunResult.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
    this.txtRunResult.Size = new System.Drawing.Size(470, 85);
    this.txtRunResult.TabIndex = 67;
    // 
    // progBarStatus
    // 
    this.progBarStatus.Location = new System.Drawing.Point(99, 126);
    this.progBarStatus.Name = "progBarStatus";
    this.progBarStatus.Size = new System.Drawing.Size(297, 17);
    this.progBarStatus.Step = 0;
    this.progBarStatus.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
    this.progBarStatus.TabIndex = 66;
    // 
    // lblStatus
    // 
    this.lblStatus.BackColor = System.Drawing.Color.Transparent;
    this.lblStatus.ForeColor = System.Drawing.Color.Brown;
    this.lblStatus.Location = new System.Drawing.Point(4, 2);
    this.lblStatus.Name = "lblStatus";
    this.lblStatus.Size = new System.Drawing.Size(468, 32);
    this.lblStatus.TabIndex = 65;
    this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
    // 
    // btnExit
    // 
    this.btnExit.Image = ((System.Drawing.Image)(resources.GetObject("btnExit.Image")));
    this.btnExit.Location = new System.Drawing.Point(4, 149);
    this.btnExit.Name = "btnExit";
    this.btnExit.Size = new System.Drawing.Size(32, 21);
    this.btnExit.TabIndex = 64;
    this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
    // 
    // btnCancel
    // 
    this.btnCancel.Location = new System.Drawing.Point(42, 149);
    this.btnCancel.Name = "btnCancel";
    this.btnCancel.Size = new System.Drawing.Size(32, 21);
    this.btnCancel.TabIndex = 70;
    // 
    // frmGenerateSatatement
    // 
    this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
    this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
    this.ClientSize = new System.Drawing.Size(476, 173);
    this.Controls.Add(this.btnCancel);
    this.Controls.Add(this.txtCurProgress);
    this.Controls.Add(this.txtTotProgress);
    this.Controls.Add(this.txtRunResult);
    this.Controls.Add(this.progBarStatus);
    this.Controls.Add(this.lblStatus);
    this.Controls.Add(this.btnExit);
    this.Name = "frmGenerateSatatement";
    this.Text = "Generate Satatement";
    this.Load += new System.EventHandler(this.frmGenerateSatatement_Load);
    this.ResumeLayout(false);
    this.PerformLayout();

  }

  #endregion

  private System.Windows.Forms.TextBox txtCurProgress;
  private System.Windows.Forms.TextBox txtTotProgress;
  private System.Windows.Forms.TextBox txtRunResult;
  private System.Windows.Forms.ProgressBar progBarStatus;
  private System.Windows.Forms.Label lblStatus;
  private System.Windows.Forms.Button btnExit;
  private System.Windows.Forms.Button btnCancel;
}

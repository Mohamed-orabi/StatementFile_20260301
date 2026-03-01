namespace StatementFile.StatementFile
    {
    partial class StatementConfig
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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.MailContValue = new System.Windows.Forms.NumericUpDown();
            this.NumCheckBox = new System.Windows.Forms.CheckBox();
            this.AccCheckBox = new System.Windows.Forms.CheckBox();
            this.lblAccNo = new System.Windows.Forms.Label();
            this.txtAccNo = new System.Windows.Forms.TextBox();
            this.groupBoxEmails = new System.Windows.Forms.GroupBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblNoStat = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.MailContValue)).BeginInit();
            this.groupBoxEmails.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.textBox1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.textBox1.Location = new System.Drawing.Point(99, 79);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(256, 20);
            this.textBox1.TabIndex = 0;
            // 
            // textBox2
            // 
            this.textBox2.AutoCompleteCustomSource.AddRange(new string[] {
            "mtaher@emp-group.com",
            "iatta@emp-group.com",
            "mabouleila@emp-group.com"});
            this.textBox2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.textBox2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.textBox2.Location = new System.Drawing.Point(99, 105);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(256, 20);
            this.textBox2.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(13, 82);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(33, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "To   :";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
            this.label2.Location = new System.Drawing.Point(16, 111);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(33, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "CC   :";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
            this.button1.Location = new System.Drawing.Point(104, 333);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "Submit";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(99, 134);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(256, 20);
            this.textBox3.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
            this.label3.Location = new System.Drawing.Point(16, 137);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(34, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "BCC :";
            // 
            // button2
            // 
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
            this.button2.Location = new System.Drawing.Point(185, 333);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 7;
            this.button2.Text = "Cancel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // MailContValue
            // 
            this.MailContValue.Location = new System.Drawing.Point(161, 90);
            this.MailContValue.Maximum = new decimal(new int[] {
            4000,
            0,
            0,
            0});
            this.MailContValue.Name = "MailContValue";
            this.MailContValue.Size = new System.Drawing.Size(53, 20);
            this.MailContValue.TabIndex = 9;
            this.MailContValue.Visible = false;
            // 
            // NumCheckBox
            // 
            this.NumCheckBox.AutoSize = true;
            this.NumCheckBox.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
            this.NumCheckBox.Location = new System.Drawing.Point(13, 68);
            this.NumCheckBox.Name = "NumCheckBox";
            this.NumCheckBox.Size = new System.Drawing.Size(183, 17);
            this.NumCheckBox.TabIndex = 11;
            this.NumCheckBox.Text = "Send  Number Of Statement";
            this.NumCheckBox.UseVisualStyleBackColor = true;
            this.NumCheckBox.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // AccCheckBox
            // 
            this.AccCheckBox.AutoSize = true;
            this.AccCheckBox.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
            this.AccCheckBox.Location = new System.Drawing.Point(13, 19);
            this.AccCheckBox.Name = "AccCheckBox";
            this.AccCheckBox.Size = new System.Drawing.Size(165, 17);
            this.AccCheckBox.TabIndex = 12;
            this.AccCheckBox.Text = "Send to specific account ";
            this.AccCheckBox.UseVisualStyleBackColor = true;
            this.AccCheckBox.CheckedChanged += new System.EventHandler(this.checkBox2_CheckedChanged);
            // 
            // lblAccNo
            // 
            this.lblAccNo.AutoSize = true;
            this.lblAccNo.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
            this.lblAccNo.Location = new System.Drawing.Point(10, 45);
            this.lblAccNo.Name = "lblAccNo";
            this.lblAccNo.Size = new System.Drawing.Size(106, 13);
            this.lblAccNo.TabIndex = 13;
            this.lblAccNo.Text = "Account Number :";
            this.lblAccNo.Visible = false;
            // 
            // txtAccNo
            // 
            this.txtAccNo.Location = new System.Drawing.Point(122, 45);
            this.txtAccNo.Name = "txtAccNo";
            this.txtAccNo.Size = new System.Drawing.Size(233, 20);
            this.txtAccNo.TabIndex = 14;
            this.txtAccNo.Visible = false;
            // 
            // groupBoxEmails
            // 
            this.groupBoxEmails.Controls.Add(this.textBox5);
            this.groupBoxEmails.Controls.Add(this.label5);
            this.groupBoxEmails.Controls.Add(this.textBox4);
            this.groupBoxEmails.Controls.Add(this.label4);
            this.groupBoxEmails.Controls.Add(this.textBox1);
            this.groupBoxEmails.Controls.Add(this.label1);
            this.groupBoxEmails.Controls.Add(this.textBox2);
            this.groupBoxEmails.Controls.Add(this.label2);
            this.groupBoxEmails.Controls.Add(this.textBox3);
            this.groupBoxEmails.Controls.Add(this.label3);
            this.groupBoxEmails.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
            this.groupBoxEmails.Location = new System.Drawing.Point(12, 12);
            this.groupBoxEmails.Name = "groupBoxEmails";
            this.groupBoxEmails.Size = new System.Drawing.Size(389, 174);
            this.groupBoxEmails.TabIndex = 15;
            this.groupBoxEmails.TabStop = false;
            this.groupBoxEmails.Text = "Internal Emails";
            // 
            // textBox5
            // 
            this.textBox5.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.textBox5.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.textBox5.Location = new System.Drawing.Point(99, 54);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(256, 20);
            this.textBox5.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
            this.label5.Location = new System.Drawing.Point(13, 57);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(80, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "From Name  :";
            // 
            // textBox4
            // 
            this.textBox4.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.textBox4.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.textBox4.Location = new System.Drawing.Point(99, 29);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(256, 20);
            this.textBox4.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
            this.label4.Location = new System.Drawing.Point(13, 32);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(48, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "From   :";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblNoStat);
            this.groupBox1.Controls.Add(this.txtAccNo);
            this.groupBox1.Controls.Add(this.lblAccNo);
            this.groupBox1.Controls.Add(this.MailContValue);
            this.groupBox1.Controls.Add(this.NumCheckBox);
            this.groupBox1.Controls.Add(this.AccCheckBox);
            this.groupBox1.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
            this.groupBox1.Location = new System.Drawing.Point(12, 192);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(389, 127);
            this.groupBox1.TabIndex = 16;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Filters";
            // 
            // lblNoStat
            // 
            this.lblNoStat.AutoSize = true;
            this.lblNoStat.Location = new System.Drawing.Point(13, 97);
            this.lblNoStat.Name = "lblNoStat";
            this.lblNoStat.Size = new System.Drawing.Size(142, 13);
            this.lblNoStat.TabIndex = 15;
            this.lblNoStat.Text = "Number Of Statements :";
            this.lblNoStat.Visible = false;
            // 
            // StatementConfig
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.button2;
            this.ClientSize = new System.Drawing.Size(417, 369);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBoxEmails);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "StatementConfig";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "StatementConfig";
            ((System.ComponentModel.ISupportInitialize)(this.MailContValue)).EndInit();
            this.groupBoxEmails.ResumeLayout(false);
            this.groupBoxEmails.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

            }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.NumericUpDown MailContValue;
        private System.Windows.Forms.CheckBox NumCheckBox;
        private System.Windows.Forms.CheckBox AccCheckBox;
        private System.Windows.Forms.Label lblAccNo;
        private System.Windows.Forms.TextBox txtAccNo;
        private System.Windows.Forms.GroupBox groupBoxEmails;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lblNoStat;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.Label label5;
        }
    }
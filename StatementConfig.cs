using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StatementFile.StatementFile
    {
    public partial class StatementConfig : Form
        {
        public StatementConfig()
            {
            InitializeComponent();
            //textBox1.Text = "mlajin@network.global";

            }
        private void textBox1_Click(object sender, EventArgs e)
            {
            int nextSpaceIndex = textBox1.Text.Substring(textBox1.SelectionStart).IndexOf(' ');
            int firstSpaceIndex = textBox1.Text.Substring(0, textBox1.SelectionStart).LastIndexOf(' ');
            nextSpaceIndex = nextSpaceIndex == -1 ? textBox1.Text.Length : nextSpaceIndex + textBox1.SelectionStart;
            firstSpaceIndex = firstSpaceIndex == -1 ? 0 : firstSpaceIndex;
            textBox1.SelectionStart = firstSpaceIndex;
            textBox1.SelectionLength = nextSpaceIndex - firstSpaceIndex;
            }
        public static bool isEmail(string inputEmail)
        {
            //string strRegex = @"^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}" +
            //    @"\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\" +
            //    @".)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$";
            //Regex re = new Regex(strRegex);
            //if (re.IsMatch(inputEmail))
            //    return (true);
            //else
            //    return (false);
            RegexUtilities util = new RegexUtilities();
            if (util.IsValidEmail(inputEmail))
                return (true);
            else
                return (false);
        }

        private void button1_Click(object sender, EventArgs e)
            {
            if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text == "")
                {
                MessageBox.Show("Please Enter at least one Email ! ");
                }
            else
                {

                if (MessageBox.Show("Are you sure that you want to send statement only for this emails ?", "Submit", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                    if (textBox1.Text != "")
                        {
                        string To = textBox1.Text;
                        string[] mailto = To.Split(';');
                        foreach (string item in mailto)
                            {
                            if (item != "" && !isEmail(item))
                                {
                                MessageBox.Show(item + " Not Valid Email");
                                return;
                                }
                            frmStatementFile.InternalMailTo.Add(item);
                            }

                        }
                    if (textBox2.Text != "")
                        {
                        string CC = textBox2.Text;
                        string[] mailcc = CC.Split(';');
                        foreach (string item in mailcc)
                            {
                            if (item != null && !isEmail(item))
                                {
                                MessageBox.Show(item + " Not Valid Email");
                                return;
                                }
                            frmStatementFile.InternalMailCC.Add(item);
                            }
                        }
                    if (textBox3.Text != "")
                        {
                        string BCC = textBox3.Text;
                        string[] mailBcc = BCC.Split(';');
                        foreach (string item in mailBcc)
                            {
                            if (item != "" && !isEmail(item))
                                {
                                MessageBox.Show(item + " Not Valid Email");
                                return;
                                }
                            frmStatementFile.InternalMailBCC.Add(item);
                            }
                        }

                    if (textBox4.Text != "")
                        {
                        frmStatementFile.InternalMailFrom = textBox4.Text;
                        }

                    if (textBox5.Text != "")
                        {
                        frmStatementFile.InternalMailFromName = textBox5.Text;
                        }


                    (this.Owner as frmStatementFile).button1.BackColor = Color.Green;
                    (this.Owner as frmStatementFile).button1.Text = "Internal Mode";
                    frmStatementFile.Internal = true;
                    if (NumCheckBox.Checked == true)
                        {
                        frmStatementFile.MailCount = Convert.ToInt32(MailContValue.Value);
                        }
                    else
                        {
                        frmStatementFile.MailCount = 1;
                        }
                    frmStatementFile.internalAccNo = txtAccNo.Text;


                    this.Close();
                    }

                }


            }

        private void button2_Click(object sender, EventArgs e)
            {
            this.Close();
            }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
            {
            if (NumCheckBox.Checked == true)
                {
                MailContValue.Visible = true;
                lblNoStat.Visible = true;
                lblAccNo.Visible = false;
                txtAccNo.Visible = false;
                AccCheckBox.Checked = false;
                txtAccNo.Text = "";
                }
            else
                {
                MailContValue.Visible = false;
                lblNoStat.Visible = false;
                }
            }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
            {
            if (AccCheckBox.Checked == true)
                {
                lblAccNo.Visible = true;
                txtAccNo.Visible = true;
                NumCheckBox.Checked = false;
                MailContValue.Visible = false;
                lblNoStat.Visible = false;
                }
            else
                {
                lblAccNo.Visible = false;
                txtAccNo.Visible = false;
                txtAccNo.Text = "";
                }
            }
        }
    }

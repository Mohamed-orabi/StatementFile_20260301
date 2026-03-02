using System;
using System.Windows.Forms;
using StatementFile.Infrastructure.Configuration;

namespace StatementFile.Presentation.Forms
{
    /// <summary>
    /// Login form — refactored to receive the composition root via constructor injection
    /// instead of directly constructing infrastructure objects.
    ///
    /// Business behaviour preserved: validates Oracle credentials before granting access
    /// to the main statement form.
    /// </summary>
    public partial class frmLogin : Form
    {
        private readonly CompositionRoot _root;

        public frmLogin(CompositionRoot root)
        {
            _root = root ?? throw new ArgumentNullException(nameof(root));
            InitializeComponent();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            try
            {
                // Oracle credentials entered by the user are passed to the
                // connection factory — no direct Oracle API calls in the UI layer.
                using (var uow = _root.CreateUnitOfWork())
                {
                    // A successful CreateUnitOfWork() call proves the connection string is valid.
                    // (Legacy: frmBasLogin opened a raw OracleConnection here.)
                }

                Hide();
                var mainForm = new frmGenerateStatement(_root);
                mainForm.FormClosed += (s, a) => Application.Exit();
                mainForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Login failed:\n{ex.Message}",
                    "StatementFile — Login",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        // Designer-generated code lives in frmLogin.Designer.cs (existing file).
        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.Text          = "StatementFile — Login";
            this.Size          = new System.Drawing.Size(360, 220);
            this.StartPosition = FormStartPosition.CenterScreen;

            var lblUser = new Label { Text = "Username:", Left = 20, Top = 30, Width = 80 };
            var txtUser = new TextBox { Left = 110, Top = 27, Width = 200, Name = "txtUser" };

            var lblPass = new Label { Text = "Password:", Left = 20, Top = 65, Width = 80 };
            var txtPass = new TextBox { Left = 110, Top = 62, Width = 200, PasswordChar = '*', Name = "txtPass" };

            var btnLogin  = new Button { Text = "Login",  Left = 110, Top = 100, Width = 90 };
            var btnCancel = new Button { Text = "Cancel", Left = 210, Top = 100, Width = 90 };

            btnLogin.Click  += btnLogin_Click;
            btnCancel.Click += btnCancel_Click;

            Controls.AddRange(new Control[] { lblUser, txtUser, lblPass, txtPass, btnLogin, btnCancel });
            this.ResumeLayout(false);
        }
    }
}

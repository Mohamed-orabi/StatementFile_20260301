using System;
using System.Windows.Forms;
using StatementFile.Infrastructure.Configuration;
using StatementFile.Presentation.Forms;

namespace StatementFile.Presentation
{
    /// <summary>
    /// Application entry point.
    /// Constructs the DI composition root once, then passes it into the
    /// login form so every subsequent form receives its dependencies
    /// through constructor injection — no static service locators.
    /// </summary>
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                // Build the full dependency graph at startup
                CompositionRoot root = DependencyInjection.Compose();

                Application.Run(new frmLogin(root));
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Fatal startup error:\n{ex.Message}",
                    "StatementFile",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}

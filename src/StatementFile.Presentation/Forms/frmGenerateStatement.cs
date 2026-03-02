using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using StatementFile.Application.UseCases.BulkProcessing;
using StatementFile.Application.UseCases.MerchantOnboarding;
using StatementFile.Application.UseCases.StatementGeneration;
using StatementFile.Domain.Enums;
using StatementFile.Infrastructure.Configuration;

namespace StatementFile.Presentation.Forms
{
    /// <summary>
    /// Main statement generation form.
    ///
    /// Refactored from frmStatementFile / frmStatementFileExtn:
    ///  - All business logic delegated to Application-layer Use Case Handlers.
    ///  - UI only handles user input, progress reporting, and error display.
    ///  - BackgroundWorker pattern preserved for UI responsiveness.
    ///  - No direct Oracle or file-system calls in this class.
    /// </summary>
    public partial class frmGenerateStatement : Form
    {
        private readonly CompositionRoot   _root;
        private          BackgroundWorker  _worker;

        // Selected products bound to the CheckedListBox
        private readonly List<ProductSelection> _selectedProducts = new List<ProductSelection>();

        public frmGenerateStatement(CompositionRoot root)
        {
            _root = root ?? throw new ArgumentNullException(nameof(root));
            InitializeComponent();
        }

        // ── UI Events ──────────────────────────────────────────────────────────────

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            if (_selectedProducts.Count == 0)
            {
                MessageBox.Show("Please select at least one product.", "StatementFile",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            btnGenerate.Enabled = false;
            progressBar.Value   = 0;
            progressBar.Maximum = _selectedProducts.Count;
            lblStatus.Text      = "Running…";

            _worker = new BackgroundWorker
            {
                WorkerReportsProgress      = true,
                WorkerSupportsCancellation = true,
            };
            _worker.DoWork             += Worker_DoWork;
            _worker.ProgressChanged    += Worker_ProgressChanged;
            _worker.RunWorkerCompleted += Worker_Completed;
            _worker.RunWorkerAsync(_selectedProducts);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            _worker?.CancelAsync();
        }

        private void btnMerchant_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog
            {
                Filter = "XML Files|*.xml",
                Title  = "Select Merchant XML File",
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;

                var cmd = new ProcessMerchantStatementCommand(
                    xmlSourceFilePath: dlg.FileName,
                    bankFullName:      txtBankFullName.Text.Trim(),
                    bankName:          txtBankName.Text.Trim(),
                    bankCode:          txtBankCode.Text.Trim(),
                    processingDate:    DateTime.Now);

                var result = _root.MerchantHandler.Handle(cmd);

                MessageBox.Show(
                    result.IsSuccess
                        ? "Merchant statement processed successfully."
                        : $"Error: {result.Error}",
                    "Merchant Statement",
                    MessageBoxButtons.OK,
                    result.IsSuccess ? MessageBoxIcon.Information : MessageBoxIcon.Error);
            }
        }

        // ── BackgroundWorker ───────────────────────────────────────────────────────

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            var products = (List<ProductSelection>)e.Argument;
            var worker   = (BackgroundWorker)sender;
            int index    = 0;

            foreach (var product in products)
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }

                worker.ReportProgress(index, $"Processing {product.BankName} / {product.CardProduct}…");

                // ── Step 1: Bulk maintenance ─────────────────────────────────────
                var bulkCmd = new RunBulkMaintenanceCommand(
                    branchCode:         product.BranchCode,
                    excludeReward:      product.ExcludeReward,
                    excludeInstallment: product.ExcludeInstallment);

                var bulkResult = _root.BulkHandler.Handle(bulkCmd);
                if (bulkResult.IsFailure)
                    worker.ReportProgress(index, $"Maintenance warning: {bulkResult.Error}");

                // ── Step 2: Statement generation ─────────────────────────────────
                var genCmd = new GenerateStatementCommand(
                    branchCode:    product.BranchCode,
                    bankName:      product.BankName,
                    bankFullName:  product.BankFullName,
                    cardProduct:   product.CardProduct,
                    cardType:      product.CardType,
                    statementType: StatementType.Normal,
                    formatterKey:  product.FormatterKey,
                    statementDate: DateTime.Now,
                    outputRootPath: _root.ConfigService.GetStatementOutputPath());

                var handler = _root.CreateStatementHandler();
                var genResult = handler.Handle(genCmd);

                if (genResult.IsFailure)
                    worker.ReportProgress(index, $"Generation error: {genResult.Error}");

                index++;
                worker.ReportProgress(index);
            }
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = Math.Min(e.ProgressPercentage, progressBar.Maximum);
            if (e.UserState is string msg)
                lblStatus.Text = msg;
        }

        private void Worker_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            btnGenerate.Enabled = true;
            lblStatus.Text      = e.Cancelled ? "Cancelled." : "Done.";

            if (e.Error != null)
                MessageBox.Show($"Unexpected error:\n{e.Error.Message}", "StatementFile",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        // ── InitializeComponent (inline, no Designer.cs dependency) ───────────────

        private Label     lblStatus;
        private ProgressBar progressBar;
        private Button    btnGenerate, btnCancel, btnMerchant;
        private CheckedListBox clbProducts;
        private TextBox   txtBankFullName, txtBankName, txtBankCode;

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.Text          = "StatementFile — Generate Statements";
            this.Size          = new System.Drawing.Size(800, 560);
            this.StartPosition = FormStartPosition.CenterScreen;

            clbProducts = new CheckedListBox
            {
                Left = 10, Top = 10, Width = 400, Height = 440,
                CheckOnClick = true,
            };

            var pnlBank = new Panel { Left = 430, Top = 10, Width = 340, Height = 120 };
            pnlBank.Controls.Add(new Label { Text = "Bank Full Name:", Left = 0, Top = 5,  Width = 100 });
            txtBankFullName = new TextBox  { Left = 110, Top = 2,  Width = 220 };
            pnlBank.Controls.Add(txtBankFullName);
            pnlBank.Controls.Add(new Label { Text = "Bank Name:",      Left = 0, Top = 35, Width = 100 });
            txtBankName     = new TextBox  { Left = 110, Top = 32, Width = 220 };
            pnlBank.Controls.Add(txtBankName);
            pnlBank.Controls.Add(new Label { Text = "Bank Code:",      Left = 0, Top = 65, Width = 100 });
            txtBankCode     = new TextBox  { Left = 110, Top = 62, Width = 100 };
            pnlBank.Controls.Add(txtBankCode);

            btnGenerate = new Button { Text = "Generate",         Left = 430, Top = 140, Width = 120 };
            btnCancel   = new Button { Text = "Cancel",           Left = 560, Top = 140, Width = 80  };
            btnMerchant = new Button { Text = "Merchant XML…",    Left = 430, Top = 180, Width = 120 };

            progressBar = new ProgressBar { Left = 430, Top = 220, Width = 340, Height = 24 };
            lblStatus   = new Label       { Left = 430, Top = 255, Width = 340, Height = 40, Text = "Ready." };

            btnGenerate.Click += btnGenerate_Click;
            btnCancel.Click   += btnCancel_Click;
            btnMerchant.Click += btnMerchant_Click;

            Controls.AddRange(new Control[]
            {
                clbProducts, pnlBank,
                btnGenerate, btnCancel, btnMerchant,
                progressBar, lblStatus,
            });

            this.ResumeLayout(false);
        }
    }

    /// <summary>
    /// Describes one bank/product selection from the UI.
    /// Populated when the user checks a product in the CheckedListBox.
    /// </summary>
    public sealed class ProductSelection
    {
        public int    BranchCode         { get; set; }
        public string BankName           { get; set; }
        public string BankFullName       { get; set; }
        public string CardProduct        { get; set; }
        public CardType CardType         { get; set; }
        public string FormatterKey       { get; set; }
        public bool   ExcludeReward      { get; set; } = true;
        public bool   ExcludeInstallment { get; set; }

        public override string ToString() => $"{BankName} — {CardProduct}";
    }
}

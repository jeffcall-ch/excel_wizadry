using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace RqGui
{
    /// <summary>
    /// Settings dialog for configuring rq.exe path.
    /// </summary>
    public partial class SettingsForm : Form
    {
        private TextBox txtRqPath = null!;
        private Button btnBrowse = null!;
        private Button btnSave = null!;
        private Button btnCancel = null!;
        private Label lblRqPath = null!;
        
        public SettingsForm()
        {
            InitializeComponent();
            LoadCurrentSettings();
        }
        
        private void InitializeComponent()
        {
            this.Text = "RQ GUI Settings";
            this.Size = new Size(550, 170);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Font = new Font("Segoe UI", 9F);
            
            // Label
            lblRqPath = new Label
            {
                Text = "RQ.exe Location:",
                Location = new Point(15, 25),
                Size = new Size(120, 20),
                Font = new Font("Segoe UI", 9F)
            };
            
            // TextBox
            txtRqPath = new TextBox
            {
                Location = new Point(140, 23),
                Size = new Size(300, 23),
                TabIndex = 0,
                Font = new Font("Segoe UI", 9F)
            };
            
            // Browse button
            btnBrowse = new Button
            {
                Text = "Browse...",
                Location = new Point(450, 21),
                Size = new Size(80, 27),
                TabIndex = 1,
                Font = new Font("Segoe UI", 9F),
                Cursor = Cursors.Hand
            };
            btnBrowse.Click += BtnBrowse_Click;
            
            // Info label
            var lblInfo = new Label
            {
                Text = "Download rq.exe from: github.com/seeyebe/rq/releases",
                Location = new Point(15, 55),
                Size = new Size(515, 40),
                Font = new Font("Segoe UI", 8.25F),
                ForeColor = Color.Gray,
                AutoSize = false
            };
            
            // Save button
            btnSave = new Button
            {
                Text = "Save",
                Location = new Point(350, 95),
                Size = new Size(85, 30),
                TabIndex = 2,
                Font = new Font("Segoe UI", 9F),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += BtnSave_Click;
            
            // Cancel button
            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(445, 95),
                Size = new Size(85, 30),
                TabIndex = 3,
                DialogResult = DialogResult.Cancel,
                Font = new Font("Segoe UI", 9F),
                Cursor = Cursors.Hand
            };
            btnCancel.Click += (s, e) => this.Close();
            
            this.CancelButton = btnCancel;
            this.AcceptButton = btnSave;
            
            this.Controls.Add(lblRqPath);
            this.Controls.Add(txtRqPath);
            this.Controls.Add(btnBrowse);
            this.Controls.Add(lblInfo);
            this.Controls.Add(btnSave);
            this.Controls.Add(btnCancel);
        }
        
        private void LoadCurrentSettings()
        {
            var settings = SettingsManager.LoadSettings();
            txtRqPath.Text = settings.RqExePath;
        }
        
        private void BtnBrowse_Click(object? sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*";
                dialog.Title = "Select rq.exe";
                
                if (!string.IsNullOrWhiteSpace(txtRqPath.Text) && File.Exists(txtRqPath.Text))
                {
                    dialog.InitialDirectory = Path.GetDirectoryName(txtRqPath.Text);
                    dialog.FileName = Path.GetFileName(txtRqPath.Text);
                }
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtRqPath.Text = dialog.FileName;
                }
            }
        }
        
        private void BtnSave_Click(object? sender, EventArgs e)
        {
            string path = txtRqPath.Text.Trim();
            
            // Validate
            if (string.IsNullOrWhiteSpace(path))
            {
                MessageBox.Show(
                    "Please specify the location of rq.exe",
                    "Validation Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                txtRqPath.Focus();
                return;
            }
            
            if (!File.Exists(path))
            {
                MessageBox.Show(
                    "The specified file does not exist.",
                    "Validation Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                txtRqPath.Focus();
                return;
            }
            
            if (!path.EndsWith("rq.exe", StringComparison.OrdinalIgnoreCase))
            {
                var result = MessageBox.Show(
                    "The selected file is not named 'rq.exe'. Continue anyway?",
                    "Warning",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );
                if (result != DialogResult.Yes)
                    return;
            }
            
            // Save
            var settings = SettingsManager.LoadSettings();
            settings.RqExePath = path;
            
            if (SettingsManager.SaveSettings(settings))
            {
                Logger.LogInfo($"Settings saved. RQ path: {path}");
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }
    }
}

using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using SpecHelper.Models;
using SpecHelper.Services;

namespace SpecHelper.UI
{
    public class SettingsForm : Form
    {
        private TextBox txtMaterials;
        private TextBox txtHardware;
        private TextBox txtComponents;
        private TextBox txtTemplates;
        private CheckBox chkAutoSync;
        private CheckBox chkAutoAdvance;
        private NumericUpDown numDebounce;
        private Button btnSave;
        private Button btnCancel;

        public PluginConfig CurrentConfig { get; private set; }

        public SettingsForm(PluginConfig config)
        {
            CurrentConfig = config ?? new PluginConfig();
            InitializeComponent();
            LoadValues();
        }

        private void InitializeComponent()
        {
            this.Text = "Настройки справочников SpecHelper";
            this.Width = 620;
            this.Height = 435;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular);

            var lblHeader = new Label
            {
                Text = "Пути к файлам баз данных (.xlsx):",
                Location = new Point(15, 12),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold)
            };
            this.Controls.Add(lblHeader);

            // Материалы
            var lblMat = new Label { Text = "Материалы:", Location = new Point(15, 42), AutoSize = true };
            txtMaterials = new TextBox { Location = new Point(15, 62), Width = 480 };
            var btnBrowseMat = new Button { Text = "Обзор...", Location = new Point(502, 60), Width = 85, Height = 26 };
            btnBrowseMat.Click += (s, e) => BrowseFile(txtMaterials, "Выберите файл базы Материалов");
            this.Controls.AddRange(new Control[] { lblMat, txtMaterials, btnBrowseMat });

            // Метизы
            var lblHw = new Label { Text = "Метизы и стандартные изделия:", Location = new Point(15, 95), AutoSize = true };
            txtHardware = new TextBox { Location = new Point(15, 115), Width = 480 };
            var btnBrowseHw = new Button { Text = "Обзор...", Location = new Point(502, 113), Width = 85, Height = 26 };
            btnBrowseHw.Click += (s, e) => BrowseFile(txtHardware, "Выберите файл базы Метизов");
            this.Controls.AddRange(new Control[] { lblHw, txtHardware, btnBrowseHw });

            // Комплектующие
            var lblComp = new Label { Text = "Покупные комплектующие:", Location = new Point(15, 148), AutoSize = true };
            txtComponents = new TextBox { Location = new Point(15, 168), Width = 480 };
            var btnBrowseComp = new Button { Text = "Обзор...", Location = new Point(502, 166), Width = 85, Height = 26 };
            btnBrowseComp.Click += (s, e) => BrowseFile(txtComponents, "Выберите файл базы Комплектующих");
            this.Controls.AddRange(new Control[] { lblComp, txtComponents, btnBrowseComp });

            // Шаблоны
            var lblTemplates = new Label { Text = "База шаблонов узлов и изделий:", Location = new Point(15, 201), AutoSize = true };
            txtTemplates = new TextBox { Location = new Point(15, 221), Width = 480 };
            var btnBrowseTemplates = new Button { Text = "Обзор...", Location = new Point(502, 219), Width = 85, Height = 26 };
            btnBrowseTemplates.Click += (s, e) => BrowseFile(txtTemplates, "Выберите файл базы Шаблонов");
            this.Controls.AddRange(new Control[] { lblTemplates, txtTemplates, btnBrowseTemplates });

            // Опции
            chkAutoSync = new CheckBox
            {
                Text = "Фоновое автообновление при изменении файлов на диске",
                Location = new Point(15, 260),
                AutoSize = true
            };

            chkAutoAdvance = new CheckBox
            {
                Text = "Автоматически переходить на строку ниже после вставки позиции",
                Location = new Point(15, 285),
                AutoSize = true
            };

            var lblDebounce = new Label
            {
                Text = "Задержка обновления (мс):",
                Location = new Point(15, 315),
                AutoSize = true
            };

            numDebounce = new NumericUpDown
            {
                Location = new Point(185, 313),
                Width = 80,
                Minimum = 200,
                Maximum = 5000,
                Increment = 100
            };

            this.Controls.AddRange(new Control[] { chkAutoSync, chkAutoAdvance, lblDebounce, numDebounce });

            // Кнопки
            btnSave = new Button { Text = "Сохранить", DialogResult = DialogResult.OK, Location = new Point(390, 350), Width = 95, Height = 30 };
            btnCancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel, Location = new Point(495, 350), Width = 92, Height = 30 };

            btnSave.Click += BtnSave_Click;
            this.AcceptButton = btnSave;
            this.CancelButton = btnCancel;

            this.Controls.AddRange(new Control[] { btnSave, btnCancel });
        }

        private void LoadValues()
        {
            txtMaterials.Text = CurrentConfig.MaterialsPath ?? string.Empty;
            txtHardware.Text = CurrentConfig.HardwarePath ?? string.Empty;
            txtComponents.Text = CurrentConfig.ComponentsPath ?? string.Empty;
            txtTemplates.Text = CurrentConfig.TemplatesPath ?? string.Empty;

            chkAutoSync.Checked = CurrentConfig.AutoSyncEnabled;
            chkAutoAdvance.Checked = CurrentConfig.AutoAdvanceRow;
            numDebounce.Value = Math.Max(200, Math.Min(5000, CurrentConfig.SyncDebounceMs));
        }

        private void BrowseFile(TextBox targetBox, string title)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Title = title;
                ofd.Filter = "Книги Excel (*.xlsx;*.xls)|*.xlsx;*.xls|Все файлы (*.*)|*.*";
                if (!string.IsNullOrWhiteSpace(targetBox.Text) && File.Exists(targetBox.Text))
                {
                    ofd.InitialDirectory = Path.GetDirectoryName(targetBox.Text);
                    ofd.FileName = Path.GetFileName(targetBox.Text);
                }
                if (ofd.ShowDialog(this) == DialogResult.OK)
                {
                    targetBox.Text = ofd.FileName;
                }
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            CurrentConfig.MaterialsPath = txtMaterials.Text.Trim();
            CurrentConfig.HardwarePath = txtHardware.Text.Trim();
            CurrentConfig.ComponentsPath = txtComponents.Text.Trim();
            CurrentConfig.TemplatesPath = txtTemplates.Text.Trim();
            CurrentConfig.AutoSyncEnabled = chkAutoSync.Checked;
            CurrentConfig.AutoAdvanceRow = chkAutoAdvance.Checked;
            CurrentConfig.SyncDebounceMs = (int)numDebounce.Value;

            ConfigManager.Save(CurrentConfig);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using SpecHelper.Models;
using SpecHelper.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace SpecHelper.UI
{
    [ComVisible(true)]
    [Guid("D86F2C74-3208-4CE0-A4F7-1B8C512F9E82")]
    [ProgId("SpecHelper.UI.SpecPaneControl")]
    public class SpecPaneControl : UserControl
    {
        public Excel.Application ExcelApp { get; set; }
        public CustomTaskPane TaskPane { get; set; }

        private PluginConfig _config;
        private DatabaseWatcher _watcher;

        private readonly Dictionary<CatalogType, CatalogNode> _catalogRoots = new Dictionary<CatalogType, CatalogNode>();
        private readonly Dictionary<CatalogType, List<CatalogItem>> _catalogItems = new Dictionary<CatalogType, List<CatalogItem>>();

        // Элементы управления
        private Panel pnlHeader;
        private FlowLayoutPanel flowSearchRow;
        private FlowLayoutPanel flowButtonsRow;

        private Label lblSearch;
        private TextBox txtSearch;
        private Button btnClearSearch;
        private CheckBox chkInFolder;

        private Button btnToggleTree;
        private Button btnRefresh;
        private Button btnSettings;
        private Button btnDockToggle;

        private TabControl tabCatalogs;
        private TabPage tabMaterials;
        private TabPage tabHardware;
        private TabPage tabComponents;
        private TabPage tabTemplates;
        private TabPage tabTechProcesses;

        // Элементы вкладки "Шаблоны"
        private Panel pnlTemplatesView;
        private Panel pnlTemplatesHeader;
        private FlowLayoutPanel flowTemplatesSearch;
        private FlowLayoutPanel flowTemplatesActions;
        private Label lblSearchTemplates;
        private TextBox txtSearchTemplates;
        private Button btnClearSearchTemplates;
        private Button btnRefreshTemplates;
        private Button btnOpenTemplateFile;
        private Button btnSettingsTemplates;
        private ListBox lstTemplates;
        private Panel pnlTemplatesFooter;
        private Label lblTemplateStatus;
        private Button btnInsertTemplate;
        private List<TemplateItem> _allTemplates = new List<TemplateItem>();

        // Элементы вкладки "ТП" (Техпроцессы)
        private Panel pnlTechProcessesView;
        private Panel pnlTpHeader;
        private FlowLayoutPanel flowTpSearchRow;
        private FlowLayoutPanel flowTpActionsRow;
        private Label lblSearchTp;
        private TextBox txtSearchTp;
        private Button btnClearSearchTp;
        private Button btnResetOps;
        private Button btnRefreshTp;
        private Button btnSettingsTp;
        private SplitContainer splitTpContainer;
        private Label lblOperationsTitle;
        private TabControl tabOperationSheets;
        private DataGridView gridOperations;
        private Label lblTechProcessesTitle;
        private DataGridView gridTechProcesses;
        private Panel pnlTpFooter;
        private Label lblTpStatus;
        private Button btnInsertTp;

        private TechProcessCatalogData _tpData = new TechProcessCatalogData();
        private readonly HashSet<string> _selectedOperationCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private bool _isUpdatingOpGrid = false;

        private SplitContainer splitContainer;
        private TreeView treeView;
        private DataGridView gridItems;

        private Panel pnlFooter;
        private Label lblStatus;
        private Button btnInsert;

        private bool _isUpdatingUi = false;

        public SpecPaneControl()
        {
            InitializeComponent();
        }

        public void Initialize(Excel.Application excelApp, CustomTaskPane taskPane = null)
        {
            ExcelApp = excelApp;
            TaskPane = taskPane;
            _config = ConfigManager.Load();

            try
            {
                AdaptOrientation();
            }
            catch { }

            if (_config.IsTreeCollapsed)
            {
                SetTreeCollapsed(true);
            }

            chkInFolder.Checked = _config.SearchCurrentFolderOnly;
            if (btnDockToggle != null)
            {
                btnDockToggle.Visible = (taskPane != null);
            }
            UpdateDockButtonText();

            _watcher = new DatabaseWatcher(_config.SyncDebounceMs);
            _watcher.FileChanged += Watcher_FileChanged;
            _watcher.UpdatePaths(_config);

            ReloadAllCatalogs();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            this.Size = new Size(500, 700);
            this.BackColor = Color.FromArgb(240, 243, 246);
            this.ForeColor = Color.FromArgb(20, 20, 20);

            // ================= 1. ВЕРХНЯЯ ПАНЕЛЬ ИНСТРУМЕНТОВ =================
            pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.FromArgb(235, 238, 242),
                Padding = new Padding(4, 4, 4, 4)
            };

            flowSearchRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = true,
                FlowDirection = FlowDirection.LeftToRight,
                BackColor = Color.Transparent,
                Margin = new Padding(0)
            };

            lblSearch = new Label
            {
                Text = "Поиск:",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(20, 20, 20),
                AutoSize = true,
                Margin = new Padding(4, 7, 2, 2)
            };

            txtSearch = new TextBox
            {
                Width = 180,
                Height = 24,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                BackColor = Color.White,
                ForeColor = Color.FromArgb(15, 15, 15),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(2, 4, 0, 2)
            };
            txtSearch.TextChanged += (s, e) => ApplyFilter();

            btnClearSearch = CreateStyledButton("✕", 26, 25);
            btnClearSearch.Font = new Font("Segoe UI", 8F, FontStyle.Bold);
            btnClearSearch.Margin = new Padding(1, 4, 6, 2);
            btnClearSearch.Click += (s, e) => { txtSearch.Clear(); txtSearch.Focus(); };

            chkInFolder = new CheckBox
            {
                Text = "В выбранной папке",
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                ForeColor = Color.FromArgb(20, 20, 20),
                AutoSize = true,
                Margin = new Padding(4, 6, 4, 2)
            };
            chkInFolder.CheckedChanged += (s, e) =>
            {
                _config.SearchCurrentFolderOnly = chkInFolder.Checked;
                ConfigManager.Save(_config);
                ApplyFilter();
            };

            flowSearchRow.Controls.AddRange(new Control[] { lblSearch, txtSearch, btnClearSearch, chkInFolder });

            flowButtonsRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = true,
                FlowDirection = FlowDirection.LeftToRight,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 4, 0, 0)
            };

            btnToggleTree = CreateStyledButton("▲ Скрыть папки", 125, 27);
            btnToggleTree.Click += BtnToggleTree_Click;

            btnRefresh = CreateStyledButton("Обновить", 85, 27);
            btnRefresh.Click += (s, e) => ReloadAllCatalogs();

            btnSettings = CreateStyledButton("Настройки", 90, 27);
            btnSettings.Click += BtnSettings_Click;

            btnDockToggle = CreateStyledButton("Закрепить внизу", 125, 27);
            btnDockToggle.Click += BtnDockToggle_Click;

            flowButtonsRow.Controls.AddRange(new Control[] { btnToggleTree, btnRefresh, btnSettings, btnDockToggle });

            pnlHeader.Controls.Add(flowButtonsRow);
            pnlHeader.Controls.Add(flowSearchRow);

            var tip = new ToolTip();
            tip.SetToolTip(btnToggleTree, "Скрыть или показать дерево папок");
            tip.SetToolTip(btnClearSearch, "Очистить строку поиска");
            tip.SetToolTip(btnRefresh, "Перечитать базы данных с диска");
            tip.SetToolTip(btnSettings, "Открыть окно настроек путей к файлам");
            tip.SetToolTip(btnDockToggle, "Переключить прикрепление панели (справа / снизу листа)");

            // ================= 2. ВКЛАДКИ (TABCONTROL) =================
            tabCatalogs = new TabControl
            {
                Dock = DockStyle.Top,
                Height = 30,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular)
            };
            tabMaterials = new TabPage("Материалы");
            tabHardware = new TabPage("Метизы и крепеж");
            tabComponents = new TabPage("Комплектующие");
            tabTemplates = new TabPage("Шаблоны");
            tabTechProcesses = new TabPage("ТП");

            tabCatalogs.TabPages.AddRange(new TabPage[] { tabMaterials, tabHardware, tabComponents, tabTemplates, tabTechProcesses });
            tabCatalogs.SelectedIndexChanged += TabCatalogs_SelectedIndexChanged;

            // ================= 3. НИЖНЯЯ ПАНЕЛЬ (FOOTER) =================
            pnlFooter = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 44,
                Padding = new Padding(8, 6, 8, 6),
                BackColor = Color.FromArgb(232, 235, 240)
            };

            btnInsert = new Button
            {
                Text = "Вставить в спецификацию [Enter]",
                Dock = DockStyle.Right,
                Width = 240,
                Height = 32,
                BackColor = Color.FromArgb(16, 124, 65),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnInsert.FlatAppearance.BorderSize = 0;
            btnInsert.Click += (s, e) => InsertSelectedItem();

            lblStatus = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "Готово к работе",
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                ForeColor = Color.FromArgb(20, 20, 20),
                AutoEllipsis = true
            };

            pnlFooter.Controls.Add(lblStatus);
            pnlFooter.Controls.Add(btnInsert);

            // ================= 4. ЦЕНТРАЛЬНЫЙ SPLITCONTAINER =================
            splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterWidth = 6,
                Panel1MinSize = 30,
                Panel2MinSize = 30,
                BackColor = Color.FromArgb(215, 220, 226)
            };
            splitContainer.SplitterMoved += (s, e) =>
            {
                if (!splitContainer.Panel1Collapsed && _config != null)
                {
                    _config.SplitterDistance = splitContainer.SplitterDistance;
                    ConfigManager.Save(_config);
                }
            };
            splitContainer.SizeChanged += (s, e) =>
            {
                if (!splitContainer.Panel1Collapsed)
                {
                    SetSafeSplitterDistance(splitContainer.Orientation == Orientation.Horizontal ? 150 : 220);
                }
            };

            treeView = new TreeView
            {
                Dock = DockStyle.Fill,
                HideSelection = false,
                BorderStyle = BorderStyle.None,
                ShowLines = true,
                ShowPlusMinus = true,
                FullRowSelect = true,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                BackColor = Color.White,
                ForeColor = Color.FromArgb(15, 15, 15),
                LineColor = Color.FromArgb(120, 130, 140),
                ItemHeight = 22
            };
            treeView.AfterSelect += TreeView_AfterSelect;
            splitContainer.Panel1.Controls.Add(treeView);

            gridItems = new DataGridView
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                BackgroundColor = Color.White,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AllowUserToResizeColumns = true,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                AutoGenerateColumns = false,
                ShowCellToolTips = true,
                GridColor = Color.FromArgb(220, 225, 232),
                RowTemplate = { Height = 26 },
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                EnableHeadersVisualStyles = false
            };

            StyleDataGridView(gridItems);

            var colDesig = new DataGridViewTextBoxColumn
            {
                HeaderText = "Обозначение",
                DataPropertyName = "Designation",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 45,
                MinimumWidth = 120
            };

            var colName = new DataGridViewTextBoxColumn
            {
                HeaderText = "Наименование",
                DataPropertyName = "Name",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 55,
                MinimumWidth = 140
            };

            gridItems.Columns.AddRange(new DataGridViewColumn[] { colDesig, colName });
            gridItems.CellDoubleClick += (s, e) => InsertSelectedItem();
            gridItems.KeyDown += GridItems_KeyDown;

            splitContainer.Panel2.Controls.Add(gridItems);

            // ================= 5. ПАНЕЛЬ ВКЛАДКИ ШАБЛОНЫ =================
            pnlTemplatesView = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(240, 243, 246),
                Visible = false
            };

            pnlTemplatesHeader = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.FromArgb(235, 238, 242),
                Padding = new Padding(4)
            };

            flowTemplatesSearch = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = true,
                FlowDirection = FlowDirection.LeftToRight,
                BackColor = Color.Transparent,
                Margin = new Padding(0)
            };

            lblSearchTemplates = new Label
            {
                Text = "Поиск:",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(20, 20, 20),
                AutoSize = true,
                Margin = new Padding(4, 7, 2, 2)
            };

            txtSearchTemplates = new TextBox
            {
                Width = 180,
                Height = 24,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                BackColor = Color.White,
                ForeColor = Color.FromArgb(15, 15, 15),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(2, 4, 0, 2)
            };
            txtSearchTemplates.TextChanged += (s, e) => ApplyTemplateFilter();

            btnClearSearchTemplates = CreateStyledButton("✕", 26, 25);
            btnClearSearchTemplates.Font = new Font("Segoe UI", 8F, FontStyle.Bold);
            btnClearSearchTemplates.Margin = new Padding(1, 4, 6, 2);
            btnClearSearchTemplates.Click += (s, e) => { txtSearchTemplates.Clear(); txtSearchTemplates.Focus(); };

            flowTemplatesSearch.Controls.AddRange(new Control[] { lblSearchTemplates, txtSearchTemplates, btnClearSearchTemplates });

            flowTemplatesActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = true,
                FlowDirection = FlowDirection.LeftToRight,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 4, 0, 0)
            };

            btnRefreshTemplates = CreateStyledButton("🔄 Обновить", 110, 27);
            btnRefreshTemplates.Click += (s, e) => RefreshTemplatesSafely();

            btnOpenTemplateFile = CreateStyledButton("📁 Открыть файл шаблонов", 195, 27);
            btnOpenTemplateFile.Click += (s, e) => OpenTemplateFile();

            btnSettingsTemplates = CreateStyledButton("Настройки", 90, 27);
            btnSettingsTemplates.Click += BtnSettings_Click;

            flowTemplatesActions.Controls.AddRange(new Control[] { btnRefreshTemplates, btnOpenTemplateFile, btnSettingsTemplates });

            pnlTemplatesHeader.Controls.Add(flowTemplatesActions);
            pnlTemplatesHeader.Controls.Add(flowTemplatesSearch);

            pnlTemplatesFooter = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 84,
                Padding = new Padding(8, 6, 8, 8),
                BackColor = Color.FromArgb(232, 235, 240)
            };

            lblTemplateStatus = new Label
            {
                Dock = DockStyle.Top,
                Height = 28,
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "Выберите шаблон и целевую ячейку в столбце A",
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                ForeColor = Color.FromArgb(20, 20, 20),
                AutoEllipsis = true
            };

            btnInsertTemplate = new Button
            {
                Text = "Вставить шаблон [Enter]",
                Dock = DockStyle.Bottom,
                Height = 38,
                BackColor = Color.FromArgb(16, 124, 65),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnInsertTemplate.FlatAppearance.BorderSize = 0;
            btnInsertTemplate.Click += (s, e) => InsertSelectedTemplate();

            pnlTemplatesFooter.Controls.Add(btnInsertTemplate);
            pnlTemplatesFooter.Controls.Add(lblTemplateStatus);

            lstTemplates = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                ItemHeight = 32,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(15, 15, 15),
                IntegralHeight = false,
                DrawMode = DrawMode.OwnerDrawFixed
            };
            lstTemplates.DrawItem += LstTemplates_DrawItem;
            lstTemplates.SelectedIndexChanged += LstTemplates_SelectedIndexChanged;
            lstTemplates.DoubleClick += (s, e) => InsertSelectedTemplate();
            lstTemplates.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    InsertSelectedTemplate();
                }
            };

            pnlTemplatesView.Controls.Add(lstTemplates);
            pnlTemplatesView.Controls.Add(pnlTemplatesHeader);
            pnlTemplatesView.Controls.Add(pnlTemplatesFooter);

            // ================= 6. ПАНЕЛЬ ВКЛАДКИ ТП (ТЕХПРОЦЕССЫ) =================
            InitializeTechProcessesView();

            // Сборка элементов
            this.Controls.Add(pnlTechProcessesView);
            this.Controls.Add(pnlTemplatesView);
            this.Controls.Add(splitContainer);
            this.Controls.Add(tabCatalogs);
            this.Controls.Add(pnlHeader);
            this.Controls.Add(pnlFooter);

            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void InitializeTechProcessesView()
        {
            pnlTechProcessesView = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(240, 243, 246),
                Visible = false
            };

            pnlTpHeader = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.FromArgb(235, 238, 242),
                Padding = new Padding(4)
            };

            flowTpSearchRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = true,
                FlowDirection = FlowDirection.LeftToRight,
                BackColor = Color.Transparent,
                Margin = new Padding(0)
            };

            lblSearchTp = new Label
            {
                Text = "Поиск:",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(20, 20, 20),
                AutoSize = true,
                Margin = new Padding(4, 7, 2, 2)
            };

            txtSearchTp = new TextBox
            {
                Width = 175,
                Height = 24,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                BackColor = Color.White,
                ForeColor = Color.FromArgb(15, 15, 15),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(2, 4, 0, 2)
            };
            txtSearchTp.TextChanged += (s, e) =>
            {
                PopulateOperationsGrid();
                ApplyTechProcessFilter();
            };

            btnClearSearchTp = CreateStyledButton("✕", 26, 25);
            btnClearSearchTp.Font = new Font("Segoe UI", 8F, FontStyle.Bold);
            btnClearSearchTp.Margin = new Padding(1, 4, 6, 2);
            btnClearSearchTp.Click += (s, e) => { txtSearchTp.Clear(); txtSearchTp.Focus(); };

            btnResetOps = CreateStyledButton("Сбросить галочки", 130, 25);
            btnResetOps.Margin = new Padding(2, 4, 4, 2);
            btnResetOps.Enabled = false;
            btnResetOps.Click += (s, e) =>
            {
                _selectedOperationCodes.Clear();
                UpdateResetOpsButton();
                PopulateOperationsGrid();
                ApplyTechProcessFilter();
            };

            flowTpSearchRow.Controls.AddRange(new Control[] { lblSearchTp, txtSearchTp, btnClearSearchTp, btnResetOps });

            flowTpActionsRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = true,
                FlowDirection = FlowDirection.LeftToRight,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 4, 0, 0)
            };

            btnRefreshTp = CreateStyledButton("Обновить", 85, 27);
            btnRefreshTp.Click += (s, e) => ReloadTechProcesses();

            btnSettingsTp = CreateStyledButton("Настройки", 90, 27);
            btnSettingsTp.Click += BtnSettings_Click;

            flowTpActionsRow.Controls.AddRange(new Control[] { btnRefreshTp, btnSettingsTp });

            pnlTpHeader.Controls.Add(flowTpActionsRow);
            pnlTpHeader.Controls.Add(flowTpSearchRow);

            pnlTpFooter = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 44,
                Padding = new Padding(8, 6, 8, 6),
                BackColor = Color.FromArgb(232, 235, 240)
            };

            btnInsertTp = new Button
            {
                Text = "Вставить в спецификацию [Enter]",
                Dock = DockStyle.Right,
                Width = 240,
                Height = 32,
                BackColor = Color.FromArgb(16, 124, 65),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnInsertTp.FlatAppearance.BorderSize = 0;
            btnInsertTp.Click += (s, e) => InsertSelectedTechProcess();

            lblTpStatus = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "Выберите операции галочками для поиска техпроцесса",
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                ForeColor = Color.FromArgb(20, 20, 20),
                AutoEllipsis = true
            };

            pnlTpFooter.Controls.Add(lblTpStatus);
            pnlTpFooter.Controls.Add(btnInsertTp);

            splitTpContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterWidth = 6,
                Panel1MinSize = 80,
                Panel2MinSize = 80,
                BackColor = Color.FromArgb(215, 220, 226)
            };
            splitTpContainer.SizeChanged += (s, e) =>
            {
                int total = splitTpContainer.Orientation == Orientation.Horizontal
                    ? splitTpContainer.Height
                    : splitTpContainer.Width;
                if (total > 180)
                {
                    int target = (int)(total * 0.48);
                    int min = splitTpContainer.Panel1MinSize + 5;
                    int max = total - splitTpContainer.Panel2MinSize - 5;
                    if (max > min)
                    {
                        try { splitTpContainer.SplitterDistance = Math.Max(min, Math.Min(target, max)); } catch { }
                    }
                }
            };

            // Верхняя секция: Операции с динамическими подвкладками и чекбоксами
            lblOperationsTitle = new Label
            {
                Dock = DockStyle.Top,
                Height = 24,
                Text = "  Операции",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                BackColor = Color.FromArgb(230, 234, 240),
                ForeColor = Color.FromArgb(20, 20, 20)
            };

            tabOperationSheets = new TabControl
            {
                Dock = DockStyle.Top,
                Height = 26,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };
            tabOperationSheets.SelectedIndexChanged += (s, e) => PopulateOperationsGrid();

            gridOperations = new DataGridView
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                BackgroundColor = Color.White,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AllowUserToResizeColumns = true,
                ReadOnly = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                AutoGenerateColumns = false,
                ShowCellToolTips = true,
                GridColor = Color.FromArgb(220, 225, 232),
                RowTemplate = { Height = 24 },
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                EnableHeadersVisualStyles = false
            };
            StyleDataGridView(gridOperations);

            var colOpCheck = new DataGridViewCheckBoxColumn
            {
                Name = "colOpSelected",
                HeaderText = "✓",
                DataPropertyName = "IsSelected",
                Width = 32,
                Resizable = DataGridViewTriState.False,
                ReadOnly = false
            };
            var colOpCode = new DataGridViewTextBoxColumn
            {
                Name = "colOpCode",
                HeaderText = "Обозначение",
                DataPropertyName = "Code",
                Width = 90,
                ReadOnly = true
            };
            var colOpName = new DataGridViewTextBoxColumn
            {
                Name = "colOpName",
                HeaderText = "Наименование",
                DataPropertyName = "Name",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 120,
                ReadOnly = true
            };

            gridOperations.Columns.AddRange(new DataGridViewColumn[] { colOpCheck, colOpCode, colOpName });
            gridOperations.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (gridOperations.IsCurrentCellDirty)
                {
                    gridOperations.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            };
            gridOperations.CellValueChanged += GridOperations_CellValueChanged;

            splitTpContainer.Panel1.Controls.Add(gridOperations);
            splitTpContainer.Panel1.Controls.Add(tabOperationSheets);
            splitTpContainer.Panel1.Controls.Add(lblOperationsTitle);

            // Нижняя секция: Отфильтрованные техпроцессы
            lblTechProcessesTitle = new Label
            {
                Dock = DockStyle.Top,
                Height = 24,
                Text = "  Техпроцессы",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                BackColor = Color.FromArgb(230, 234, 240),
                ForeColor = Color.FromArgb(20, 20, 20)
            };

            gridTechProcesses = new DataGridView
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                BackgroundColor = Color.White,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AllowUserToResizeColumns = true,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                AutoGenerateColumns = false,
                ShowCellToolTips = true,
                GridColor = Color.FromArgb(220, 225, 232),
                RowTemplate = { Height = 26 },
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                EnableHeadersVisualStyles = false
            };
            StyleDataGridView(gridTechProcesses);

            var colTpDesig = new DataGridViewTextBoxColumn
            {
                HeaderText = "Обозначение",
                DataPropertyName = "Designation",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 42,
                MinimumWidth = 110
            };
            var colTpComp = new DataGridViewTextBoxColumn
            {
                HeaderText = "Наименование / состав техпроцесса",
                DataPropertyName = "DisplayComposition",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 58,
                MinimumWidth = 140
            };

            gridTechProcesses.Columns.AddRange(new DataGridViewColumn[] { colTpDesig, colTpComp });
            gridTechProcesses.CellDoubleClick += (s, e) => InsertSelectedTechProcess();
            gridTechProcesses.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    InsertSelectedTechProcess();
                }
            };

            splitTpContainer.Panel2.Controls.Add(gridTechProcesses);
            splitTpContainer.Panel2.Controls.Add(lblTechProcessesTitle);

            pnlTechProcessesView.Controls.Add(splitTpContainer);
            pnlTechProcessesView.Controls.Add(pnlTpHeader);
            pnlTechProcessesView.Controls.Add(pnlTpFooter);
        }

        private void StyleDataGridView(DataGridView dgv)
        {
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(225, 230, 238);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(20, 20, 20);
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            dgv.ColumnHeadersHeight = 28;

            dgv.DefaultCellStyle.BackColor = Color.White;
            dgv.DefaultCellStyle.ForeColor = Color.FromArgb(15, 15, 15);
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(16, 124, 65);
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;

            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 247, 251);
            dgv.AlternatingRowsDefaultCellStyle.ForeColor = Color.FromArgb(15, 15, 15);
            dgv.AlternatingRowsDefaultCellStyle.SelectionBackColor = Color.FromArgb(16, 124, 65);
            dgv.AlternatingRowsDefaultCellStyle.SelectionForeColor = Color.White;
        }

        private Button CreateStyledButton(string text, int width, int height)
        {
            var btn = new Button
            {
                Text = text,
                Width = width,
                Height = height,
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(226, 230, 236),
                ForeColor = Color.FromArgb(20, 20, 20),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Margin = new Padding(2, 2, 4, 2)
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(170, 176, 186);
            btn.FlatAppearance.BorderSize = 1;
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(210, 218, 228);
            return btn;
        }

        private void AdaptOrientation()
        {
            try
            {
                if (TaskPane == null)
                {
                    splitContainer.Orientation = Orientation.Horizontal;
                    if (splitTpContainer != null) splitTpContainer.Orientation = Orientation.Horizontal;
                    SetSafeSplitterDistance(150);
                    return;
                }

                if (TaskPane.DockPosition == MsoCTPDockPosition.msoCTPDockPositionBottom ||
                    TaskPane.DockPosition == MsoCTPDockPosition.msoCTPDockPositionTop)
                {
                    splitContainer.Orientation = Orientation.Vertical;
                    if (splitTpContainer != null) splitTpContainer.Orientation = Orientation.Vertical;
                    SetSafeSplitterDistance(220);
                }
                else
                {
                    splitContainer.Orientation = Orientation.Horizontal;
                    if (splitTpContainer != null) splitTpContainer.Orientation = Orientation.Horizontal;
                    SetSafeSplitterDistance(150);
                }
            }
            catch { }
        }

        private void SetSafeSplitterDistance(int targetDistance)
        {
            try
            {
                int total = splitContainer.Orientation == Orientation.Horizontal
                    ? splitContainer.Height
                    : splitContainer.Width;

                if (total <= 50) return;

                int min = splitContainer.Panel1MinSize + 5;
                int max = total - splitContainer.Panel2MinSize - 5;

                if (max > min)
                {
                    splitContainer.SplitterDistance = Math.Max(min, Math.Min(targetDistance, max));
                }
            }
            catch { }
        }

        public void OnDockPositionChanged()
        {
            AdaptOrientation();
            UpdateDockButtonText();
        }

        private void UpdateDockButtonText()
        {
            if (TaskPane == null)
            {
                btnDockToggle.Text = "Закрепить внизу";
                return;
            }

            if (TaskPane.DockPosition == MsoCTPDockPosition.msoCTPDockPositionBottom ||
                TaskPane.DockPosition == MsoCTPDockPosition.msoCTPDockPositionTop)
            {
                btnDockToggle.Text = "Закрепить справа";
            }
            else
            {
                btnDockToggle.Text = "Закрепить внизу";
            }
        }

        private void BtnDockToggle_Click(object sender, EventArgs e)
        {
            if (TaskPane == null) return;

            try
            {
                if (TaskPane.DockPosition == MsoCTPDockPosition.msoCTPDockPositionBottom ||
                    TaskPane.DockPosition == MsoCTPDockPosition.msoCTPDockPositionTop)
                {
                    TaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                    try { TaskPane.Width = 500; } catch { }
                }
                else
                {
                    TaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
                    try { TaskPane.Height = 350; } catch { }
                }

                AdaptOrientation();
                UpdateDockButtonText();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SpecPaneControl] Ошибка смены Dock: {ex.Message}");
            }
        }

        private CatalogType CurrentCatalogType
        {
            get
            {
                if (tabCatalogs.SelectedTab == tabMaterials) return CatalogType.Materials;
                if (tabCatalogs.SelectedTab == tabHardware) return CatalogType.Hardware;
                return CatalogType.Components;
            }
        }

        public void ReloadAllCatalogs()
        {
            try
            {
                lblStatus.Text = "Чтение баз данных...";
                Application.DoEvents();

                LoadSingleCatalog(CatalogType.Materials, _config.MaterialsPath);
                LoadSingleCatalog(CatalogType.Hardware, _config.HardwarePath);
                LoadSingleCatalog(CatalogType.Components, _config.ComponentsPath);

                RefreshCurrentView();
                lblStatus.Text = $"Базы обновлены в {DateTime.Now:HH:mm:ss}";
            }
            catch (Exception ex)
            {
                lblStatus.Text = $"Ошибка загрузки: {ex.Message}";
            }
        }

        private void LoadSingleCatalog(CatalogType type, string path)
        {
            try
            {
                var result = ExcelCatalogReader.LoadCatalog(path, type);
                _catalogRoots[type] = result.RootNode;
                _catalogItems[type] = result.AllItems;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SpecPaneControl] Ошибка загрузки {type}: {ex.Message}");
            }
        }

        private void RefreshCurrentView()
        {
            _isUpdatingUi = true;
            try
            {
                var type = CurrentCatalogType;
                treeView.Nodes.Clear();

                if (_catalogRoots.TryGetValue(type, out var root) && root != null)
                {
                    var rootNode = CreateTreeNode(root);
                    treeView.Nodes.Add(rootNode);
                    rootNode.Expand();
                    treeView.SelectedNode = rootNode;
                }

                ApplyFilter();
            }
            finally
            {
                _isUpdatingUi = false;
            }
        }

        private TreeNode CreateTreeNode(CatalogNode catalogNode)
        {
            int count = catalogNode.TotalItemCount;
            string label = count > 0 ? $"{catalogNode.Name} [{count}]" : catalogNode.Name;
            var treeNode = new TreeNode(label) { Tag = catalogNode };

            foreach (var child in catalogNode.Children)
            {
                treeNode.Nodes.Add(CreateTreeNode(child));
            }

            return treeNode;
        }

        private void ApplyFilter()
        {
            if (_isUpdatingUi) return;

            var type = CurrentCatalogType;
            if (!_catalogItems.TryGetValue(type, out var allItems) || allItems == null)
            {
                gridItems.DataSource = null;
                lblStatus.Text = "База пуста или файл не выбран в настройках.";
                return;
            }

            IEnumerable<CatalogItem> candidateItems;

            if (chkInFolder.Checked && treeView.SelectedNode?.Tag is CatalogNode selectedFolder)
            {
                candidateItems = selectedFolder.GetAllItems();
            }
            else
            {
                candidateItems = allItems;
            }

            var query = txtSearch.Text.Trim();
            if (!string.IsNullOrEmpty(query))
            {
                candidateItems = candidateItems.Where(i =>
                    (i.Designation != null && i.Designation.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) ||
                    (i.Name != null && i.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) ||
                    (i.CategoryPath != null && i.CategoryPath.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)
                );
            }

            var resultList = candidateItems.ToList();
            gridItems.DataSource = resultList;
            lblStatus.Text = $"Найдено: {resultList.Count} из {allItems.Count}";
        }

        private void TreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (_isUpdatingUi) return;

            if (e.Node?.Tag is CatalogNode selectedNode)
            {
                if (chkInFolder.Checked || string.IsNullOrEmpty(txtSearch.Text))
                {
                    var items = selectedNode.GetAllItems().ToList();
                    var query = txtSearch.Text.Trim();
                    if (!string.IsNullOrEmpty(query))
                    {
                        items = items.Where(i =>
                            (i.Designation != null && i.Designation.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) ||
                            (i.Name != null && i.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)
                        ).ToList();
                    }
                    gridItems.DataSource = items;
                    lblStatus.Text = $"Папка: {selectedNode.Name} | Позиций: {items.Count}";
                }
            }
        }

        private void TabCatalogs_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabCatalogs.SelectedTab == tabTemplates)
            {
                pnlHeader.Visible = false;
                pnlFooter.Visible = false;
                splitContainer.Visible = false;
                pnlTechProcessesView.Visible = false;
                pnlTemplatesView.Visible = true;
                pnlTemplatesView.BringToFront();
                ReloadTemplates();
            }
            else if (tabCatalogs.SelectedTab == tabTechProcesses)
            {
                pnlHeader.Visible = false;
                pnlFooter.Visible = false;
                splitContainer.Visible = false;
                pnlTemplatesView.Visible = false;
                pnlTechProcessesView.Visible = true;
                pnlTechProcessesView.BringToFront();
                ReloadTechProcesses();
            }
            else
            {
                pnlTemplatesView.Visible = false;
                pnlTechProcessesView.Visible = false;
                pnlHeader.Visible = true;
                pnlFooter.Visible = true;
                splitContainer.Visible = true;
                txtSearch.Clear();
                RefreshCurrentView();
            }
        }

        // ================= ЛОГИКА ВКЛАДКИ ТП =================

        private string GetTechProcessesFilePath()
        {
            if (_config != null && !string.IsNullOrWhiteSpace(_config.TechProcessesPath) && File.Exists(_config.TechProcessesPath))
            {
                return _config.TechProcessesPath;
            }

            string baseDir = Path.GetDirectoryName(typeof(SpecPaneControl).Assembly.Location) ?? AppDomain.CurrentDomain.BaseDirectory;
            string[] candidates = new[]
            {
                Path.Combine(baseDir, "databases", "Техпроцессы.xls"),
                Path.Combine(baseDir, "databases", "Техпроцессы.xlsx"),
                Path.Combine(baseDir, "Техпроцессы.xls"),
                Path.Combine(baseDir, "Техпроцессы.xlsx")
            };

            foreach (string c in candidates)
            {
                if (File.Exists(c)) return c;
            }

            return _config?.TechProcessesPath ?? string.Empty;
        }

        private void ReloadTechProcesses()
        {
            try
            {
                string path = GetTechProcessesFilePath();
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    _tpData = new TechProcessCatalogData();
                    tabOperationSheets.TabPages.Clear();
                    gridOperations.DataSource = null;
                    gridTechProcesses.DataSource = null;
                    lblTpStatus.Text = "Укажите файл Техпроцессы.xls в Настройках";
                    return;
                }

                string prevTabName = tabOperationSheets.SelectedTab?.Text;
                _tpData = ExcelCatalogReader.LoadTechProcesses(path);

                tabOperationSheets.TabPages.Clear();
                foreach (var sheetPair in _tpData.OperationSheets)
                {
                    tabOperationSheets.TabPages.Add(new TabPage(sheetPair.Key));
                }

                if (!string.IsNullOrEmpty(prevTabName))
                {
                    foreach (TabPage tp in tabOperationSheets.TabPages)
                    {
                        if (string.Equals(tp.Text, prevTabName, StringComparison.OrdinalIgnoreCase))
                        {
                            tabOperationSheets.SelectedTab = tp;
                            break;
                        }
                    }
                }

                PopulateOperationsGrid();
                ApplyTechProcessFilter();
            }
            catch (Exception ex)
            {
                lblTpStatus.Text = $"Ошибка загрузки ТП: {ex.Message}";
            }
        }

        private void PopulateOperationsGrid()
        {
            _isUpdatingOpGrid = true;
            try
            {
                int idx = tabOperationSheets.SelectedIndex;
                if (idx < 0 || idx >= _tpData.OperationSheets.Count)
                {
                    gridOperations.DataSource = null;
                    return;
                }

                var ops = _tpData.OperationSheets[idx].Value;
                string query = txtSearchTp.Text.Trim();

                IEnumerable<TechOperationItem> filteredOps = ops;
                if (!string.IsNullOrEmpty(query))
                {
                    filteredOps = filteredOps.Where(o =>
                        (o.Code != null && o.Code.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) ||
                        (o.Name != null && o.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0));
                }

                var rows = filteredOps.Select(o => new TechOperationRowView
                {
                    IsSelected = _selectedOperationCodes.Contains(o.Code),
                    Code = o.Code,
                    Name = o.Name
                }).ToList();

                gridOperations.DataSource = rows;
            }
            finally
            {
                _isUpdatingOpGrid = false;
            }
        }

        private void GridOperations_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_isUpdatingOpGrid || e.RowIndex < 0 || e.ColumnIndex != 0)
                return;

            if (gridOperations.Rows[e.RowIndex].DataBoundItem is TechOperationRowView rowView)
            {
                if (rowView.IsSelected)
                {
                    _selectedOperationCodes.Add(rowView.Code);
                }
                else
                {
                    _selectedOperationCodes.Remove(rowView.Code);
                }

                UpdateResetOpsButton();
                ApplyTechProcessFilter();
            }
        }

        private void UpdateResetOpsButton()
        {
            int count = _selectedOperationCodes.Count;
            btnResetOps.Enabled = count > 0;
            btnResetOps.Text = count > 0 ? $"Сбросить ({count})" : "Сбросить галочки";
        }

        private void ApplyTechProcessFilter()
        {
            if (_tpData == null || _tpData.Processes == null)
            {
                gridTechProcesses.DataSource = null;
                return;
            }

            IEnumerable<TechProcessItem> candidates = _tpData.Processes;

            if (_selectedOperationCodes.Count > 0)
            {
                // Фильтруем по ВСЕМ отмеченным галочками операциям (включая разные вкладки)
                candidates = candidates.Where(tp =>
                    _selectedOperationCodes.All(reqCode =>
                        tp.OperationCodes.Any(c => string.Equals(c, reqCode, StringComparison.OrdinalIgnoreCase))));
            }
            else
            {
                string query = txtSearchTp.Text.Trim();
                if (!string.IsNullOrEmpty(query))
                {
                    candidates = candidates.Where(tp =>
                        (tp.Designation != null && tp.Designation.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) ||
                        (tp.Name != null && tp.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) ||
                        (tp.Composition != null && tp.Composition.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0));
                }
            }

            var list = candidates.ToList();
            gridTechProcesses.DataSource = list;
            lblTechProcessesTitle.Text = $"  Техпроцессы ({list.Count} из {_tpData.Processes.Count})";

            if (_selectedOperationCodes.Count > 0)
            {
                string selectedList = string.Join(" + ", _selectedOperationCodes);
                lblTpStatus.Text = $"Выбрано операций: {_selectedOperationCodes.Count} [{selectedList}] | Найдено ТП: {list.Count}";
            }
            else
            {
                lblTpStatus.Text = $"Всего техпроцессов: {list.Count}. Отметьте операции галочками для подбора.";
            }
        }

        private void InsertSelectedTechProcess()
        {
            if (gridTechProcesses.CurrentRow?.DataBoundItem is TechProcessItem selectedTp)
            {
                var result = ExcelInsertionService.InsertSingleValue(ExcelApp, selectedTp.Designation, _config);
                lblTpStatus.Text = result.Message;
                if (!result.Success)
                {
                    System.Media.SystemSounds.Beep.Play();
                }
            }
            else
            {
                lblTpStatus.Text = "Выберите техпроцесс в нижней таблице для вставки.";
                System.Media.SystemSounds.Beep.Play();
            }
        }

        private class TechOperationRowView
        {
            public bool IsSelected { get; set; }
            public string Code { get; set; }
            public string Name { get; set; }
        }

        // ================= ЛОГИКА ВКЛАДКИ ШАБЛОНЫ =================

        private void LstTemplates_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= lstTemplates.Items.Count) return;
            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            Color bg = selected ? Color.FromArgb(16, 124, 65) : (e.Index % 2 == 0 ? Color.White : Color.FromArgb(248, 249, 251));
            Color fg = selected ? Color.White : Color.FromArgb(20, 20, 20);
            Color subFg = selected ? Color.FromArgb(220, 240, 225) : Color.FromArgb(120, 130, 140);

            using (var b = new SolidBrush(bg))
            {
                e.Graphics.FillRectangle(b, e.Bounds);
            }

            if (lstTemplates.Items[e.Index] is TemplateItem t)
            {
                using (var fontMain = new Font(e.Font.FontFamily, 9.5F, FontStyle.Bold))
                using (var fontSub = new Font(e.Font.FontFamily, 8.5F, FontStyle.Regular))
                using (var brushFg = new SolidBrush(fg))
                using (var brushSub = new SolidBrush(subFg))
                {
                    var rectMain = new Rectangle(e.Bounds.X + 8, e.Bounds.Y + 4, e.Bounds.Width - 95, e.Bounds.Height - 8);
                    var sf = new StringFormat { LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };
                    e.Graphics.DrawString(t.Name, fontMain, brushFg, rectMain, sf);

                    string countText = $"{t.RowCount} стр.";
                    var rectSub = new Rectangle(e.Bounds.Right - 85, e.Bounds.Y + 4, 75, e.Bounds.Height - 8);
                    var sfRight = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Center };
                    e.Graphics.DrawString(countText, fontSub, brushSub, rectSub, sfRight);
                }
            }
        }

        private void LstTemplates_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstTemplates.SelectedItem is TemplateItem t)
            {
                lblTemplateStatus.Text = $"Выбран: {t.Name} (строки {t.StartRow}–{t.EndRow}, {t.RowCount} строк)";
            }
            else
            {
                lblTemplateStatus.Text = "Выберите шаблон и целевую ячейку в столбце A";
            }
        }

        private void ReloadTemplates()
        {
            try
            {
                lblTemplateStatus.Text = "Чтение файла шаблонов...";
                Application.DoEvents();

                string path = TemplateService.GetTemplatesFilePath(_config);
                _allTemplates = TemplateService.LoadTemplates(path);
                ApplyTemplateFilter();

                TemplateService.EnsureBackgroundWorkbookOpened(ExcelApp, path);

                if (_allTemplates.Count > 0)
                {
                    lblTemplateStatus.Text = $"Загружено шаблонов: {_allTemplates.Count}";
                }
                else
                {
                    lblTemplateStatus.Text = $"Шаблоны не найдены в файле {Path.GetFileName(path)}";
                }
            }
            catch (Exception ex)
            {
                lblTemplateStatus.Text = $"Ошибка загрузки шаблонов: {ex.Message}";
            }
        }

        private void RefreshTemplatesSafely()
        {
            try
            {
                lblTemplateStatus.Text = "Перезагрузка книги шаблонов...";
                Application.DoEvents();

                TemplateService.CloseBackgroundWorkbook();

                string path = TemplateService.GetTemplatesFilePath(_config);
                _allTemplates = TemplateService.LoadTemplates(path);
                ApplyTemplateFilter();

                TemplateService.EnsureBackgroundWorkbookOpened(ExcelApp, path);

                if (_allTemplates.Count > 0)
                {
                    lblTemplateStatus.Text = $"Шаблоны обновлены в {DateTime.Now:HH:mm:ss} (загружено: {_allTemplates.Count})";
                }
                else
                {
                    lblTemplateStatus.Text = $"Шаблоны не найдены в файле {Path.GetFileName(path)}";
                }
            }
            catch (Exception ex)
            {
                lblTemplateStatus.Text = $"Ошибка обновления: {ex.Message}";
            }
        }

        private void ApplyTemplateFilter()
        {
            string query = txtSearchTemplates.Text.Trim();
            IEnumerable<TemplateItem> items = _allTemplates;
            if (!string.IsNullOrEmpty(query))
            {
                items = items.Where(t => t.Name != null && t.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0);
            }

            lstTemplates.BeginUpdate();
            lstTemplates.Items.Clear();
            foreach (var t in items)
            {
                lstTemplates.Items.Add(t);
            }
            if (lstTemplates.Items.Count > 0)
            {
                lstTemplates.SelectedIndex = 0;
            }
            lstTemplates.EndUpdate();
        }

        private void InsertSelectedTemplate()
        {
            if (lstTemplates.SelectedItem is TemplateItem selected)
            {
                string path = TemplateService.GetTemplatesFilePath(_config);
                var result = TemplateService.InsertTemplate(ExcelApp, path, selected);
                lblTemplateStatus.Text = result.Message;
                if (!result.Success)
                {
                    System.Media.SystemSounds.Beep.Play();
                }
            }
            else
            {
                lblTemplateStatus.Text = "Выберите шаблон из списка для вставки.";
                System.Media.SystemSounds.Beep.Play();
            }
        }

        private void OpenTemplateFile()
        {
            string path = TemplateService.GetTemplatesFilePath(_config);
            TemplateService.OpenTemplateFile(ExcelApp, path);
        }

        private void BtnToggleTree_Click(object sender, EventArgs e)
        {
            SetTreeCollapsed(!splitContainer.Panel1Collapsed);
        }

        private void SetTreeCollapsed(bool collapsed)
        {
            splitContainer.Panel1Collapsed = collapsed;
            btnToggleTree.Text = collapsed ? "▼ Показать папки" : "▲ Скрыть папки";

            if (_config != null)
            {
                _config.IsTreeCollapsed = collapsed;
                ConfigManager.Save(_config);
            }
        }

        private void BtnSettings_Click(object sender, EventArgs e)
        {
            using (var form = new SettingsForm(_config))
            {
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    _config = form.CurrentConfig;
                    _watcher.UpdatePaths(_config);
                    ReloadAllCatalogs();
                    if (tabCatalogs.SelectedTab == tabTemplates)
                    {
                        ReloadTemplates();
                    }
                    else if (tabCatalogs.SelectedTab == tabTechProcesses)
                    {
                        ReloadTechProcesses();
                    }
                }
            }
        }

        private void Watcher_FileChanged(CatalogType type, string filePath)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() => Watcher_FileChanged(type, filePath)));
                return;
            }

            LoadSingleCatalog(type, filePath);
            if (CurrentCatalogType == type)
            {
                RefreshCurrentView();
            }

            lblStatus.Text = $"Автосинхронизация: {type} обновлена в {DateTime.Now:HH:mm:ss}";
        }

        private void InsertSelectedItem()
        {
            if (gridItems.CurrentRow?.DataBoundItem is CatalogItem selectedItem)
            {
                var result = ExcelInsertionService.InsertItem(ExcelApp, selectedItem, _config);
                lblStatus.Text = result.Message;
                if (!result.Success)
                {
                    System.Media.SystemSounds.Beep.Play();
                }
            }
            else
            {
                lblStatus.Text = "Выберите строку в таблице для вставки.";
            }
        }

        private void GridItems_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                InsertSelectedItem();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                TemplateService.CloseBackgroundWorkbook();
                _watcher?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}

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

            // Служба автосинхронизации баз данных
            _watcher = new DatabaseWatcher(_config.SyncDebounceMs);
            _watcher.FileChanged += Watcher_FileChanged;
            _watcher.UpdatePaths(_config);

            // Первичная загрузка каталогов
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

            // Строка 1: Поиск и фильтрация
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

            // Строка 2: Кнопки управления
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

            tabCatalogs.TabPages.AddRange(new TabPage[] { tabMaterials, tabHardware, tabComponents });
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
                BackColor = Color.FromArgb(16, 124, 65), // Excel Green
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
            // В вертикальной боковой панели TreeView размещается СВЕРХУ, DataGridView — СНИЗУ.
            // Это гарантирует, что таблица занимает 100% ширины панели и не сплющивается!
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

            // Дерево разделов (TreeView)
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

            // Таблица позиций (DataGridView)
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

            // Стиль заголовков колонок
            gridItems.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(225, 230, 238);
            gridItems.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(20, 20, 20);
            gridItems.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            gridItems.ColumnHeadersHeight = 28;

            // Стиль строк данных: ЧЕТКИЙ ТЕМНЫЙ ТЕКСТ НА БЕЛОМ ФОНЕ
            gridItems.DefaultCellStyle.BackColor = Color.White;
            gridItems.DefaultCellStyle.ForeColor = Color.FromArgb(15, 15, 15);
            gridItems.DefaultCellStyle.SelectionBackColor = Color.FromArgb(16, 124, 65);
            gridItems.DefaultCellStyle.SelectionForeColor = Color.White;

            gridItems.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 247, 251);
            gridItems.AlternatingRowsDefaultCellStyle.ForeColor = Color.FromArgb(15, 15, 15);
            gridItems.AlternatingRowsDefaultCellStyle.SelectionBackColor = Color.FromArgb(16, 124, 65);
            gridItems.AlternatingRowsDefaultCellStyle.SelectionForeColor = Color.White;

            // Колонка 1: Обозначение
            var colDesig = new DataGridViewTextBoxColumn
            {
                HeaderText = "Обозначение",
                DataPropertyName = "Designation",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 45,
                MinimumWidth = 120
            };

            // Колонка 2: Наименование
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

            // Сборка элементов
            this.Controls.Add(splitContainer);
            this.Controls.Add(tabCatalogs);
            this.Controls.Add(pnlHeader);
            this.Controls.Add(pnlFooter);

            this.ResumeLayout(false);
            this.PerformLayout();
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
                    SetSafeSplitterDistance(150);
                    return;
                }

                if (TaskPane.DockPosition == MsoCTPDockPosition.msoCTPDockPositionBottom ||
                    TaskPane.DockPosition == MsoCTPDockPosition.msoCTPDockPositionTop)
                {
                    // Если панель закреплена снизу (широкая) - дерево слева, таблица справа
                    splitContainer.Orientation = Orientation.Vertical;
                    SetSafeSplitterDistance(220);
                }
                else
                {
                    // Если панель сбоку (высокая и узкая) - дерево сверху, таблица снизу во всю ширину!
                    splitContainer.Orientation = Orientation.Horizontal;
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

                if (total <= 50) return; // Контейнер еще не получил реальный размер от Excel

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
            txtSearch.Clear();
            RefreshCurrentView();
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
                _watcher?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}

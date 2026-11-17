using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using SpecHelper.Models;
using SpecHelper.Services;
using SpecHelper.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace SpecHelper
{
    [ComVisible(true)]
    [Guid("A19845E6-5F24-4B78-8316-D4E03C1247AA")]
    [ProgId("SpecHelper.AddIn")]
    public class AddIn : IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
    {
        private static readonly string LogFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "SpecHelper", "addin_debug.log"
        );

        public static void Log(string message)
        {
            try
            {
                var dir = Path.GetDirectoryName(LogFile);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                File.AppendAllText(LogFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}\r\n");
            }
            catch { }
        }

        static AddIn()
        {
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                try
                {
                    var requestedName = new System.Reflection.AssemblyName(args.Name).Name;
                    var assemblyDir = System.IO.Path.GetDirectoryName(typeof(AddIn).Assembly.Location);
                    if (!string.IsNullOrEmpty(assemblyDir))
                    {
                        var candidateFile = System.IO.Path.Combine(assemblyDir, requestedName + ".dll");
                        if (System.IO.File.Exists(candidateFile))
                        {
                            return System.Reflection.Assembly.LoadFrom(candidateFile);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"AssemblyResolve error for {args.Name}: {ex.Message}");
                }
                return null;
            };
        }

        private Excel.Application _excelApp;
        private CustomTaskPane _taskPane;
        private Form _floatingForm;
        private SpecPaneControl _paneControl;
        private ICTPFactory _ctpFactory;
        private IRibbonUI _ribbon;

        #region IDTExtensibility2 Members

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            Log("OnConnection start");
            _excelApp = (Excel.Application)Application;

            try
            {
                var appEvents = (Excel.AppEvents_Event)_excelApp;
                appEvents.WorkbookOpen += (Excel.Workbook wb) =>
                {
                    Log($"WorkbookOpen: {wb?.Name}");
                    EnsureTaskPaneCreatedAndVisible();
                };
                appEvents.NewWorkbook += (Excel.Workbook wb) =>
                {
                    Log($"NewWorkbook: {wb?.Name}");
                    EnsureTaskPaneCreatedAndVisible();
                };
                appEvents.WindowActivate += (Excel.Workbook wb, Excel.Window wn) =>
                {
                    Log($"WindowActivate: {wb?.Name}");
                    EnsureTaskPaneCreatedAndVisible();
                };
            }
            catch (Exception ex)
            {
                Log($"Error hooking events: {ex.Message}");
            }

            Log("OnConnection completed");
        }

        public void OnStartupComplete(ref Array custom)
        {
            Log("OnStartupComplete start");
            EnsureTaskPaneCreatedAndVisible();
            Log("OnStartupComplete end");
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Log("OnDisconnection");
            if (_taskPane != null)
            {
                try { _taskPane.Visible = false; } catch { }
            }

            if (_floatingForm != null && !_floatingForm.IsDisposed)
            {
                try { _floatingForm.Close(); _floatingForm.Dispose(); } catch { }
                _floatingForm = null;
            }

            _paneControl?.Dispose();
            _paneControl = null;
            _taskPane = null;
            _excelApp = null;
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Log("OnBeginShutdown");
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        #endregion

        #region ICustomTaskPaneConsumer Members

        public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
        {
            Log("CTPFactoryAvailable received");
            _ctpFactory = CTPFactoryInst;
            EnsureTaskPaneCreatedAndVisible();
        }

        private void EnsureTaskPaneCreatedAndVisible()
        {
            if (_taskPane == null && (_floatingForm == null || _floatingForm.IsDisposed))
            {
                bool ctpCreated = false;
                if (_ctpFactory != null)
                {
                    try
                    {
                        Log("Creating CustomTaskPane via CTPFactory...");
                        _taskPane = _ctpFactory.CreateCTP("SpecHelper.UI.SpecPaneControl", "Справочник спецификаций", Type.Missing);
                        Log("CreateCTP succeeded!");

                        _taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                        _taskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;

                        try
                        {
                            _taskPane.Width = 560;
                        }
                        catch (Exception wEx)
                        {
                            Log($"Width set error: {wEx.Message}");
                        }

                        _taskPane.DockPositionStateChange += (CustomTaskPane ctp) =>
                        {
                            try
                            {
                                if (ctp.DockPosition == MsoCTPDockPosition.msoCTPDockPositionRight ||
                                    ctp.DockPosition == MsoCTPDockPosition.msoCTPDockPositionLeft)
                                {
                                    try { ctp.Width = 560; } catch { }
                                }
                                else if (ctp.DockPosition == MsoCTPDockPosition.msoCTPDockPositionBottom ||
                                         ctp.DockPosition == MsoCTPDockPosition.msoCTPDockPositionTop)
                                {
                                    try { ctp.Height = 360; } catch { }
                                }

                                _paneControl?.OnDockPositionChanged();
                            }
                            catch { }
                        };

                        _paneControl = _taskPane.ContentControl as SpecPaneControl;
                        if (_paneControl != null && _excelApp != null)
                        {
                            Log("Initializing SpecPaneControl in CustomTaskPane...");
                            _paneControl.Initialize(_excelApp, _taskPane);
                            Log("SpecPaneControl initialized successfully");
                        }

                        ctpCreated = true;
                    }
                    catch (Exception ex)
                    {
                        Log($"CreateCTP failed or blocked by policy: {ex.Message}. Switching to fallback floating window.");
                        _taskPane = null;
                        ctpCreated = false;
                    }
                }
                else
                {
                    Log("CTPFactory is null. Switching to fallback floating window.");
                }

                if (!ctpCreated)
                {
                    CreateOrShowFloatingWindow();
                    return;
                }
            }

            if (_taskPane != null)
            {
                try
                {
                    if (!_taskPane.Visible)
                    {
                        Log("Setting _taskPane.Visible = true");
                        _taskPane.Visible = true;
                        Log("_taskPane.Visible is now true");
                    }
                }
                catch (Exception vEx)
                {
                    Log($"Error setting Visible = true on CTP: {vEx.Message}. Attempting floating window fallback.");
                    CreateOrShowFloatingWindow();
                }
            }
            else if (_floatingForm != null && !_floatingForm.IsDisposed)
            {
                CreateOrShowFloatingWindow();
            }
        }

        private void CreateOrShowFloatingWindow()
        {
            try
            {
                if (_floatingForm == null || _floatingForm.IsDisposed)
                {
                    Log("Creating fallback floating window (Form TopMost)...");
                    _floatingForm = new Form
                    {
                        Text = "Справочник спецификаций",
                        Width = 580,
                        Height = 720,
                        MinimumSize = new System.Drawing.Size(350, 400),
                        StartPosition = FormStartPosition.CenterScreen,
                        TopMost = true,
                        ShowIcon = false
                    };

                    _paneControl = new SpecPaneControl
                    {
                        Dock = DockStyle.Fill
                    };
                    _floatingForm.Controls.Add(_paneControl);

                    _floatingForm.FormClosing += (s, e) =>
                    {
                        if (e.CloseReason == CloseReason.UserClosing)
                        {
                            e.Cancel = true;
                            _floatingForm.Hide();
                        }
                    };

                    if (_excelApp != null)
                    {
                        _paneControl.Initialize(_excelApp, null);
                    }
                    Log("Fallback floating window created successfully.");
                }

                if (!_floatingForm.Visible)
                {
                    _floatingForm.Show();
                }
                _floatingForm.BringToFront();
            }
            catch (Exception ex)
            {
                Log($"Error in CreateOrShowFloatingWindow: {ex.Message}");
            }
        }

        #endregion

        #region IRibbonExtensibility Members

        public string GetCustomUI(string RibbonID)
        {
            Log($"GetCustomUI requested for {RibbonID}");
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab id=""tabSpecHelper"" label=""Спецификация"">
        <group id=""grpSpecHelper"" label=""Справочник БД"">
          <button id=""btnTogglePane"" label=""Справочник БД"" size=""large"" onAction=""OnTogglePaneClicked"" imageMso=""TableInsert"" supertip=""Открыть или скрыть боковую панель справочника"" />
          <separator id=""sep1"" />
          <button id=""btnRefresh"" label=""Обновить базы"" size=""normal"" onAction=""OnRefreshClicked"" imageMso=""Refresh"" supertip=""Перечитать файлы баз данных с диска"" />
          <button id=""btnSettings"" label=""Настройки"" size=""normal"" onAction=""OnSettingsClicked"" imageMso=""ServerSettings"" supertip=""Настроить пути к файлам баз данных"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            Log("OnRibbonLoad");
            _ribbon = ribbonUI;
        }

        public void OnTogglePaneClicked(IRibbonControl control)
        {
            Log("OnTogglePaneClicked");
            TogglePane();
        }

        public void OnRefreshClicked(IRibbonControl control)
        {
            Log("OnRefreshClicked");
            _paneControl?.ReloadAllCatalogs();
        }

        public void OnSettingsClicked(IRibbonControl control)
        {
            Log("OnSettingsClicked");
            var config = ConfigManager.Load();
            using (var form = new SettingsForm(config))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    _paneControl?.ReloadAllCatalogs();
                }
            }
        }

        public void TogglePane()
        {
            try
            {
                if (_taskPane != null)
                {
                    _taskPane.Visible = !_taskPane.Visible;
                    Log($"TogglePane: CTP visible now {_taskPane.Visible}");
                }
                else if (_floatingForm != null && !_floatingForm.IsDisposed)
                {
                    if (_floatingForm.Visible)
                    {
                        _floatingForm.Hide();
                    }
                    else
                    {
                        _floatingForm.Show();
                        _floatingForm.BringToFront();
                    }
                    Log($"TogglePane: Floating form visible now {_floatingForm.Visible}");
                }
                else
                {
                    EnsureTaskPaneCreatedAndVisible();
                }
            }
            catch (Exception ex)
            {
                Log($"TogglePane error: {ex.Message}");
            }
        }

        #endregion
    }
}

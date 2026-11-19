using System;
using System.Collections.Generic;
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
        private readonly Dictionary<int, CustomTaskPane> _taskPanes = new Dictionary<int, CustomTaskPane>();
        private readonly Dictionary<int, SpecPaneControl> _paneControls = new Dictionary<int, SpecPaneControl>();
        private Form _floatingForm;
        private SpecPaneControl _floatingPaneControl;
        private ICTPFactory _ctpFactory;
        private IRibbonUI _ribbon;

        #region IDTExtensibility2 Members

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            Log("OnConnection start");
            _excelApp = (Excel.Application)Application;
            Log("OnConnection completed");
        }

        public void OnStartupComplete(ref Array custom)
        {
            Log("OnStartupComplete start");
            UpdateService.CheckForUpdatesAsync(_excelApp);
            Log("OnStartupComplete end");
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Log("OnDisconnection");
            try { TemplateService.CloseBackgroundWorkbook(); } catch { }

            foreach (var pane in _taskPanes.Values)
            {
                try { pane.Visible = false; } catch { }
            }
            _taskPanes.Clear();

            foreach (var ctrl in _paneControls.Values)
            {
                try { ctrl?.Dispose(); } catch { }
            }
            _paneControls.Clear();

            if (_floatingForm != null && !_floatingForm.IsDisposed)
            {
                try { _floatingForm.Close(); _floatingForm.Dispose(); } catch { }
                _floatingForm = null;
            }

            _floatingPaneControl?.Dispose();
            _floatingPaneControl = null;
            _excelApp = null;
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Log("OnBeginShutdown");
            try { TemplateService.CloseBackgroundWorkbook(); } catch { }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        #endregion

        #region ICustomTaskPaneConsumer Members

        public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
        {
            Log("CTPFactoryAvailable received (on-demand mode)");
            _ctpFactory = CTPFactoryInst;
        }

        private void CreateAndShowPaneForActiveWindow()
        {
            if (_excelApp == null) return;

            dynamic activeWindow = _excelApp.ActiveWindow;
            if (activeWindow == null)
            {
                Log("ActiveWindow is null. Switching to fallback floating window.");
                CreateOrShowFloatingWindow();
                return;
            }

            int hwnd = (int)activeWindow.Hwnd;

            if (_ctpFactory != null)
            {
                try
                {
                    Log($"Creating CustomTaskPane for active window Hwnd={hwnd}...");
                    CustomTaskPane pane = _ctpFactory.CreateCTP(
                        "SpecHelper.UI.SpecPaneControl",
                        "Справочник спецификаций",
                        activeWindow
                    );
                    Log("CreateCTP succeeded!");

                    pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                    pane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;

                    try
                    {
                        pane.Width = 560;
                    }
                    catch (Exception wEx)
                    {
                        Log($"Width set error: {wEx.Message}");
                    }

                    SpecPaneControl control = pane.ContentControl as SpecPaneControl;
                    if (control != null)
                    {
                        Log("Initializing SpecPaneControl in CustomTaskPane...");
                        control.Initialize(_excelApp, pane);
                        Log("SpecPaneControl initialized successfully");
                    }

                    pane.DockPositionStateChange += (CustomTaskPane ctp) =>
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

                            control?.OnDockPositionChanged();
                        }
                        catch { }
                    };

                    _taskPanes[hwnd] = pane;
                    _paneControls[hwnd] = control;

                    pane.Visible = true;
                    Log($"CustomTaskPane visible is now true for Hwnd={hwnd}");
                    return;
                }
                catch (Exception ex)
                {
                    Log($"CreateCTP failed or blocked by policy: {ex.Message}. Switching to fallback floating window.");
                    _taskPanes.Remove(hwnd);
                    _paneControls.Remove(hwnd);
                }
            }
            else
            {
                Log("CTPFactory is null. Switching to fallback floating window.");
            }

            CreateOrShowFloatingWindow();
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

                    _floatingPaneControl = new SpecPaneControl
                    {
                        Dock = DockStyle.Fill
                    };
                    _floatingForm.Controls.Add(_floatingPaneControl);

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
                        _floatingPaneControl.Initialize(_excelApp, null);
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

        private void ReloadAllOpenCatalogs()
        {
            foreach (var ctrl in _paneControls.Values)
            {
                try { ctrl?.ReloadAllCatalogs(); } catch { }
            }
            try { _floatingPaneControl?.ReloadAllCatalogs(); } catch { }
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
            ReloadAllOpenCatalogs();
        }

        public void OnSettingsClicked(IRibbonControl control)
        {
            Log("OnSettingsClicked");
            var config = ConfigManager.Load();
            using (var form = new SettingsForm(config))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    ReloadAllOpenCatalogs();
                }
            }
        }

        public void TogglePane()
        {
            try
            {
                if (_excelApp != null && _excelApp.ActiveWindow != null)
                {
                    dynamic activeWindow = _excelApp.ActiveWindow;
                    int hwnd = (int)activeWindow.Hwnd;

                    if (_taskPanes.TryGetValue(hwnd, out CustomTaskPane existingPane))
                    {
                        try
                        {
                            existingPane.Visible = !existingPane.Visible;
                            Log($"TogglePane: Hwnd={hwnd}, CTP visible now {existingPane.Visible}");
                            return;
                        }
                        catch
                        {
                            _taskPanes.Remove(hwnd);
                            _paneControls.Remove(hwnd);
                        }
                    }

                    CreateAndShowPaneForActiveWindow();
                    return;
                }

                if (_floatingForm != null && !_floatingForm.IsDisposed)
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
                    CreateOrShowFloatingWindow();
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

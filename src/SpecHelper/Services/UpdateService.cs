using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace SpecHelper.Services
{
    public static class UpdateService
    {
        public const string DefaultVersionJsonUrl = "https://raw.githubusercontent.com/pururuall/SpecHelper/main/version.json";
        public static string VersionJsonUrl = DefaultVersionJsonUrl;

        public class UpdateManifest
        {
            [JsonProperty("version")]
            public string version { get; set; }

            [JsonProperty("url")]
            public string url { get; set; }

            [JsonProperty("notes")]
            public string notes { get; set; }
        }

        public static void CheckForUpdatesAsync(Excel.Application excelApp)
        {
            Task.Run(() =>
            {
                try
                {
                    CheckForUpdates(excelApp);
                }
                catch (Exception ex)
                {
                    AddIn.Log($"Silent update check failed: {ex.Message}");
                }
            });
        }

        public static bool CheckForUpdates(Excel.Application excelApp, string manifestUrl = null, bool? autoConfirm = null)
        {
            string url = manifestUrl ?? VersionJsonUrl;
            if (string.IsNullOrWhiteSpace(url))
            {
                return false;
            }

            try
            {
                AddIn.Log($"Checking for updates from {url}...");

                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

                string jsonContent = null;
                using (var client = new TimeoutWebClient(5000))
                {
                    client.Headers[HttpRequestHeader.UserAgent] = "SpecHelper-Updater";
                    client.Encoding = System.Text.Encoding.UTF8;
                    jsonContent = client.DownloadString(url);
                }

                if (string.IsNullOrWhiteSpace(jsonContent))
                {
                    AddIn.Log("Update manifest is empty.");
                    return false;
                }

                var manifest = ParseManifest(jsonContent);
                if (manifest == null || string.IsNullOrWhiteSpace(manifest.version) || string.IsNullOrWhiteSpace(manifest.url))
                {
                    AddIn.Log("Failed to parse update manifest.");
                    return false;
                }

                if (!Version.TryParse(manifest.version, out var remoteVersion))
                {
                    AddIn.Log($"Invalid remote version format: {manifest.version}");
                    return false;
                }

                var currentVersion = GetCurrentAssemblyFileVersion();
                AddIn.Log($"Current version: {currentVersion}, Remote version: {remoteVersion}");

                if (remoteVersion <= currentVersion)
                {
                    AddIn.Log("No update required. Current version is up-to-date.");
                    return false;
                }

                // Новая версия найдена, запускаем сценарий обновления
                HandleUpdateAvailable(excelApp, manifest, currentVersion, remoteVersion, autoConfirm);
                return true;
            }
            catch (Exception ex)
            {
                AddIn.Log($"Update check exception (silent): {ex.Message}");
                return false;
            }
        }

        public static Version GetCurrentAssemblyFileVersion()
        {
            try
            {
                var asm = typeof(UpdateService).Assembly;
                var attr = (AssemblyFileVersionAttribute)Attribute.GetCustomAttribute(asm, typeof(AssemblyFileVersionAttribute));
                if (attr != null && Version.TryParse(attr.Version, out var v))
                {
                    return v;
                }

                var fvi = FileVersionInfo.GetVersionInfo(asm.Location);
                if (!string.IsNullOrEmpty(fvi.FileVersion) && Version.TryParse(fvi.FileVersion, out var fv))
                {
                    return fv;
                }
            }
            catch (Exception ex)
            {
                AddIn.Log($"Error reading AssemblyFileVersion: {ex.Message}");
            }

            return new Version(1, 0, 0, 0);
        }

        private static UpdateManifest ParseManifest(string json)
        {
            try
            {
                var manifest = JsonConvert.DeserializeObject<UpdateManifest>(json);
                if (manifest != null && !string.IsNullOrWhiteSpace(manifest.version))
                {
                    return manifest;
                }
            }
            catch { }

            try
            {
                var matchVer = Regex.Match(json, @"""version""\s*:\s*""([^""]+)""", RegexOptions.IgnoreCase);
                var matchUrl = Regex.Match(json, @"""url""\s*:\s*""([^""]+)""", RegexOptions.IgnoreCase);
                var matchNotes = Regex.Match(json, @"""notes""\s*:\s*""([^""]+)""", RegexOptions.IgnoreCase);

                if (matchVer.Success && matchUrl.Success)
                {
                    return new UpdateManifest
                    {
                        version = matchVer.Groups[1].Value.Trim(),
                        url = matchUrl.Groups[1].Value.Trim(),
                        notes = matchNotes.Success ? matchNotes.Groups[1].Value : ""
                    };
                }
            }
            catch { }

            return null;
        }

        private static void HandleUpdateAvailable(Excel.Application excelApp, UpdateManifest manifest, Version currentVersion, Version remoteVersion, bool? autoConfirm = null)
        {
            DialogResult result;

            if (autoConfirm.HasValue)
            {
                result = autoConfirm.Value ? DialogResult.Yes : DialogResult.No;
            }
            else
            {
                string message = $"Доступно обновление надстройки SpecHelper!\r\n\r\n" +
                                 $"Текущая версия: {currentVersion}\r\n" +
                                 $"Новая версия: {remoteVersion}\r\n\r\n" +
                                 (string.IsNullOrWhiteSpace(manifest.notes) ? "" : $"Описание изменений:\r\n{manifest.notes}\r\n\r\n") +
                                 "Чтобы установить обновление, необходимо закрыть Excel.\r\n" +
                                 "Закрыть Excel и обновить сейчас?";

                IntPtr hwnd = IntPtr.Zero;
                try
                {
                    if (excelApp != null)
                    {
                        hwnd = new IntPtr(excelApp.Hwnd);
                    }
                }
                catch { }

                if (hwnd != IntPtr.Zero)
                {
                    result = MessageBox.Show(new WindowOwner(hwnd), message, "Обновление SpecHelper", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                }
                else
                {
                    result = MessageBox.Show(message, "Обновление SpecHelper", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                }
            }

            if (result != DialogResult.Yes)
            {
                AddIn.Log("User postponed update.");
                return;
            }

            ExecuteUpdate(excelApp, manifest);
        }

        private static void ExecuteUpdate(Excel.Application excelApp, UpdateManifest manifest)
        {
            try
            {
                string tempDir = Path.Combine(Path.GetTempPath(), "SpecHelper_Update");
                if (Directory.Exists(tempDir))
                {
                    try { Directory.Delete(tempDir, true); } catch { }
                }
                Directory.CreateDirectory(tempDir);

                string zipPath = Path.Combine(tempDir, "update.zip");
                string extractedPath = Path.Combine(tempDir, "extracted");

                AddIn.Log($"Downloading update archive from {manifest.url}...");
                using (var client = new TimeoutWebClient(60000))
                {
                    client.Headers[HttpRequestHeader.UserAgent] = "SpecHelper-Updater";
                    client.DownloadFile(manifest.url, zipPath);
                }

                AddIn.Log($"Extracting archive to {extractedPath}...");
                ZipFile.ExtractToDirectory(zipPath, extractedPath);

                string targetDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "SpecHelper"
                );

                string batPath = Path.Combine(tempDir, "apply_update.bat");
                string batContent =
@"@echo off
setlocal EnableDelayedExpansion
title Obnovlenie SpecHelper...
set ""LOGFILE=%~dp0update.log""
echo [START] %DATE% %TIME% > ""%LOGFILE%""
echo ========================================================
echo           OBNOVLENIE SPECHELPER DLYA EXCEL
echo ========================================================
echo.
echo Ozhidanie zakrytiya Excel...

set /a WAIT_COUNT=0
:wait_loop
tasklist /FI ""IMAGENAME eq EXCEL.EXE"" 2>NUL | find /I /N ""EXCEL.EXE"">NUL
if ""%ERRORLEVEL%""==""0"" (
    set /a WAIT_COUNT+=1
    if !WAIT_COUNT! geq 12 (
        echo Prinuditelnoe zakrytie Excel... >> ""%LOGFILE%""
        taskkill /F /IM EXCEL.EXE >nul 2>&1
    )
    timeout /t 1 /nobreak >nul
    goto wait_loop
)

echo Primenenie obnovleniya... >> ""%LOGFILE%""
set ""TARGET_DIR=" + targetDir + @"""
set ""EXTRACTED_DIR=" + extractedPath + @"""

if exist ""%EXTRACTED_DIR%\bin\SpecHelper.dll"" set ""EXTRACTED_DIR=%EXTRACTED_DIR%\bin""
if exist ""%EXTRACTED_DIR%\SpecHelper_РабочийПК\bin\SpecHelper.dll"" set ""EXTRACTED_DIR=%EXTRACTED_DIR%\SpecHelper_РабочийПК\bin""

echo Kopirovanie bibliotek iz %EXTRACTED_DIR% v %TARGET_DIR%... >> ""%LOGFILE%""
xcopy ""%EXTRACTED_DIR%\*"" ""%TARGET_DIR%\"" /E /Y /I >> ""%LOGFILE%"" 2>&1

if exist ""%TARGET_DIR%\mscorlib.dll"" del /f /q ""%TARGET_DIR%\mscorlib.dll"" >nul 2>&1
if exist ""%TARGET_DIR%\System.*.dll"" del /f /q ""%TARGET_DIR%\System.*.dll"" >nul 2>&1

echo Snyatie blokirovki Windows... >> ""%LOGFILE%""
powershell -NoProfile -ExecutionPolicy Bypass -Command ""Get-ChildItem -Path '%TARGET_DIR%' -Recurse | Unblock-File"" >> ""%LOGFILE%"" 2>&1

echo Zapusk Excel... >> ""%LOGFILE%""
start """" excel.exe

echo [SUCCESS] Obnovlenie zaversheno. >> ""%LOGFILE%""
timeout /t 2 /nobreak >nul
cd /d ""%TEMP%""
rmdir /s /q """ + tempDir + @""" >nul 2>&1
exit
";

                File.WriteAllText(batPath, batContent, System.Text.Encoding.Default);

                AddIn.Log($"Starting updater script: {batPath}");
                var psi = new ProcessStartInfo
                {
                    FileName = batPath,
                    UseShellExecute = true,
                    WorkingDirectory = tempDir
                };
                Process.Start(psi);

                AddIn.Log("Requesting Excel exit...");
                try
                {
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                    }
                }
                catch (Exception ex)
                {
                    AddIn.Log($"excelApp.Quit error: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                AddIn.Log($"ExecuteUpdate error: {ex}");
                MessageBox.Show($"Не удалось выполнить обновление:\r\n{ex.Message}", "Ошибка обновления", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private class TimeoutWebClient : WebClient
        {
            private readonly int _timeoutMs;

            public TimeoutWebClient(int timeoutMs = 5000)
            {
                _timeoutMs = timeoutMs;
            }

            protected override WebRequest GetWebRequest(Uri address)
            {
                var request = base.GetWebRequest(address);
                if (request != null)
                {
                    request.Timeout = _timeoutMs;
                }
                return request;
            }
        }

        private class WindowOwner : IWin32Window
        {
            private readonly IntPtr _hwnd;
            public WindowOwner(IntPtr handle) { _hwnd = handle; }
            public IntPtr Handle => _hwnd;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using ExcelDataReader;
using SpecHelper.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace SpecHelper.Services
{
    public static class TemplateService
    {
        private const int VBA_E_IGNORE = unchecked((int)0x800AC472);

        private static Excel.Workbook _cachedWorkbook = null;
        private static bool _isOpenedByUs = false;
        private static readonly object _syncLock = new object();

        public static string GetTemplatesFilePath(PluginConfig config)
        {
            if (config != null && !string.IsNullOrWhiteSpace(config.TemplatesPath) && File.Exists(config.TemplatesPath))
            {
                return config.TemplatesPath;
            }

            string[] searchFolders = new string[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SpecHelper", "databases"),
                Path.Combine(Path.GetDirectoryName(typeof(TemplateService).Assembly.Location) ?? "", "databases"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SpecHelper", "databases"),
                @"D:\EXCEL mibil\databases"
            };

            foreach (var folder in searchFolders)
            {
                if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder)) continue;
                string candidate = Path.Combine(folder, "Шаблоны.xlsx");
                if (File.Exists(candidate))
                {
                    if (config != null)
                    {
                        config.TemplatesPath = candidate;
                        ConfigManager.Save(config);
                    }
                    return candidate;
                }
            }

            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SpecHelper", "databases", "Шаблоны.xlsx");
        }

        public static List<TemplateItem> LoadTemplates(string filePath)
        {
            var templates = new List<TemplateItem>();
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return templates;
            }

            FileStream stream = null;
            int retries = 4;
            int delayMs = 250;

            for (int i = 0; i < retries; i++)
            {
                try
                {
                    stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    break;
                }
                catch (IOException)
                {
                    if (i == retries - 1)
                        return templates;
                    Thread.Sleep(delayMs);
                }
            }

            if (stream == null) return templates;

            try
            {
                using (stream)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = false
                        }
                    });

                    if (result.Tables.Count == 0) return templates;
                    var table = result.Tables[0];
                    int totalRows = table.Rows.Count;

                    for (int r = 0; r < totalRows; r++)
                    {
                        var row = table.Rows[r];
                        string colA = row[0]?.ToString()?.Trim() ?? string.Empty;
                        string colB = row.Table.Columns.Count > 1 ? (row[1]?.ToString()?.Trim() ?? string.Empty) : string.Empty;

                        if (colA == "#" && !string.IsNullOrEmpty(colB))
                        {
                            string templateName = colB;
                            int headerRow = r + 1; // 1-based index in Excel

                            // Пропускаем пустые строки после заголовка (по схеме 2 пустые строки)
                            int startR = r + 1;
                            while (startR < totalRows && IsRowEmpty(table.Rows[startR]))
                            {
                                startR++;
                            }

                            if (startR >= totalRows) break;

                            int endR = startR;
                            int currR = startR;
                            while (currR < totalRows)
                            {
                                var cRow = table.Rows[currR];
                                string cA = cRow[0]?.ToString()?.Trim() ?? string.Empty;
                                if (cA == "#")
                                {
                                    break;
                                }

                                if (IsRowEmpty(cRow))
                                {
                                    // 2 последовательные пустые строки означают окончание блока шаблона
                                    if (currR + 1 < totalRows && IsRowEmpty(table.Rows[currR + 1]))
                                    {
                                        break;
                                    }
                                }
                                else
                                {
                                    endR = currR;
                                }

                                currR++;
                            }

                            templates.Add(new TemplateItem
                            {
                                Name = templateName,
                                HeaderRow = headerRow,
                                StartRow = startR + 1, // 1-based index in Excel
                                EndRow = endR + 1      // 1-based index in Excel
                            });

                            r = endR;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TemplateService] Ошибка парсинга шаблонов: {ex.Message}");
            }

            return templates;
        }

        private static bool IsRowEmpty(DataRow row)
        {
            if (row == null) return true;
            for (int c = 0; c < row.Table.Columns.Count; c++)
            {
                var val = row[c]?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(val)) return false;
            }
            return true;
        }

        public static void EnsureBackgroundWorkbookOpened(Excel.Application excelApp, string templatePath)
        {
            if (excelApp == null || string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
                return;

            try
            {
                GetOrOpenTemplateWorkbook(excelApp, templatePath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TemplateService] Ошибка предзагрузки фоновой книги: {ex.Message}");
            }
        }

        public static Excel.Workbook GetOrOpenTemplateWorkbook(Excel.Application excelApp, string templatePath)
        {
            lock (_syncLock)
            {
                if (_cachedWorkbook != null)
                {
                    try
                    {
                        // Проверка доступности COM-объекта
                        string name = _cachedWorkbook.Name;
                        return _cachedWorkbook;
                    }
                    catch
                    {
                        _cachedWorkbook = null;
                        _isOpenedByUs = false;
                    }
                }

                if (excelApp == null || string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
                    return null;

                // Проверяем, открыта ли книга уже в Excel (например, самим пользователем)
                string templateFileName = Path.GetFileName(templatePath);
                try
                {
                    foreach (Excel.Workbook wb in excelApp.Workbooks)
                    {
                        if (string.Equals(wb.FullName, templatePath, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(wb.Name, templateFileName, StringComparison.OrdinalIgnoreCase))
                        {
                            _cachedWorkbook = wb;
                            _isOpenedByUs = false;
                            return _cachedWorkbook;
                        }
                    }
                }
                catch { }

                // Открываем скрыто в фоновом режиме (ReadOnly: true, Visible = false)
                try
                {
                    _cachedWorkbook = excelApp.Workbooks.Open(
                        templatePath,
                        UpdateLinks: false,
                        ReadOnly: true
                    );
                    _isOpenedByUs = true;

                    try
                    {
                        if (_cachedWorkbook.Windows.Count > 0)
                        {
                            _cachedWorkbook.Windows[1].Visible = false;
                        }
                    }
                    catch { }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[TemplateService] Ошибка фонового открытия книги шаблонов: {ex.Message}");
                    _cachedWorkbook = null;
                    _isOpenedByUs = false;
                }

                return _cachedWorkbook;
            }
        }

        public static void CloseBackgroundWorkbook()
        {
            lock (_syncLock)
            {
                if (_cachedWorkbook != null)
                {
                    try
                    {
                        if (_isOpenedByUs)
                        {
                            _cachedWorkbook.Close(false);
                        }
                    }
                    catch { }
                    finally
                    {
                        try
                        {
                            Marshal.ReleaseComObject(_cachedWorkbook);
                        }
                        catch { }
                        _cachedWorkbook = null;
                        _isOpenedByUs = false;
                    }
                }
            }
        }

        public static (bool Success, string Message) InsertTemplate(Excel.Application excelApp, string templatePath, TemplateItem template)
        {
            if (excelApp == null)
            {
                return (false, "Excel не подключен.");
            }

            if (template == null)
            {
                return (false, "Шаблон не выбран. Выберите шаблон из списка.");
            }

            if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
            {
                return (false, $"Файл шаблонов не найден:\n{templatePath}");
            }

            try
            {
                if (!excelApp.Ready)
                {
                    return (false, "Excel занят. Завершите текущую операцию в Excel.");
                }

                Excel.Range targetCell = null;
                try
                {
                    targetCell = excelApp.ActiveCell;
                }
                catch (COMException comEx) when (comEx.ErrorCode == VBA_E_IGNORE)
                {
                    return (false, "Excel находится в режиме редактирования ячейки. Нажмите Enter или Esc на листе и повторите.");
                }

                if (targetCell == null)
                {
                    return (false, "На листе Excel не выбрана активная ячейка. Выделите ячейку в столбце A.");
                }

                Excel.Worksheet targetWs = targetCell.Worksheet;
                if (targetWs != null && targetWs.ProtectContents)
                {
                    return (false, "Лист Excel защищен от изменений. Снимите защиту листа.");
                }

                // Быстрое получение уже открытой фоновой книги (без задержки открытия с диска)
                Excel.Workbook templateWb = GetOrOpenTemplateWorkbook(excelApp, templatePath);
                if (templateWb == null)
                {
                    return (false, "Не удалось получить доступ к книге шаблонов.");
                }

                try
                {
                    // Отключаем обновление экрана для исключения визуальных рывков
                    try
                    {
                        excelApp.ScreenUpdating = false;
                    }
                    catch { }

                    Excel.Worksheet templateWs = (Excel.Worksheet)templateWb.Worksheets[1];
                    Excel.Range templateRowsRange = templateWs.Range[
                        templateWs.Rows[template.StartRow],
                        templateWs.Rows[template.EndRow]
                    ];

                    // 1. Копируем строки шаблона целиком
                    templateRowsRange.EntireRow.Copy();

                    // 2. Вставляем со сдвигом существующих данных вниз
                    targetCell.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

                    // 3. Снимаем бегущую рамку выделения
                    try
                    {
                        excelApp.CutCopyMode = (Excel.XlCutCopyMode)0;
                    }
                    catch { }

                    return (true, $"Вставлен шаблон «{template.Name}» ({template.RowCount} строк)");
                }
                finally
                {
                    // Гарантированно возвращаем обновление экрана
                    try
                    {
                        excelApp.ScreenUpdating = true;
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                return (false, $"Ошибка при вставке шаблона: {ex.Message}");
            }
        }

        public static void OpenTemplateFile(Excel.Application excelApp, string templatePath)
        {
            if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
            {
                MessageBox.Show($"Файл шаблонов не найден по пути:\n{templatePath}", "Шаблоны SpecHelper", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // Безопасно закрываем фоновую скрытую книгу, чтобы пользователь мог открыть файл для редактирования
                CloseBackgroundWorkbook();

                if (excelApp != null)
                {
                    foreach (Excel.Workbook wb in excelApp.Workbooks)
                    {
                        if (string.Equals(wb.FullName, templatePath, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(wb.Name, Path.GetFileName(templatePath), StringComparison.OrdinalIgnoreCase))
                        {
                            try
                            {
                                if (wb.Windows.Count > 0)
                                {
                                    wb.Windows[1].Visible = true;
                                }
                            }
                            catch { }
                            wb.Activate();
                            excelApp.Visible = true;
                            return;
                        }
                    }

                    excelApp.Workbooks.Open(templatePath, ReadOnly: false);
                    excelApp.Visible = true;
                    return;
                }
            }
            catch { }

            try
            {
                Process.Start(new ProcessStartInfo(templatePath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось открыть файл шаблонов: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}

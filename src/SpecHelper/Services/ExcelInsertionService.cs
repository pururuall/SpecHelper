using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SpecHelper.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace SpecHelper.Services
{
    public static class ExcelInsertionService
    {
        private const int VBA_E_IGNORE = unchecked((int)0x800AC472);

        public static (bool Success, string Message) InsertItem(Excel.Application excelApp, CatalogItem item, PluginConfig config)
        {
            if (excelApp == null)
            {
                return (false, "Приложение Excel не подключено.");
            }

            if (item == null)
            {
                return (false, "Позиция не выбрана.");
            }

            try
            {
                // Проверка, что Excel не в модальном диалоге
                if (!excelApp.Ready)
                {
                    return (false, "Excel занят. Завершите текущую операцию в Excel.");
                }

                Excel.Range activeCell = null;
                try
                {
                    activeCell = excelApp.ActiveCell;
                }
                catch (COMException comEx) when (comEx.ErrorCode == VBA_E_IGNORE)
                {
                    return (false, "Excel находится в режиме редактирования ячейки. Нажмите Enter или Esc на листе и повторите.");
                }

                if (activeCell == null)
                {
                    return (false, "На листе Excel не выбрана активная ячейка.");
                }

                Excel.Worksheet worksheet = activeCell.Worksheet;
                if (worksheet != null && worksheet.ProtectContents)
                {
                    return (false, "Лист Excel защищен от изменений. Снимите защиту листа.");
                }

                // Очистка значений от разрывов строк и табов для корректного разделения по столбцам
                string designation = item.Designation ?? string.Empty;
                string name = item.Name ?? string.Empty;

                designation = designation.Replace("\r", " ").Replace("\n", " ").Replace("\t", " ");
                name = name.Replace("\r", " ").Replace("\n", " ").Replace("\t", " ");

                // Формируем строку: Обозначение [TAB] Наименование
                string pasteText = $"{designation}\t{name}";

                // 1. Сохраняем снимок текущего буфера обмена пользователя
                DataObject previousClipboardSnapshot = SnapshotClipboard();

                try
                {
                    // Помещаем в буфер обмена для вставки в Excel
                    Clipboard.SetDataObject(pasteText, false, 5, 50);

                    // Гарантируем, что выделена ровно одна активная ячейка
                    activeCell.Select();

                    // Вызываем нативную команду вставки Excel Ribbon (ExecuteMso "Paste").
                    // Это создает полноценный шаг в стеке отмены Excel и позволяет отменять через Ctrl + Z!
                    dynamic commandBars = excelApp.CommandBars;
                    commandBars.ExecuteMso("Paste");
                }
                catch
                {
                    // Резервный способ (COM Interop) на случай, если команда ленты недоступна
                    try
                    {
                        activeCell.Value2 = designation;
                        Excel.Range nameCell = activeCell.Offset[0, 1];
                        nameCell.Value2 = name;
                    }
                    catch (Exception ex)
                    {
                        return (false, $"Ошибка записи в ячейку: {ex.Message}");
                    }
                }
                finally
                {
                    // 2. Немедленно восстанавливаем буфер обмена пользователя
                    if (previousClipboardSnapshot != null)
                    {
                        try
                        {
                            Clipboard.SetDataObject(previousClipboardSnapshot, true, 5, 50);
                        }
                        catch
                        {
                            // Игнорируем ошибку восстановления пользовательского буфера
                        }
                    }
                }

                // 3. Автопереход на строку ниже для удобства серийного ввода
                if (config != null && config.AutoAdvanceRow)
                {
                    try
                    {
                        Excel.Range nextRowCell = activeCell.Offset[1, 0];
                        nextRowCell.Select();
                    }
                    catch
                    {
                        // Игнорируем ошибку выделения, если достигнут край листа
                    }
                }

                return (true, $"Вставлено: {designation}");
            }
            catch (COMException comEx)
            {
                if (comEx.ErrorCode == VBA_E_IGNORE)
                {
                    return (false, "Excel находится в режиме редактирования ячейки. Нажмите Enter или Esc на листе и повторите.");
                }

                return (false, $"Ошибка Excel COM (0x{comEx.ErrorCode:X8}): {comEx.Message}");
            }
            catch (Exception ex)
            {
                return (false, $"Ошибка вставки: {ex.Message}");
            }
        }

        private static DataObject SnapshotClipboard()
        {
            try
            {
                IDataObject current = Clipboard.GetDataObject();
                if (current == null) return null;

                DataObject snapshot = new DataObject();
                string[] formats = current.GetFormats();
                if (formats == null || formats.Length == 0) return null;

                bool hasData = false;
                foreach (string format in formats)
                {
                    try
                    {
                        object data = current.GetData(format);
                        if (data != null)
                        {
                            snapshot.SetData(format, data);
                            hasData = true;
                        }
                    }
                    catch
                    {
                        // Некоторые OLE-форматы не могут быть сериализованы в память, пропускаем их
                    }
                }

                return hasData ? snapshot : null;
            }
            catch
            {
                return null;
            }
        }
    }
}

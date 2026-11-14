using System;
using System.Runtime.InteropServices;
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

                // 1. Вставка обозначения в активную ячейку
                activeCell.Value2 = item.Designation;

                // 2. Вставка наименования в соседнюю ячейку справа (Offset[0, 1])
                Excel.Range nameCell = activeCell.Offset[0, 1];
                nameCell.Value2 = item.Name;

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

                return (true, $"Вставлено: {item.Designation}");
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
    }
}

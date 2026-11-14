using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Threading;
using ExcelDataReader;
using SpecHelper.Models;

namespace SpecHelper.Services
{
    public static class ExcelCatalogReader
    {
        public static (CatalogNode RootNode, List<CatalogItem> AllItems) LoadCatalog(string filePath, CatalogType catalogType)
        {
            var root = new CatalogNode("Все разделы", string.Empty);
            var allItems = new List<CatalogItem>();

            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return (root, allItems);
            }

            // Retry policy при открытии файла
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
                        throw;
                    Thread.Sleep(delayMs);
                }
            }

            if (stream == null)
            {
                return (root, allItems);
            }

            using (stream)
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = false // Ручной парсинг заголовков для максимальной гибкости
                    }
                });

                if (result.Tables.Count == 0)
                {
                    return (root, allItems);
                }

                var table = result.Tables[0];
                if (table.Rows.Count < 2)
                {
                    return (root, allItems);
                }

                // Определение индексов колонок
                int colDesignation = -1;
                int colName = -1;
                int colFolder = -1;

                var headerRow = table.Rows[0];
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    var val = headerRow[c]?.ToString()?.Trim() ?? string.Empty;
                    if (string.Equals(val, "Обозначение", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(val, "Артикул", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(val, "Код", StringComparison.OrdinalIgnoreCase))
                    {
                        colDesignation = c;
                    }
                    else if (string.Equals(val, "Наименование", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(val, "Название", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(val, "Описание", StringComparison.OrdinalIgnoreCase))
                    {
                        colName = c;
                    }
                    else if (string.Equals(val, "Папка", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(val, "Наименование папки", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(val, "Категория", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(val, "Раздел", StringComparison.OrdinalIgnoreCase))
                    {
                        colFolder = c;
                    }
                }

                // Fallback, если заголовки не распознаны
                if (colDesignation == -1) colDesignation = 0;
                if (colName == -1 && table.Columns.Count > 1) colName = 1;
                if (colFolder == -1 && table.Columns.Count > 2) colFolder = 2;

                // Чтение строк данных
                for (int r = 1; r < table.Rows.Count; r++)
                {
                    var row = table.Rows[r];
                    var designation = colDesignation >= 0 && colDesignation < table.Columns.Count
                        ? row[colDesignation]?.ToString()?.Trim() ?? string.Empty
                        : string.Empty;

                    var name = colName >= 0 && colName < table.Columns.Count
                        ? row[colName]?.ToString()?.Trim() ?? string.Empty
                        : string.Empty;

                    var folder = colFolder >= 0 && colFolder < table.Columns.Count
                        ? row[colFolder]?.ToString()?.Trim() ?? string.Empty
                        : string.Empty;

                    if (string.IsNullOrEmpty(designation) && string.IsNullOrEmpty(name))
                    {
                        continue; // Пропуск пустых строк
                    }

                    var item = new CatalogItem
                    {
                        Designation = designation,
                        Name = name,
                        CategoryPath = folder,
                        CatalogType = catalogType
                    };

                    allItems.Add(item);

                    // Добавление в дерево узлов
                    AddCatalogItemToTree(root, item);
                }
            }

            return (root, allItems);
        }

        private static void AddCatalogItemToTree(CatalogNode root, CatalogItem item)
        {
            if (string.IsNullOrWhiteSpace(item.CategoryPath))
            {
                root.DirectItems.Add(item);
                return;
            }

            var parts = item.CategoryPath.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
            var currentNode = root;

            foreach (var part in parts)
            {
                currentNode = currentNode.FindOrCreateChild(part);
            }

            currentNode.DirectItems.Add(item);
        }
    }
}

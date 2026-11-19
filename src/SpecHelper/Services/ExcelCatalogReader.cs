using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
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

            FileStream stream = OpenFileWithRetry(filePath);
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
                        UseHeaderRow = false
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

                if (colDesignation == -1) colDesignation = 0;
                if (colName == -1 && table.Columns.Count > 1) colName = 1;
                if (colFolder == -1 && table.Columns.Count > 2) colFolder = 2;

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
                        continue;
                    }

                    var item = new CatalogItem
                    {
                        Designation = designation,
                        Name = name,
                        CategoryPath = folder,
                        CatalogType = catalogType
                    };

                    allItems.Add(item);
                    AddCatalogItemToTree(root, item);
                }
            }

            return (root, allItems);
        }

        public static TechProcessCatalogData LoadTechProcesses(string filePath)
        {
            var data = new TechProcessCatalogData();
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return data;
            }

            FileStream stream = OpenFileWithRetry(filePath);
            if (stream == null)
            {
                return data;
            }

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

                if (result.Tables.Count == 0)
                {
                    return data;
                }

                // Ищем лист "Техпроцессы" (или берем первый лист)
                DataTable tpTable = null;
                foreach (DataTable t in result.Tables)
                {
                    if (string.Equals(t.TableName?.Trim(), "Техпроцессы", StringComparison.OrdinalIgnoreCase))
                    {
                        tpTable = t;
                        break;
                    }
                }
                if (tpTable == null)
                {
                    tpTable = result.Tables[0];
                }

                // Парсинг листа техпроцессов
                if (tpTable.Rows.Count >= 2)
                {
                    int colDesig = -1, colName = -1, colComment = -1, colComp = -1;
                    var header = tpTable.Rows[0];
                    for (int c = 0; c < tpTable.Columns.Count; c++)
                    {
                        var val = header[c]?.ToString()?.Trim() ?? string.Empty;
                        if (string.Equals(val, "Обозначение", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(val, "Код", StringComparison.OrdinalIgnoreCase))
                            colDesig = c;
                        else if (string.Equals(val, "Наименование", StringComparison.OrdinalIgnoreCase) ||
                                 string.Equals(val, "Название", StringComparison.OrdinalIgnoreCase))
                            colName = c;
                        else if (string.Equals(val, "Комментарий", StringComparison.OrdinalIgnoreCase))
                            colComment = c;
                        else if (val.IndexOf("Состав", StringComparison.OrdinalIgnoreCase) >= 0)
                            colComp = c;
                    }

                    if (colDesig == -1) colDesig = 0;
                    if (colName == -1 && tpTable.Columns.Count > 1) colName = 1;
                    if (colComment == -1 && tpTable.Columns.Count > 2) colComment = 2;
                    if (colComp == -1)
                        colComp = tpTable.Columns.Count >= 5 ? 4 : tpTable.Columns.Count - 1;

                    for (int r = 1; r < tpTable.Rows.Count; r++)
                    {
                        var row = tpTable.Rows[r];
                        string desig = GetCellText(row, colDesig);
                        string name = GetCellText(row, colName);
                        string comment = GetCellText(row, colComment);
                        string comp = GetCellText(row, colComp);

                        if (string.IsNullOrEmpty(desig) && string.IsNullOrEmpty(name))
                            continue;

                        var codes = comp
                            .Split(new[] { '+' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(x => x.Trim())
                            .Where(x => !string.IsNullOrEmpty(x))
                            .ToList();

                        data.Processes.Add(new TechProcessItem
                        {
                            Designation = desig,
                            Name = name,
                            Comment = comment,
                            Composition = comp,
                            OperationCodes = codes
                        });
                    }
                }

                // Все остальные листы в книге ("Операции новые", "ЦМК" и будущие) — вкладки операций
                foreach (DataTable opTable in result.Tables)
                {
                    if (opTable == tpTable || opTable.Rows.Count < 2)
                        continue;

                    int colCode = -1, colName = -1, colComment = -1;
                    var header = opTable.Rows[0];
                    for (int c = 0; c < opTable.Columns.Count; c++)
                    {
                        var val = header[c]?.ToString()?.Trim() ?? string.Empty;
                        if (string.Equals(val, "Обозначение", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(val, "Код", StringComparison.OrdinalIgnoreCase))
                            colCode = c;
                        else if (string.Equals(val, "Наименование", StringComparison.OrdinalIgnoreCase) ||
                                 string.Equals(val, "Название", StringComparison.OrdinalIgnoreCase))
                            colName = c;
                        else if (string.Equals(val, "Комментарий", StringComparison.OrdinalIgnoreCase))
                            colComment = c;
                    }

                    if (colCode == -1) colCode = 0;
                    if (colName == -1 && opTable.Columns.Count > 1) colName = 1;
                    if (colComment == -1 && opTable.Columns.Count > 2) colComment = 2;

                    var ops = new List<TechOperationItem>();
                    for (int r = 1; r < opTable.Rows.Count; r++)
                    {
                        var row = opTable.Rows[r];
                        string code = GetCellText(row, colCode);
                        string name = GetCellText(row, colName);
                        string comment = GetCellText(row, colComment);

                        if (string.IsNullOrEmpty(code) && string.IsNullOrEmpty(name))
                            continue;

                        ops.Add(new TechOperationItem
                        {
                            Code = code,
                            Name = name,
                            Comment = comment,
                            SheetName = opTable.TableName
                        });
                    }

                    data.OperationSheets.Add(new KeyValuePair<string, List<TechOperationItem>>(opTable.TableName, ops));
                }
            }

            return data;
        }

        private static string GetCellText(DataRow row, int colIndex)
        {
            if (colIndex < 0 || colIndex >= row.Table.Columns.Count)
                return string.Empty;
            return row[colIndex]?.ToString()?.Trim() ?? string.Empty;
        }

        private static FileStream OpenFileWithRetry(string filePath)
        {
            int retries = 4;
            int delayMs = 250;

            for (int i = 0; i < retries; i++)
            {
                try
                {
                    return new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                }
                catch (IOException)
                {
                    if (i == retries - 1)
                        throw;
                    Thread.Sleep(delayMs);
                }
            }

            return null;
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

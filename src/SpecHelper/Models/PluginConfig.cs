using System;
using System.Collections.Generic;

namespace SpecHelper.Models
{
    public class PluginConfig
    {
        public string MaterialsPath { get; set; } = string.Empty;
        public string HardwarePath { get; set; } = string.Empty;
        public string ComponentsPath { get; set; } = string.Empty;
        public string TemplatesPath { get; set; } = string.Empty;
        public string TechProcessesPath { get; set; } = string.Empty;

        public bool AutoSyncEnabled { get; set; } = true;
        public int SyncDebounceMs { get; set; } = 800;
        public int SplitterDistance { get; set; } = 240;
        public bool IsTreeCollapsed { get; set; } = false;
        public bool AutoAdvanceRow { get; set; } = true;
        public bool SearchCurrentFolderOnly { get; set; } = false;
        public int WindowWidth { get; set; } = 450;
    }

    public class TechOperationItem
    {
        public string Code { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
    }

    public class TechProcessItem
    {
        public string Designation { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;
        public string Composition { get; set; } = string.Empty;
        public List<string> OperationCodes { get; set; } = new List<string>();

        public string DisplayComposition
        {
            get
            {
                if (string.IsNullOrWhiteSpace(Name))
                    return Composition;
                if (string.IsNullOrWhiteSpace(Composition))
                    return Name;
                return $"{Name} ({Composition})";
            }
        }
    }

    public class TechProcessCatalogData
    {
        public List<TechProcessItem> Processes { get; set; } = new List<TechProcessItem>();
        public List<KeyValuePair<string, List<TechOperationItem>>> OperationSheets { get; set; } =
            new List<KeyValuePair<string, List<TechOperationItem>>>();
    }
}

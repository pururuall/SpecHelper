using System;

namespace SpecHelper.Models
{
    public class PluginConfig
    {
        public string MaterialsPath { get; set; } = string.Empty;
        public string HardwarePath { get; set; } = string.Empty;
        public string ComponentsPath { get; set; } = string.Empty;

        public bool AutoSyncEnabled { get; set; } = true;
        public int SyncDebounceMs { get; set; } = 800;
        public int SplitterDistance { get; set; } = 240;
        public bool IsTreeCollapsed { get; set; } = false;
        public bool AutoAdvanceRow { get; set; } = true;
        public bool SearchCurrentFolderOnly { get; set; } = false;
        public int WindowWidth { get; set; } = 450;
    }
}

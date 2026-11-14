using System;
using System.Collections.Generic;

namespace SpecHelper.Models
{
    public enum CatalogType
    {
        Materials,
        Hardware,
        Components
    }

    public class CatalogItem
    {
        public string Designation { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string CategoryPath { get; set; } = string.Empty;
        public CatalogType CatalogType { get; set; }
        public Dictionary<string, string> ExtraColumns { get; set; } = new Dictionary<string, string>();

        public override string ToString()
        {
            return $"{Designation} | {Name}";
        }
    }
}

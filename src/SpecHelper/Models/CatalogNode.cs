using System;
using System.Collections.Generic;
using System.Linq;

namespace SpecHelper.Models
{
    public class CatalogNode
    {
        public string Name { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        public List<CatalogNode> Children { get; set; } = new List<CatalogNode>();
        public List<CatalogItem> DirectItems { get; set; } = new List<CatalogItem>();

        public CatalogNode(string name, string fullPath)
        {
            Name = name;
            FullPath = fullPath;
        }

        public IEnumerable<CatalogItem> GetAllItems()
        {
            foreach (var item in DirectItems)
            {
                yield return item;
            }

            foreach (var child in Children)
            {
                foreach (var item in child.GetAllItems())
                {
                    yield return item;
                }
            }
        }

        public int TotalItemCount => DirectItems.Count + Children.Sum(c => c.TotalItemCount);

        public CatalogNode FindOrCreateChild(string childName)
        {
            var trimmed = childName.Trim();
            var existing = Children.FirstOrDefault(c => string.Equals(c.Name, trimmed, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                return existing;
            }

            var childPath = string.IsNullOrEmpty(FullPath) ? trimmed : $"{FullPath} / {trimmed}";
            var newNode = new CatalogNode(trimmed, childPath);
            Children.Add(newNode);
            return newNode;
        }
    }
}

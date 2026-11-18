using System;

namespace SpecHelper.Models
{
    public class TemplateItem
    {
        public string Name { get; set; } = string.Empty;
        public int HeaderRow { get; set; }
        public int StartRow { get; set; }
        public int EndRow { get; set; }
        public int RowCount => (EndRow >= StartRow) ? (EndRow - StartRow + 1) : 0;

        public override string ToString()
        {
            return Name;
        }
    }
}

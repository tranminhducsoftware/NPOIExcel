using System.Collections.Generic;

namespace Application.Attribute
{
    public class SampleExportAttribute : System.Attribute
    {
        #region Properties

        public Dictionary<string, List<object>> Dependency { get; set; }

        public int RowIndex { get; }

        public int ColIndex { get; set; }

        public int ColSpan { get; }

        public int RowSpan { get; }

        public string Name { get; }

        public string ColName { get; set; }

        public int ColWidth { get; }

        public bool Required { get; }

        #endregion

        #region Constructor

        public SampleExportAttribute(int rowIndex = 0, int colIndex = 0, int colSpan = 1, int rowSpan = 1,
            string name = "", int colWidth = 50,
            bool required = false, Dictionary<string, List<object>> dependency = null)
        {
            RowIndex = rowIndex;
            ColIndex = colIndex;
            ColSpan = colSpan;
            RowSpan = rowSpan;
            Name = name;
            ColWidth = colWidth;
            Required = required;
            Dependency = dependency;
        }

        public SampleExportAttribute()
        {
            RowIndex = 0;
            ColSpan = 1;
            ColWidth = 50;
            RowSpan = 1;
            Required = false;
        }

        #endregion
    }
}
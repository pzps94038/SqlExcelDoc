using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SqlExcelDoc.Model
{
    public class TableSpecifications
    {
        public string TableName { get; set; } = string.Empty;
        public string ColumnName { get; set; } = string.Empty;
        public string DataType { get; set; } = string.Empty;
        public string NotNull { get; set; } = string.Empty;
        public string Length {  get; set; } = string.Empty;
        public string ConstraintType { get; set; } = string.Empty;
        public string IsUnique { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
    }
}

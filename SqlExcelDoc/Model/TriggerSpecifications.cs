using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SqlExcelDoc
{
    public class TriggerSpecifications
    {
        public string TableName { get; set; } = string.Empty;
        public string TriggerName { get; set; } = string.Empty;
        public string TypeDesc {  get; set; } = string.Empty;
    }
}

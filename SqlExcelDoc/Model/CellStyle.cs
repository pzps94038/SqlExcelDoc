using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SqlExcelDoc.Model
{
    public class CellStyle
    {
        public double FontHeightInPoints = 10;
        public string FontName = "微軟正黑體";
        public bool IsBold = false;
        public short FontColor = IndexedColors.Black.Index;
        public BorderStyle BorderStyle = BorderStyle.Medium;
        public short BorderColor = IndexedColors.Black.Index;
        public HorizontalAlignment Alignment = HorizontalAlignment.Left;
        public VerticalAlignment VerticalAlignment = VerticalAlignment.Center;
        public short FillForegroundColor = IndexedColors.White.Index;
        public FillPattern FillPattern = FillPattern.SolidForeground;
        public FontUnderlineType Underline = FontUnderlineType.None;
    }
}

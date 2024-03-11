using NPOI.SS.UserModel;
using SqlExcelDoc.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace SqlExcelDoc.Helper
{
    public static class SheetHelper
    {
        private static string _fontName = "微軟正黑體";
        public static void AutoSheetSize(this ISheet sheet, int maxColumn)
        {
            for (int columnNum = 0; columnNum <= maxColumn; columnNum++)
            {
                int columnWidth = 0; // 初始化列宽为0
                for (int rowNum = 0; rowNum <= sheet.LastRowNum; rowNum++)
                {
                    IRow currentRow = sheet.GetRow(rowNum) ?? sheet.CreateRow(rowNum); // 直接使用 currentRow

                    ICell currentCell = currentRow.GetCell(columnNum);
                    if (currentCell != null)
                    {
                        int len = Encoding.UTF8.GetBytes(currentCell.ToString()).Length;
                        if (columnWidth < len)
                        {
                            columnWidth = len; // 更新列宽为当前最大宽度
                        }
                    }
                }
                columnWidth = (columnWidth + 8) * 256;
                int maxColumnWidth = 255 * 256; // The maximum column width for an individual cell is 255 characters
                if (columnWidth > maxColumnWidth)
                {
                    columnWidth = maxColumnWidth;
                }
                sheet.SetColumnWidth(columnNum, columnWidth);
            }
        }

        public static ICell CreateTitleStyleCell(this IRow row, int column, HorizontalAlignment alignment = HorizontalAlignment.Left, VerticalAlignment verticalAlignment = VerticalAlignment.Center)
        {
            var workbook = row.Sheet.Workbook;
            var cell = row.CreateCell(column);
            // 創建一個字體樣式
            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = 12;
            font.FontName = _fontName;
            font.IsBold = true;
            font.Color = IndexedColors.Black.Index; // 字體顏色
            // 創建一個單元格樣式並應用這個字體
            ICellStyle style = workbook.CreateCellStyle();
            style.SetFont(font);
            var borderStyle = BorderStyle.Medium;
            var borderColor = IndexedColors.Black.Index;
            style.BorderBottom = borderStyle; // 下邊框
            style.BottomBorderColor = borderColor; // 下邊框顏色
            style.BorderLeft = borderStyle; // 左邊框
            style.LeftBorderColor = borderColor; // 左邊框顏色
            style.BorderRight = borderStyle; // 右邊框
            style.RightBorderColor = borderColor; // 右邊框顏色
            style.BorderTop = borderStyle; // 上邊框
            style.TopBorderColor = borderColor; // 上邊框顏色
            style.Alignment = alignment; // 水平居中
            style.VerticalAlignment = verticalAlignment; // 垂直居中
            // 填充背景顏色
            style.FillForegroundColor = IndexedColors.White.Index;
            style.FillPattern = FillPattern.SolidForeground;
            cell.CellStyle = style;
            return cell;
        }

        public static ICell CreateStyleCell(this IRow row, int column)
        {
            var cellStyle = new CellStyle();
            var workbook = row.Sheet.Workbook;
            var cell = row.CreateCell(column);
            // 創建一個字體樣式
            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = cellStyle.FontHeightInPoints;
            font.FontName = cellStyle.FontName;
            font.IsBold = cellStyle.IsBold;
            font.Color = cellStyle.FontColor; // 字體顏色
            // 創建一個單元格樣式並應用這個字體
            ICellStyle style = workbook.CreateCellStyle();
            style.SetFont(font);
            var borderStyle = cellStyle.BorderStyle;
            var borderColor = cellStyle.BorderColor;
            style.BorderBottom = borderStyle; // 下邊框
            style.BottomBorderColor = borderColor; // 下邊框顏色
            style.BorderLeft = borderStyle; // 左邊框
            style.LeftBorderColor = borderColor; // 左邊框顏色
            style.BorderRight = borderStyle; // 右邊框
            style.RightBorderColor = borderColor; // 右邊框顏色
            style.BorderTop = borderStyle; // 上邊框
            style.TopBorderColor = borderColor; // 上邊框顏色
            style.Alignment = cellStyle.Alignment; // 水平居中
            style.VerticalAlignment = cellStyle.VerticalAlignment; // 垂直居中
            // 填充背景顏色
            style.FillForegroundColor = cellStyle.FillForegroundColor;
            style.FillPattern = cellStyle.FillPattern;
            cell.CellStyle = style;
            return cell;
        }

        public static ICell CreateStyleCell(this IRow row, int column, CellStyle cellStyle)
        {
            var workbook = row.Sheet.Workbook;
            var cell = row.CreateCell(column);
            // 創建一個字體樣式
            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = cellStyle.FontHeightInPoints;
            font.FontName = cellStyle.FontName;
            font.IsBold = cellStyle.IsBold;
            font.Color = cellStyle.FontColor; // 字體顏色
            font.Underline = cellStyle.Underline;
            // 創建一個單元格樣式並應用這個字體
            ICellStyle style = workbook.CreateCellStyle();
            style.SetFont(font);
            var borderStyle = cellStyle.BorderStyle;
            var borderColor = cellStyle.BorderColor;
            style.BorderBottom = borderStyle; // 下邊框
            style.BottomBorderColor = borderColor; // 下邊框顏色
            style.BorderLeft = borderStyle; // 左邊框
            style.LeftBorderColor = borderColor; // 左邊框顏色
            style.BorderRight = borderStyle; // 右邊框
            style.RightBorderColor = borderColor; // 右邊框顏色
            style.BorderTop = borderStyle; // 上邊框
            style.TopBorderColor = borderColor; // 上邊框顏色
            style.Alignment = cellStyle.Alignment; // 水平居中
            style.VerticalAlignment = cellStyle.VerticalAlignment; // 垂直居中
            // 填充背景顏色
            style.FillForegroundColor = cellStyle.FillForegroundColor;
            style.FillPattern = cellStyle.FillPattern;
            cell.CellStyle = style;
            return cell;
        }
    }
}

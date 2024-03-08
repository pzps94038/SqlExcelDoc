using NPOI.SS.UserModel;
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
                sheet.SetColumnWidth(columnNum, (columnWidth + 5) *  256);
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

        public static ICell CreateHeaderStyleCell(this IRow row, int column, HorizontalAlignment alignment = HorizontalAlignment.Left, VerticalAlignment verticalAlignment = VerticalAlignment.Center)
        {
            var workbook = row.Sheet.Workbook;
            var cell = row.CreateCell(column);
            // 創建一個字體樣式
            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = 12;
            font.FontName = _fontName;
            font.IsBold = true;
            font.Color = IndexedColors.White.Index; // 字體顏色
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
            style.FillForegroundColor = IndexedColors.RoyalBlue.Index;
            style.FillPattern = FillPattern.SolidForeground;
            cell.CellStyle = style;
            return cell;
        }

        public static ICell CreateContentStyleCell(this IRow row, int column, HorizontalAlignment alignment = HorizontalAlignment.Left, VerticalAlignment verticalAlignment = VerticalAlignment.Center)
        {
            var workbook = row.Sheet.Workbook;
            var cell = row.CreateCell(column);
            // 創建一個字體樣式
            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = 10;
            font.FontName = _fontName;
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

        public static ICell CreatePrimaryKeyStyleCell(this IRow row, int column, HorizontalAlignment alignment = HorizontalAlignment.Left, VerticalAlignment verticalAlignment = VerticalAlignment.Center)
        {
            var workbook = row.Sheet.Workbook;
            var cell = row.CreateCell(column);
            // 創建一個字體樣式
            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = 10;
            font.FontName = _fontName;
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
            style.FillForegroundColor = IndexedColors.Yellow.Index;
            style.FillPattern = FillPattern.SolidForeground;
            cell.CellStyle = style;
            return cell;
        }

        public static ICell CreateForeignKeyStyleCell(this IRow row, int column, HorizontalAlignment alignment = HorizontalAlignment.Left, VerticalAlignment verticalAlignment = VerticalAlignment.Center)
        {
            var workbook = row.Sheet.Workbook;
            var cell = row.CreateCell(column);
            // 創建一個字體樣式
            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = 10;
            font.FontName = _fontName;
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
            style.FillForegroundColor = IndexedColors.LightOrange.Index;
            style.FillPattern = FillPattern.SolidForeground;
            cell.CellStyle = style;
            return cell;
        }
    }
}

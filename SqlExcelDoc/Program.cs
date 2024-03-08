using NPOI.POIFS.Crypt;
using NPOI.POIFS.Crypt.Agile;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using SqlExcelDoc.Helper;
using SqlExcelDoc.Model;
using System.Diagnostics;

namespace SqlExcelDoc
{
    public class Program
    {
        static void Main(string[] args)
        {
            // Add console trace listener
            Trace.Listeners.Add(new ConsoleTraceListener());

            // Run actions
            Consolery.Run();
        }

        [Action("產生資料庫規格文件")]
        public static void CreateDocumentation(
            [Required(Description = "連線字串")] string connection,
            [Required(Description = "輸出路徑")] string fileName,
            [Optional("MsSql", Description = "資料庫格式")] string type,
            [Optional("y", Description = "如果已存在是否覆蓋")] string overwriteString
            )
        {
            try
            {
                SqlDoc sqlDoc;
                var upperType = type.ToUpper();
                var overwrite = overwriteString.ToUpper() == "Y";
                switch (upperType)
                {
                    case "MSSQL":
                        sqlDoc = new MsSqlSqlDoc(connection);
                        break;
                    default:
                        PrintMessage("不支援的資料庫類型");
                        return;
                }
                var isExists = File.Exists(fileName);
                if (overwrite)
                {
                    if (isExists)
                    {
                        File.Delete(fileName);
                    }
                }
                else
                {
                    PrintMessage("發生錯誤: 文檔已存在. 使用 /y 來做覆蓋");
                    return;
                }
                IWorkbook workbook = new XSSFWorkbook();
                PrintMessage("產生資料庫規格中...");
                var databaseSpecifications = sqlDoc.GetDatabaseSpecifications();
                var databaseViewSpecifications = sqlDoc.GetDatabaseViewSpecifications();
                GenerateDatabaseSpecifications(workbook, databaseSpecifications, databaseViewSpecifications);
                PrintMessage("產生資料庫規格完成...");
                PrintMessage("產生預存程序規格中...");
                var storedProcedureSpecifications = sqlDoc.GetStoredProcedureSpecifications();
                if (storedProcedureSpecifications.Any())
                {
                    GenerateProcedureSpecifications(workbook, storedProcedureSpecifications);
                    PrintMessage("產生預存程序規格完成...");
                }
                else
                {
                    PrintMessage("無任何預存程序...");
                }
                PrintMessage("產生表格規格中...");
                var tableSpecifications = sqlDoc.GetTableSpecifications();
                GenerateDatabaseSpecifications(workbook, tableSpecifications);
                PrintMessage("產生表格規格完成...");
               
                FileStream sw = File.Create(fileName);
                workbook.Write(sw);
                sw.Close();
                sqlDoc.Dispose();
                PrintMessage("文件已產生完成...");
            }
            catch (Exception ex) 
            {
                PrintMessage("發生錯誤" + ex.Message);
            }
        }

        /// <summary>
        /// 輸出訊息包含時間
        /// </summary>
        /// <param name="msg"></param>
        private static void PrintMessage(string msg) 
        {
            Console.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " " + msg);
        }

        /// <summary>
        /// 產生資料庫表格清單
        /// </summary>
        private static void GenerateDatabaseSpecifications(IWorkbook workbook, IEnumerable<DatabaseSpecifications> databaseSpecifications, IEnumerable<DatabaseSpecifications> databaseViewSpecifications) 
        {
            var sheet = workbook.CreateSheet("表格清單目錄");
            // 表頭
            var headerRow = sheet.CreateRow(0);
            headerRow.CreateHeaderStyleCell(0).SetCellValue("項次");
            headerRow.CreateHeaderStyleCell(1).SetCellValue("表格名稱");
            headerRow.CreateHeaderStyleCell(2).SetCellValue("描述");
            
            int i = 1;
            foreach (var item in databaseSpecifications)
            {
                var row = sheet.CreateRow(i);
                i++;
                row.CreateContentStyleCell(0).SetCellValue(i - 1);
                var tableNameCell = row.CreateContentStyleCell(1);
                var hyperlink = new XSSFHyperlink(HyperlinkType.Document)
                {
                    Address = $"'{item.TableName}'!A1"
                };
                tableNameCell.SetCellValue((item.TableName as string) ?? "");
                tableNameCell.Hyperlink = hyperlink;
                var style = tableNameCell.CellStyle;
                var font = style.GetFont(workbook);
                font.Color = IndexedColors.Blue.Index;
                font.Underline = FontUnderlineType.Single;
                style.SetFont(font);
                tableNameCell.CellStyle = style;
                row.CreateContentStyleCell(2).SetCellValue((item.Description as string) ?? "");
            }

            foreach (var item in databaseViewSpecifications)
            {
                var row = sheet.CreateRow(i);
                i++;
                row.CreateContentStyleCell(0).SetCellValue(i - 1);
                row.CreateContentStyleCell(1).SetCellValue((item.TableName as string) ?? "");
                row.CreateContentStyleCell(2).SetCellValue((item.Description as string) ?? "");
            }
            CellRangeAddress filterRange = new CellRangeAddress(0, i, 0, 2);

            // 在工作表上設置自動篩選的範圍
            sheet.SetAutoFilter(filterRange);
            sheet.AutoSheetSize(3);
        }

        /// <summary>
        /// 產生資料庫表格細項
        /// </summary>
        private static void GenerateDatabaseSpecifications(IWorkbook workbook, IEnumerable<TableSpecifications> tableSpecifications)
        {
            var group = tableSpecifications.GroupBy(a => a.TableName);
            foreach (var keyPair in group) 
            {
                var sheet = workbook.CreateSheet(keyPair.Key);
                var headerRow = sheet.CreateRow(0);
                headerRow.CreateHeaderStyleCell(0).SetCellValue("項次");
                headerRow.CreateHeaderStyleCell(1).SetCellValue("欄位名稱");
                headerRow.CreateHeaderStyleCell(2).SetCellValue("型態");
                headerRow.CreateHeaderStyleCell(3).SetCellValue("長度");
                headerRow.CreateHeaderStyleCell(4).SetCellValue("NOT NULL");
                headerRow.CreateHeaderStyleCell(5).SetCellValue("UNIQUE");
                headerRow.CreateHeaderStyleCell(6).SetCellValue("描述");
                int i = 1;
                foreach (var item in keyPair.ToList()) 
                {
                    var row = sheet.CreateRow(i);
                    i++;
                    var isPrimaryKey = item.ConstraintType == "PRIMARY KEY";
                    var isForeignKey = item.ConstraintType == "FOREIGN KEY";
                    if (isPrimaryKey)
                    {
                        row.CreatePrimaryKeyStyleCell(0).SetCellValue(i - 1);
                        row.CreatePrimaryKeyStyleCell(1).SetCellValue((item.ColumnName as string) ?? "");
                        row.CreatePrimaryKeyStyleCell(2).SetCellValue((item.DataType as string).ToUpper() ?? "");
                        var len = ((item.Length as string) ?? "") == "-1" ? "MAX" : ((item.Length as string) ?? "");
                        row.CreatePrimaryKeyStyleCell(3).SetCellValue(len);
                        row.CreatePrimaryKeyStyleCell(4).SetCellValue((item.NotNull as string) ?? "");
                        row.CreatePrimaryKeyStyleCell(5).SetCellValue((item.IsUnique as string) ?? "");
                        row.CreatePrimaryKeyStyleCell(6).SetCellValue((item.Description as string) ?? "");
                    }
                    else if (isForeignKey) 
                    {
                        row.CreateForeignKeyStyleCell(0).SetCellValue(i - 1);
                        row.CreateForeignKeyStyleCell(1).SetCellValue((item.ColumnName as string) ?? "");
                        row.CreateForeignKeyStyleCell(2).SetCellValue((item.DataType as string).ToUpper() ?? "");
                        var len = ((item.Length as string) ?? "") == "-1" ? "MAX" : ((item.Length as string) ?? "");
                        row.CreateForeignKeyStyleCell(3).SetCellValue(len);
                        row.CreateForeignKeyStyleCell(4).SetCellValue((item.NotNull as string) ?? "");
                        row.CreateForeignKeyStyleCell(5).SetCellValue((item.IsUnique as string) ?? "");
                        row.CreateForeignKeyStyleCell(6).SetCellValue((item.Description as string) ?? "");
                    }
                    else
                    {
                        row.CreateContentStyleCell(0).SetCellValue(i - 1);
                        row.CreateContentStyleCell(1).SetCellValue((item.ColumnName as string) ?? "");
                        row.CreateContentStyleCell(2).SetCellValue((item.DataType as string).ToUpper() ?? "");
                        var len = ((item.Length as string) ?? "") == "-1" ? "MAX" : ((item.Length as string) ?? "");
                        row.CreateContentStyleCell(3).SetCellValue(len);
                        row.CreateContentStyleCell(4).SetCellValue((item.NotNull as string) ?? "");
                        row.CreateContentStyleCell(5).SetCellValue((item.IsUnique as string) ?? "");
                        row.CreateContentStyleCell(6).SetCellValue((item.Description as string) ?? "");
                    }
                    
                }
                CellRangeAddress filterRange = new CellRangeAddress(0, i, 0, 6);

                // 在工作表上設置自動篩選的範圍
                sheet.SetAutoFilter(filterRange);
                sheet.AutoSheetSize(6);
            }
        }

        /// <summary>
        /// 產生預存程序細項
        /// </summary>
        private static void GenerateProcedureSpecifications(IWorkbook workbook, IEnumerable<ProcedureSpecifications> storedProcedureSpecifications)
        {
            var sheet = workbook.CreateSheet("預存程序清單目錄");
            var headerRow = sheet.CreateRow(0);
            headerRow.CreateHeaderStyleCell(0).SetCellValue("項次");
            headerRow.CreateHeaderStyleCell(1).SetCellValue("預存程序名稱");
            headerRow.CreateHeaderStyleCell(2).SetCellValue("描述");
            int i = 1;
            foreach (var item in storedProcedureSpecifications)
            {
                var row = sheet.CreateRow(i);
                i++;
                row.CreateContentStyleCell(0).SetCellValue(i - 1);
                row.CreateContentStyleCell(1).SetCellValue((item.ProcedureName as string) ?? "");
                row.CreateContentStyleCell(2).SetCellValue((item.Description as string) ?? "");
            }
            CellRangeAddress filterRange = new CellRangeAddress(0, i, 0, 2);
            // 在工作表上設置自動篩選的範圍
            sheet.SetAutoFilter(filterRange);
            sheet.AutoSheetSize(6);
            
        }

    }
}

using Dapper;
using Microsoft.Data.SqlClient;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using Org.BouncyCastle.Utilities.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace SqlExcelDoc.Model
{
    internal class MsSqlSqlDoc : SqlDoc
    {
        private SqlConnection _connection;
        public MsSqlSqlDoc(string connectionString) : base(connectionString)
        {
            _connection = new SqlConnection(connectionString);
            _connection.Open();
        }

        public override IEnumerable<DatabaseSpecifications> GetDatabaseSpecifications()
        {
            // 一般表格
            var sql = @"SELECT
	                    S.name + '.' + O.name as TableName, 
	                    O.type_desc		AS [Type], 
	                    P.value			AS [Description]
                    FROM 
	                    sys.objects AS O
	                    LEFT JOIN sys.schemas AS S on S.schema_id = O.schema_id
	                    LEFT JOIN sys.extended_properties AS P ON P.major_id = O.object_id AND P.minor_id = 0 and P.name = 'MS_Description' 
                    WHERE 
	                    is_ms_shipped = 0 
	                    AND parent_object_id = 0
                        AND type_desc = 'USER_TABLE'";
            var result = _connection.Query<DatabaseSpecifications>(sql);
            return result;
        }

        public override IEnumerable<DatabaseSpecifications> GetDatabaseViewSpecifications()
        {
            // 檢視表
            var sql = @"SELECT
	                    S.name + '.' + O.name AS TableName, 
	                    O.type_desc		AS [Type], 
	                    P.value			AS [Description]
                    FROM 
	                    sys.objects AS O
	                    LEFT JOIN sys.schemas AS S on S.schema_id = O.schema_id
	                    LEFT JOIN sys.extended_properties AS P ON P.major_id = O.object_id AND P.minor_id = 0 and P.name = 'MS_Description' 
                    WHERE 
	                    is_ms_shipped = 0 
	                    AND parent_object_id = 0
                        AND type_desc = 'VIEW'";
            var result = _connection.Query<DatabaseSpecifications>(sql);
            return result;
        }

        public override void GenerateDatabaseStoredProcedureSpecifications(IWorkbook workbook)
        {
            throw new NotImplementedException();
        }

        public override IEnumerable<TableSpecifications> GetTableSpecifications()
        {
            //// 一般表格
            var sql = @"SELECT
                t.TABLE_SCHEMA + '.' + t.TABLE_NAME AS TableName,
                c.COLUMN_NAME as ColumnName,
                c.DATA_TYPE as DataType,
                CASE 
				    WHEN c.IS_NULLABLE = 'YES' THEN ''
				    WHEN c.IS_NULLABLE = 'NO' THEN 'Y'
				    ELSE ''
				END as NotNull,
                c.CHARACTER_MAXIMUM_LENGTH as Length,
                CASE 
				    WHEN k.CONSTRAINT_TYPE = 'PRIMARY KEY' THEN 'PRIMARY KEY'
				    WHEN k.CONSTRAINT_TYPE = 'FOREIGN KEY' THEN 'FOREIGN KEY'
				    ELSE ''
				END as ConstraintType,
                CASE WHEN k.CONSTRAINT_TYPE = 'UNIQUE' THEN 'Y' ELSE 'N' END as IsUnique,
                ISNULL(ep.value, '') AS Description
                FROM 
                    INFORMATION_SCHEMA.TABLES t
                INNER JOIN 
                    INFORMATION_SCHEMA.COLUMNS c ON t.TABLE_SCHEMA = c.TABLE_SCHEMA AND t.TABLE_NAME = c.TABLE_NAME
                LEFT JOIN 
                    INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE ccu ON c.TABLE_SCHEMA = ccu.TABLE_SCHEMA AND c.TABLE_NAME = ccu.TABLE_NAME AND c.COLUMN_NAME = ccu.COLUMN_NAME
                LEFT JOIN 
                    INFORMATION_SCHEMA.TABLE_CONSTRAINTS k ON ccu.CONSTRAINT_SCHEMA = k.CONSTRAINT_SCHEMA AND ccu.CONSTRAINT_NAME = k.CONSTRAINT_NAME
               	LEFT JOIN 
    				sys.extended_properties ep ON ep.major_id = OBJECT_ID(t.TABLE_SCHEMA + '.' + t.TABLE_NAME) AND ep.minor_id = c.ORDINAL_POSITION AND ep.name = 'MS_Description'
                WHERE 
                    t.TABLE_TYPE = 'BASE TABLE'
                ORDER BY 
                    t.TABLE_NAME, k.CONSTRAINT_TYPE DESC
            ";
            var result = _connection.Query<TableSpecifications>(sql);
            return result;
        }

        public override void GenerateDatabaseTriggerSpecifications(IWorkbook workbook)
        {
            throw new NotImplementedException();
        }

        public override void Dispose()
        {
            _connection.Dispose();
        }
    }
}

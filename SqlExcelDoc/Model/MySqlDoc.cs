using Dapper;
using Microsoft.Data.SqlClient;
using MySql.Data.MySqlClient;
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
    internal class MySqlDoc : SqlDoc
    {
        private MySqlConnection _connection;
        public MySqlDoc(string connectionString) : base(connectionString)
        {
            _connection = new MySqlConnection(connectionString);
            _connection.Open();
        }

        public override IEnumerable<DatabaseSpecifications> GetDatabaseSpecifications()
        {
            var sql = @"SELECT 
                            TABLE_NAME AS TableName,
                            TABLE_TYPE AS Type,
                            TABLE_COMMENT AS Description
                        FROM 
                            information_schema.TABLES
                        WHERE 
                            TABLE_SCHEMA NOT IN ('information_schema', 'mysql', 'performance_schema', 'sys')
                        ORDER BY 
                            TableName;";
            var result = _connection.Query<DatabaseSpecifications>(sql);
            return result;
        }

        public override IEnumerable<DatabaseSpecifications> GetDatabaseViewSpecifications()
        {
            var sql = @"SELECT
                            TABLE_NAME AS TableName, 
                            'VIEW' AS Type, 
                            TABLE_COMMENT AS Description
                        FROM 
                            information_schema.TABLES
                        WHERE 
                            TABLE_TYPE = 'VIEW'
                            AND TABLE_SCHEMA NOT IN ('information_schema', 'mysql', 'performance_schema', 'sys')
                        ORDER BY 
                            TableName;";
            var result = _connection.Query<DatabaseSpecifications>(sql);
            return result;
        }

        public override IEnumerable<ProcedureSpecifications> GetStoredProcedureSpecifications()
        {
            var sql = @"SELECT 
                            routine_name AS ProcedureName,
                            routine_comment AS Description
                        FROM 
                            information_schema.ROUTINES
                        WHERE 
                            routine_type = 'PROCEDURE'
                            AND routine_schema NOT IN ('information_schema', 'mysql', 'performance_schema', 'sys')
                        ORDER BY 
                            ProcedureName;";
            var result = _connection.Query<ProcedureSpecifications>(sql);
            return result;
        }

        public override IEnumerable<TableSpecifications> GetTableSpecifications()
        {
            var sql = @"SELECT
                            t.TABLE_NAME AS TableName,
                            c.COLUMN_NAME AS ColumnName,
                            c.DATA_TYPE AS DataType,
                            CASE 
                                WHEN c.IS_NULLABLE = 'YES' THEN 'N'
                                ELSE 'Y'
                            END AS NotNull,
                            c.CHARACTER_MAXIMUM_LENGTH AS Length,
                            IFNULL(c.COLUMN_COMMENT, '') AS Description,
                            MAX(CASE WHEN k.CONSTRAINT_TYPE = 'UNIQUE' THEN 'Y' ELSE 'N' END) AS IsUnique,
                            MAX(CASE WHEN k.CONSTRAINT_TYPE = 'PRIMARY KEY' THEN 'Y' ELSE 'N' END) AS IsPrimaryKey,
                            MAX(CASE WHEN k.CONSTRAINT_TYPE = 'FOREIGN KEY' THEN 'Y' ELSE 'N' END) AS IsForeignKey,
                            fk.REFERENCED_TABLE_NAME AS ReferencedTableName,
                            fk.REFERENCED_COLUMN_NAME AS ReferencedColumnName
                        FROM 
                            INFORMATION_SCHEMA.TABLES t
                        INNER JOIN 
                            INFORMATION_SCHEMA.COLUMNS c ON t.TABLE_SCHEMA = c.TABLE_SCHEMA AND t.TABLE_NAME = c.TABLE_NAME
                        LEFT JOIN 
                            INFORMATION_SCHEMA.KEY_COLUMN_USAGE kcu ON kcu.TABLE_SCHEMA = c.TABLE_SCHEMA AND kcu.TABLE_NAME = c.TABLE_NAME AND kcu.COLUMN_NAME = c.COLUMN_NAME
                        LEFT JOIN 
                            INFORMATION_SCHEMA.TABLE_CONSTRAINTS k ON kcu.CONSTRAINT_SCHEMA = k.CONSTRAINT_SCHEMA AND kcu.CONSTRAINT_NAME = k.CONSTRAINT_NAME
                        LEFT JOIN 
                            (SELECT 
                                kcu.TABLE_SCHEMA,
                                kcu.TABLE_NAME,
                                kcu.COLUMN_NAME,
                                kcu.REFERENCED_TABLE_NAME,
                                kcu.REFERENCED_COLUMN_NAME
                             FROM 
                                INFORMATION_SCHEMA.KEY_COLUMN_USAGE kcu
                             WHERE 
                                kcu.REFERENCED_TABLE_NAME IS NOT NULL
                            ) AS fk ON kcu.TABLE_SCHEMA = fk.TABLE_SCHEMA AND kcu.TABLE_NAME = fk.TABLE_NAME AND kcu.COLUMN_NAME = fk.COLUMN_NAME
                        WHERE 
                            t.TABLE_TYPE = 'BASE TABLE'
                            AND t.TABLE_SCHEMA NOT IN ('information_schema', 'mysql', 'performance_schema', 'sys')
                        GROUP BY
                            t.TABLE_SCHEMA, t.TABLE_NAME, c.COLUMN_NAME, c.DATA_TYPE, c.IS_NULLABLE, c.CHARACTER_MAXIMUM_LENGTH, c.COLUMN_COMMENT, fk.REFERENCED_TABLE_NAME, fk.REFERENCED_COLUMN_NAME
                        ORDER BY 
                            TableName, IsPrimaryKey DESC, IsForeignKey DESC, ColumnName;
            ";
            var result = _connection.Query<TableSpecifications>(sql);
            return result;
        }

        public override IEnumerable<TriggerSpecifications> GetTriggerSpecifications()
        {
            var sql = @"SELECT 
                            event_object_table AS TableName,
                            trigger_name AS TriggerName,
                            action_timing AS TypeDesc
                        FROM 
                            information_schema.TRIGGERS
                        WHERE 
                            trigger_schema NOT IN ('information_schema', 'mysql', 'performance_schema', 'sys')
                        ORDER BY 
                            TableName, 
                            TriggerName;";
            var result = _connection.Query<TriggerSpecifications>(sql);
            return result;
        }
        public override void Dispose()
        {
            _connection.Dispose();
        }
    }
}

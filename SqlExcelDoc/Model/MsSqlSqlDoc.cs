﻿using Dapper;
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
                        AND type_desc = 'USER_TABLE'
                        ORDER BY TableName";
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
                        AND type_desc = 'VIEW'
                        ORDER BY TableName";
            var result = _connection.Query<DatabaseSpecifications>(sql);
            return result;
        }

        public override IEnumerable<ProcedureSpecifications> GetStoredProcedureSpecifications()
        {
            // 檢視表
            var sql = @"SELECT
                    SCHEMA_NAME(p.schema_id) + '.' + p.name as ProcedureName,
                    ep.value AS Description
                FROM
                    sys.procedures p
                LEFT JOIN
                    sys.extended_properties ep ON p.object_id = ep.major_id
                    AND ep.minor_id = 0
                    AND ep.name = 'MS_Description'
                WHERE
                    p.is_ms_shipped = 0
                ORDER BY
                    ProcedureName;";
            var result = _connection.Query<ProcedureSpecifications>(sql);
            return result;
        }

        public override IEnumerable<TableSpecifications> GetTableSpecifications()
        {
            var sql = @"
                    SELECT
                    t.TABLE_SCHEMA + '.' + t.TABLE_NAME AS TableName,
                    c.COLUMN_NAME AS ColumnName,
                    c.DATA_TYPE AS DataType,
                    CASE 
                        WHEN c.IS_NULLABLE = 'YES' THEN 'N'
                        WHEN c.IS_NULLABLE = 'NO' THEN 'Y'
                        ELSE 'N'
                    END AS NotNull,
                    c.CHARACTER_MAXIMUM_LENGTH AS Length,
                    ISNULL(ep.value, '') AS Description,
                    MAX(CASE WHEN k.CONSTRAINT_TYPE = 'UNIQUE' THEN 'Y' ELSE 'N' END) AS IsUnique,
                    MAX(CASE WHEN k.CONSTRAINT_TYPE = 'PRIMARY KEY' THEN 'Y' ELSE 'N' END) AS IsPrimaryKey,
                    MAX(CASE WHEN k.CONSTRAINT_TYPE = 'FOREIGN KEY' THEN 'Y' ELSE 'N' END) AS IsForeignKey,
                    (fk.TABLE_SCHEMA + '.' + fk.REFERENCED_TABLE_NAME) AS ReferencedTableName,
                    fk.REFERENCED_COLUMN_NAME AS ReferencedColumnName
                    FROM 
                        INFORMATION_SCHEMA.TABLES t
                    INNER JOIN 
                        INFORMATION_SCHEMA.COLUMNS c ON t.TABLE_SCHEMA = c.TABLE_SCHEMA AND t.TABLE_NAME = c.TABLE_NAME
                    LEFT JOIN 
                        sys.extended_properties ep ON ep.major_id = OBJECT_ID(t.TABLE_SCHEMA + '.' + t.TABLE_NAME) AND ep.minor_id = c.ORDINAL_POSITION AND ep.name = 'MS_Description'
                    LEFT JOIN 
                        INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE ccu ON c.TABLE_SCHEMA = ccu.TABLE_SCHEMA AND c.TABLE_NAME = ccu.TABLE_NAME AND c.COLUMN_NAME = ccu.COLUMN_NAME
                    LEFT JOIN 
                        INFORMATION_SCHEMA.TABLE_CONSTRAINTS k ON ccu.CONSTRAINT_SCHEMA = k.CONSTRAINT_SCHEMA AND ccu.CONSTRAINT_NAME = k.CONSTRAINT_NAME
                    LEFT JOIN 
                        (SELECT 
                            rc.CONSTRAINT_SCHEMA,
                            rc.CONSTRAINT_NAME,
                            kcu.TABLE_SCHEMA,
                            kcu.TABLE_NAME,
                            kcu.COLUMN_NAME,
                            kcu2.TABLE_NAME AS REFERENCED_TABLE_NAME,
                            kcu2.COLUMN_NAME AS REFERENCED_COLUMN_NAME
                         FROM 
                            INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS rc
                         INNER JOIN 
                            INFORMATION_SCHEMA.KEY_COLUMN_USAGE kcu ON kcu.CONSTRAINT_NAME = rc.CONSTRAINT_NAME
                         INNER JOIN 
                            INFORMATION_SCHEMA.KEY_COLUMN_USAGE kcu2 ON kcu2.CONSTRAINT_NAME = rc.UNIQUE_CONSTRAINT_NAME AND kcu2.CONSTRAINT_SCHEMA = rc.UNIQUE_CONSTRAINT_SCHEMA
                        ) AS fk ON fk.TABLE_SCHEMA = c.TABLE_SCHEMA AND fk.TABLE_NAME = c.TABLE_NAME AND fk.COLUMN_NAME = c.COLUMN_NAME
                    WHERE 
                        t.TABLE_TYPE = 'BASE TABLE'
                    GROUP BY
                        t.TABLE_SCHEMA, t.TABLE_NAME, c.COLUMN_NAME, c.DATA_TYPE, c.IS_NULLABLE, c.CHARACTER_MAXIMUM_LENGTH, ep.value, (fk.TABLE_SCHEMA + '.' + fk.REFERENCED_TABLE_NAME), fk.REFERENCED_COLUMN_NAME
                    ORDER BY 
                        TableName, IsPrimaryKey DESC, IsForeignKey DESC, ColumnName
            ";
            var result = _connection.Query<TableSpecifications>(sql);
            return result;
        }

        public override IEnumerable<TriggerSpecifications> GetTriggerSpecifications()
        {
            var sql = @"SELECT 
                s.name + '.' + t.name AS TableName,
                tr.name AS TriggerName,
                tr.type_desc AS TypeDesc
                FROM 
                    sys.triggers tr
                INNER JOIN 
                    sys.tables t ON tr.parent_id = t.object_id
                INNER JOIN 
                    sys.schemas s ON t.schema_id = s.schema_id -- 加入sys.schemas视图来获取架构名
                ORDER BY 
                    TableName, 
                    TriggerName;
            ";
            var result = _connection.Query<TriggerSpecifications>(sql);
            return result;
        }

        public override void Dispose()
        {
            _connection.Dispose();
        }
    }
}

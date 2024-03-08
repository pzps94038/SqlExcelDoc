using Microsoft.Data.SqlClient;
using NPOI.SS.UserModel;
using Org.BouncyCastle.Utilities.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SqlExcelDoc.Model
{
    internal abstract class SqlDoc : IDisposable
    {
        protected string _connectionString;
        public SqlDoc(string connectionString)
        {
            _connectionString = connectionString;
        }

        public abstract void Dispose();

        public abstract IEnumerable<DatabaseSpecifications> GetDatabaseSpecifications();
        public abstract IEnumerable<DatabaseSpecifications> GetDatabaseViewSpecifications();
        public abstract IEnumerable<ProcedureSpecifications> GetStoredProcedureSpecifications();
        public abstract IEnumerable<TableSpecifications> GetTableSpecifications();
        public abstract IEnumerable<TriggerSpecifications> GetTriggerSpecifications();
    }

}

using Dapper;
using ExcelTool.Models.DTO;
using Microsoft.Data.SqlClient;
using System.Data;

namespace ExcelTool.Data
{
    public class SqlRepository : ISqlRepository
    {        
            private readonly IConfiguration _config;

            public SqlRepository(IConfiguration config)
            {
                _config = config;
            }

            public SqlConnection GetConnection()
                => new SqlConnection(_config.GetConnectionString("DefaultConnection"));

            public async Task<List<string>> GetTableColumnsAsync(string tableName, string schema)
            {
                using var conn = GetConnection();

            var sql = """
        SELECT COLUMN_NAME
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = @Schema
          AND TABLE_NAME = @Table
    """;

            var result = await conn.QueryAsync<string>(sql,
                new { Schema = schema, Table = tableName });

            return result.ToList();
        }



        public async Task BulkInsertAsync(string schema, string table, DataTable dataTable)
        {
            using var conn = GetConnection();
            await conn.OpenAsync();

            using var bulk = new SqlBulkCopy(conn)
            {
                DestinationTableName = $"{schema}.{table}"
            };

            foreach (DataColumn col in dataTable.Columns)
            {
                bulk.ColumnMappings.Add(col.ColumnName, col.ColumnName);
            }

            await bulk.WriteToServerAsync(dataTable);
        }


        public async Task<Dictionary<string, (string RefSchema, string RefTable, string RefColumn)>>
       GetForeignKeysAsync(string schema, string table)
        {
            using var conn = GetConnection();

            var sql = @"
        SELECT 
            ccu.COLUMN_NAME  AS FKColumn,
            kcu.TABLE_SCHEMA AS RefSchema,
            kcu.TABLE_NAME   AS RefTable,
            kcu.COLUMN_NAME  AS RefColumn
        FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS rc
        JOIN INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE ccu
            ON rc.CONSTRAINT_NAME = ccu.CONSTRAINT_NAME
        JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE kcu
            ON rc.UNIQUE_CONSTRAINT_NAME = kcu.CONSTRAINT_NAME
        WHERE ccu.TABLE_SCHEMA = @Schema
          AND ccu.TABLE_NAME = @Table;
    ";

            var rows = await conn.QueryAsync<ForeignKeyInfo>(
                sql,
                new { Schema = schema, Table = table });

            var result = new Dictionary<string, (string, string, string)>();

            foreach (var r in rows)
            {
                result[r.FKColumn] = (r.RefSchema, r.RefTable, r.RefColumn);
            }

            return result;
        }


        public async Task<bool> ForeignKeyExistsAsync(string schema,string table, string column,object value)
        {
            using var conn = GetConnection();

            var sql = $@" SELECT COUNT(1) FROM {schema}.{table} WHERE {column} = @Value";

            var count = await conn.ExecuteScalarAsync<int>(sql, new { Value = value });
            return count > 0;
        }



        public async Task BulkUpsertAsync( string schema, string table,DataTable dataTable)
        {
            if (!dataTable.Columns.Contains("RecordNo"))
                throw new Exception("RecordNo column is required for UPSERT.");

            using var conn = GetConnection();
            await conn.OpenAsync();

            using var tran = conn.BeginTransaction();

            try
            {
                string fullTableName = $"[{schema}].[{table}]";

                /* 1️⃣ Create temp table with same structure */
                string createTempTableSql = $@"
            SELECT TOP 0 *
            INTO #TempData
            FROM {fullTableName};";

                using (var cmd = new SqlCommand(createTempTableSql, conn, tran))
                {
                    await cmd.ExecuteNonQueryAsync();
                }

                /* 2️⃣ Bulk copy Excel data into temp table */
                using (var bulk = new SqlBulkCopy(conn, SqlBulkCopyOptions.Default, tran))
                {
                    bulk.DestinationTableName = "#TempData";

                    foreach (DataColumn col in dataTable.Columns)
                    {
                        bulk.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                    }

                    await bulk.WriteToServerAsync(dataTable);
                }

                /* 3️⃣ MERGE with change detection */
               
                string identityColumn  = dataTable.Columns[0].ColumnName;
                string updateSetClause = string.Join(",",
    dataTable.Columns.Cast<DataColumn>()
        .Where(c => c.ColumnName != identityColumn)
        .Select(c => $"target.[{c.ColumnName}] = source.[{c.ColumnName}]")
);

                string changeDetection = string.Join(" OR ",
                    dataTable.Columns.Cast<DataColumn>()
                     .Where(c => c.ColumnName != identityColumn)                       
                        .Select(c =>
                            $"ISNULL(target.[{c.ColumnName}], '') <> ISNULL(source.[{c.ColumnName}], '')")
                );

                string mergeSql = $@"
            MERGE {fullTableName} AS target
            USING #TempData AS source
            ON target.RecordNo = source.RecordNo

            WHEN MATCHED AND ({changeDetection})
            THEN UPDATE SET {updateSetClause}

            WHEN NOT MATCHED BY TARGET
            THEN INSERT ({string.Join(",", dataTable.Columns.Cast<DataColumn>()
                            .Where(c => c.ColumnName != identityColumn)
                            .Select(c => $"[{c.ColumnName}]"))})
            VALUES ({string.Join(",", dataTable.Columns.Cast<DataColumn>()
                           .Where(c => c.ColumnName != identityColumn)
                            .Select(c => $"source.[{c.ColumnName}]"))});";

                using (var cmd = new SqlCommand(mergeSql, conn, tran))
                {
                    await cmd.ExecuteNonQueryAsync();
                }

                tran.Commit();
            }
            catch
            {
                tran.Rollback();
                throw;
            }
        }


    }
}

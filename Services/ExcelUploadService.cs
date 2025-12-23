using Dapper;
using ExcelTool.Data;
using ExcelTool.Helper;
using ExcelTool.Models;
using System.Data;

namespace ExcelTool.Services
{
    public class ExcelUploadService : IExcelUploadService
    {
        private readonly ISqlRepository _repo;

        public ExcelUploadService(ISqlRepository repo)
        {
            _repo = repo;
        }

        public async Task<ExcelUploadResult> UploadAsyncWithoutFK( string tableName, IFormFile excelFile)
        {
            var result = new ExcelUploadResult();

            try
            {
                // 1. Parse schema + table
                var (schema, table) = ParseSchemaAndTable(tableName);

                // 2. Read Excel
                var excelTable = ExcelHelper.ReadExcelWithoutDataCheck(excelFile);

                // 3. Get DB Columns (SCHEMA-AWARE)
                var dbColumns = await _repo.GetTableColumnsAsync(table, schema);

                if (!dbColumns.Any())
                    throw new Exception($"Table '{schema}.{table}' does not exist or has no columns.");

                // 4. Match Columns
                var mapping = ColumnMatcher.MatchColumns(
                    excelTable.Columns.Cast<DataColumn>()
                        .Select(c => c.ColumnName)
                        .ToList(),
                    dbColumns);

                if (!mapping.Any())
                    throw new Exception("No matching columns found between Excel and database table.");

                // 5. Build final DataTable
                DataTable finalTable = new();

                foreach (var dbCol in mapping.Values.Distinct())
                    finalTable.Columns.Add(dbCol);

                foreach (DataRow row in excelTable.Rows)
                {
                    var newRow = finalTable.NewRow();

                    foreach (var map in mapping)
                    {
                        newRow[map.Value] =
                            row.Table.Columns.Contains(map.Key)
                                ? row[map.Key] ?? DBNull.Value
                                : DBNull.Value;
                    }

                    finalTable.Rows.Add(newRow);
                }

                // 6. Bulk Insert (SCHEMA-AWARE)
                await _repo.BulkInsertAsync(schema, table, finalTable);

                result.Success = true;
                result.InsertedRows = finalTable.Rows.Count;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Errors.Add(ex.Message);
            }

            return result;
        }

        public async Task<ExcelUploadResult> UploadAsync( string tableName, IFormFile excelFile)
        {
            var result = new ExcelUploadResult();

            try
            {
                var (schema, table) = ParseSchemaAndTable(tableName);

                var excelTable = ExcelHelper.ReadExcelWithoutDataCheck(excelFile);

                // 3. Get DB Columns (SCHEMA-AWARE)
                var dbColumns = await _repo.GetTableColumnsAsync(table, schema);
           
                if (!dbColumns.Any())
                    throw new Exception($"Table '{schema}.{table}' does not exist.");

                var fkMap = await _repo.GetForeignKeysAsync(schema, table);

                var mapping = ColumnMatcher.MatchColumns(
                    excelTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList(),
                    dbColumns);

                DataTable finalTable = new();
                foreach (var dbCol in mapping.Values.Distinct())
                    finalTable.Columns.Add(dbCol);

                foreach (DataRow row in excelTable.Rows)
                {
                    bool fkValid = true;
                    var newRow = finalTable.NewRow();

                    foreach (var map in mapping)
                    {
                        var value = row[map.Key];

                        // FK validation
                        if (fkMap.ContainsKey(map.Value) && value != DBNull.Value)
                        {
                            var fk = fkMap[map.Value];
                            var exists = await _repo.ForeignKeyExistsAsync(
                                fk.RefSchema, fk.RefTable, fk.RefColumn, value);

                            if (!exists)
                            {
                                fkValid = false;
                                result.Errors.Add(
                                    $"FK violation: {map.Value}={value} not found in {fk.RefSchema}.{fk.RefTable}");
                                break;
                            }
                        }

                        newRow[map.Value] = value ?? DBNull.Value;
                    }

                    if (fkValid)
                        finalTable.Rows.Add(newRow);
                }
                await _repo.BulkUpsertAsync(schema, table, finalTable);
             
                result.Success = true;
                result.InsertedRows = finalTable.Rows.Count;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Errors.Add(ex.Message);
            }

            return result;
        }

        public async Task<bool> ForeignKeyExistsAsync( string schema, string table, string column, object value)
        {
            using var conn = _repo.GetConnection();

            var sql = $@"
        SELECT COUNT(1)
        FROM {schema}.{table}
        WHERE {column} = @Value";

            var count = await conn.ExecuteScalarAsync<int>(sql, new { Value = value });
            return count > 0;
        }

        private static (string Schema, string Table) ParseSchemaAndTable(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                throw new ArgumentException("Table name is required.");

            if (input.Contains('.'))
            {
                var parts = input.Split('.', 2);
                return (parts[0], parts[1]);
            }

            // Default schema fallback
            return ("master", input);
        }


    }

}

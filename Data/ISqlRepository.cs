using Microsoft.Data.SqlClient;
using System.Data;

namespace ExcelTool.Data
{
    public interface ISqlRepository
    {
        Task<List<string>> GetTableColumnsAsync(string schema, string table);
        Task BulkInsertAsync(string schema, string table, DataTable dataTable);
        SqlConnection GetConnection();
        Task<Dictionary<string, (string RefSchema, string RefTable, string RefColumn)>>
     GetForeignKeysAsync(string schema, string table);
        Task<bool> ForeignKeyExistsAsync(string schema, string table, string column, object value);

        Task BulkUpsertAsync(string schema, string table, DataTable dataTable);
    }
}

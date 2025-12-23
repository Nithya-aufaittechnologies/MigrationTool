using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using System.Data;

public static class ExcelHelper
{
    public static DataTable ReadExcelOld(IFormFile file)
    {
        using var stream = new MemoryStream();
        file.CopyTo(stream);
        stream.Position = 0;

        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet(1);

        var table = new DataTable();

        // Header
        var headerRow = worksheet.FirstRowUsed();
        foreach (var cell in headerRow.CellsUsed())
        {
            table.Columns.Add(cell.GetValue<string>().Trim());
        }

        // Data rows
        foreach (var row in worksheet.RowsUsed().Skip(1))
        {
            var dataRow = table.NewRow();
            int colIndex = 0;

            foreach (var cell in row.Cells(1, table.Columns.Count))
            {
                dataRow[colIndex++] = cell.IsEmpty()
                    ? DBNull.Value
                    : cell.GetValue<string>();
            }

            table.Rows.Add(dataRow);
        }

        return table;
    }


    public static DataTable ReadExcelWithoutStatus(IFormFile file)
    {
        using var stream = new MemoryStream();
        file.CopyTo(stream);
        stream.Position = 0;

        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet(1);

        var table = new DataTable();

        // ✅ Use LastColumnUsed to preserve indexes
        var headerRow = worksheet.FirstRowUsed();
        int lastColumn = headerRow.LastCellUsed().Address.ColumnNumber;

        // Read headers by index (NOT CellsUsed)
        for (int col = 1; col <= lastColumn; col++)
        {
            var headerText = headerRow.Cell(col).GetValue<string>().Trim();
            table.Columns.Add(headerText);
        }

        // Read data rows using exact column indexes
        foreach (var row in worksheet.RowsUsed().Skip(1))
        {
            var dataRow = table.NewRow();

            for (int col = 1; col <= lastColumn; col++)
            {
                var cell = row.Cell(col);
                dataRow[col - 1] = cell.IsEmpty()
                    ? DBNull.Value
                    : cell.GetValue<string>();
            }

            table.Rows.Add(dataRow);
        }

        return table;
    }

    public static DataTable ReadExcelWithoutDataCheck(IFormFile file)
    {
        using var stream = new MemoryStream();
        file.CopyTo(stream);
        stream.Position = 0;

        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet(1);

        var table = new DataTable();

        var headerRow = worksheet.FirstRowUsed();
        int lastColumn = headerRow.LastCellUsed().Address.ColumnNumber;

        int statusColumnIndex = -1;
        int customerIdFk = -1;

        // 1️⃣ Read headers and detect Status column
        for (int col = 1; col <= lastColumn; col++)
        {
            var headerText = headerRow.Cell(col).GetValue<string>().Trim();
            table.Columns.Add(headerText);

            if (headerText.Equals("status", StringComparison.OrdinalIgnoreCase))
            {
                statusColumnIndex = col - 1; // DataTable is 0-based
            }
            if (headerText.Equals("uot_sold_party_dp", StringComparison.OrdinalIgnoreCase))
            {
                customerIdFk = col - 1; // DataTable is 0-based
            }
        }

        // 2️⃣ Read data rows
        foreach (var row in worksheet.RowsUsed().Skip(1))
        {
            var dataRow = table.NewRow();

            for (int col = 1; col <= lastColumn; col++)
            {
                var cell = row.Cell(col);
                var columnIndex = col - 1;

                if (columnIndex == statusColumnIndex)
                {
                    // 🔥 STATUS TRANSFORMATION LOGIC
                    if (cell.IsEmpty())
                    {
                        dataRow[columnIndex] = 2;
                    }
                    else
                    {
                        var statusValue = cell.GetValue<string>().Trim();
                        

                        dataRow[columnIndex] =
                            statusValue.Equals("Active", StringComparison.OrdinalIgnoreCase) ? 1 :
                            statusValue.Equals("Terminated", StringComparison.OrdinalIgnoreCase) ? 2 :
                            2; // default fallback
                        
                    }
                }
                
                else if (columnIndex == customerIdFk)
                {
                    var customerIdValue = cell.GetValue<string>().Trim();
                    dataRow[columnIndex] =
                            customerIdValue.Equals("0", StringComparison.OrdinalIgnoreCase) ? null : customerIdValue;
                }
                else
                {
                    dataRow[columnIndex] = cell.IsEmpty()
                        ? DBNull.Value
                        : cell.GetValue<string>();
                }
            }

            table.Rows.Add(dataRow);
        }

        return table;
    } 


    private static string Normalize(string value)
    {
        return value
            .ToLowerInvariant()
            .Replace("_", "")
            .Replace(" ", "")
            .Replace("-", "")
            .Trim();
    }


    public static (DataTable InsertTable, DataTable UpdateTable) ReadExcelWithUpsert(
    IFormFile file,
    DataTable dbTable,           // Existing DB data
    string recordNoColumnName    // e.g. "RecordNo"
)
    {
        using var stream = new MemoryStream();
        file.CopyTo(stream);
        stream.Position = 0;

        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet(1);

        var insertTable = dbTable.Clone();
        var updateTable = dbTable.Clone();

        // Fast lookup by RecordNo
        var dbLookup = dbTable.AsEnumerable()
            .Where(r => !string.IsNullOrWhiteSpace(r[recordNoColumnName]?.ToString()))
            .ToDictionary(r => r[recordNoColumnName].ToString(), r => r);

        var headerRow = worksheet.FirstRowUsed();
        int lastColumn = headerRow.LastCellUsed().Address.ColumnNumber;

        int statusColumnIndex = -1;
        int customerIdFk = -1;
        int recordNoIndex = -1;

        // 1️⃣ Read headers
        for (int col = 1; col <= lastColumn; col++)
        {
            var headerText = headerRow.Cell(col).GetValue<string>().Trim();

            if (headerText.Equals("status", StringComparison.OrdinalIgnoreCase))
                statusColumnIndex = col - 1;

            if (headerText.Equals("uot_sold_party_dp", StringComparison.OrdinalIgnoreCase))
                customerIdFk = col - 1;

            if (headerText.Equals(recordNoColumnName, StringComparison.OrdinalIgnoreCase))
                recordNoIndex = col - 1;
        }

        if (recordNoIndex == -1)
            throw new Exception($"RecordNo column '{recordNoColumnName}' not found in Excel.");

        // 2️⃣ Read data rows
        foreach (var row in worksheet.RowsUsed().Skip(1))
        {
            var recordNo = row.Cell(recordNoIndex + 1).GetValue<string>()?.Trim();
            if (string.IsNullOrWhiteSpace(recordNo))
                continue;

            var newRow = dbTable.NewRow();
            bool exists = dbLookup.TryGetValue(recordNo, out DataRow dbRow);
            bool isChanged = false;

            for (int col = 1; col <= lastColumn; col++)
            {
                var cell = row.Cell(col);
                int columnIndex = col - 1;
                object newValue;

                // 🔥 STATUS LOGIC
                if (columnIndex == statusColumnIndex)
                {
                    if (cell.IsEmpty())
                        newValue = 2;
                    else
                    {
                        var statusValue = cell.GetValue<string>().Trim();
                        newValue =
                            statusValue.Equals("Active", StringComparison.OrdinalIgnoreCase) ? 1 :
                            statusValue.Equals("Terminated", StringComparison.OrdinalIgnoreCase) ? 2 :
                            2;
                    }
                }
                // 🔥 CUSTOMER FK LOGIC
                else if (columnIndex == customerIdFk)
                {
                    var value = cell.GetValue<string>()?.Trim();
                    newValue = value == "0" ? DBNull.Value : value;
                }
                else
                {
                    newValue = cell.IsEmpty()
                        ? DBNull.Value
                        : cell.GetValue<string>();
                }

                newRow[columnIndex] = newValue;

                // 🔥 CHANGE DETECTION
                if (exists)
                {
                    var oldValue = dbRow[columnIndex];
                    if (!Equals(
                            oldValue == DBNull.Value ? null : oldValue,
                            newValue == DBNull.Value ? null : newValue))
                    {
                        isChanged = true;
                    }
                }
            }

            // 🔥 UPSERT DECISION
            if (!exists)
            {
                insertTable.Rows.Add(newRow); // INSERT
            }
            else if (isChanged)
            {
                newRow["ProjectID"] = dbRow["ProjectID"]; // Preserve PK
                updateTable.Rows.Add(newRow);             // UPDATE
            }
            // else → NO CHANGE → IGNORE
        }

        return (insertTable, updateTable);
    }

}

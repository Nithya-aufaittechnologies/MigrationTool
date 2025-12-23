namespace ExcelTool.Models
{
    public class ExcelUploadResult
    {
        public bool Success { get; set; }
        public int InsertedRows { get; set; }
        public List<string> Errors { get; set; } = new();
    }
    public class ExcelUploadRequest
    {
        public string TableName { get; set; } = string.Empty;
        public IFormFile File { get; set; } = default!;
    }
}

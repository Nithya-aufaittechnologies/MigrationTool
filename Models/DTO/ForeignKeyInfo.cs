namespace ExcelTool.Models.DTO
{
    public class ForeignKeyInfo
    {
        public string FKColumn { get; set; } = string.Empty;
        public string RefSchema { get; set; } = string.Empty;
        public string RefTable { get; set; } = string.Empty;
        public string RefColumn { get; set; } = string.Empty;
    }

}

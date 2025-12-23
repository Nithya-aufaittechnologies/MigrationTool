namespace ExcelTool.Helper
{
    public static class ColumnMatcher
    {

        public static Dictionary<string, string> MatchColumns(
                List<string> excelColumns,
                List<string> dbColumns)
        {
            var map = new Dictionary<string, string>();

            // Normalize DB columns once
            var normalizedDb = dbColumns.ToDictionary(
                db => Normalize(db),
                db => db,
                StringComparer.OrdinalIgnoreCase);

            foreach (var excelCol in excelColumns)
            {
                var normalizedExcel = Normalize(excelCol);
                
                var logicalName = "";
                #region CustomerMaster
               
                 logicalName = excelCol switch
                {
                    "uot_sold_party_code_sdt120" => "CompanyCode",
                    "ucm_comp_name_sdt120" => "CompanyName",
                    "ucm_url_hp" => "CompanyURL",
                    "phone_number004" => "WorkPhoneNumber",
                    "uvetaxidtb16" => "GST",
                    "uot_sold_party_dp"=>"CustomerID",

                    //CustomerContacts
                    "uircntctfstnmtb"=> "ContactFirstName",
                    "uricntctlstnmtb" => "ContactLastName",
                    "ugenzipcodetxt16"=> "ZipPostalCode",
                    //Project
                    "us_xcpr_id"=> "ProjectTemplateID",
                    "shellnumber"=> "ProjectNumber",
                    "shellname"=> "ProjectName",                    
                    "pid" => "RecordNo",
                    //VendorMaster
                    "vendor_master_vendor"=> "VendorName",
                    "vendor_master_con_person"=> "ContactPerson",
                    "vendor_master_con_number"=> "ContactNumber",
                    "vendor_master_manu_add"=> "ManufacturingAddress",
                    "vendor_master_code"=> "VendorCode",

                    _ => excelCol   // fallback
                
            }
            ;
                #endregion
                if (normalizedDb.TryGetValue(logicalName, out var dbMatch))
                {
                    map[excelCol] = dbMatch;
                    continue;
                }


                // 2️⃣ Fallback fuzzy match (only for normal columns)
                var fuzzy = normalizedDb.FirstOrDefault(db =>
                    db.Key.Contains(normalizedExcel) ||
                    normalizedExcel.Contains(db.Key));

                if (!string.IsNullOrWhiteSpace(fuzzy.Value))
                {
                    map[excelCol] = fuzzy.Value;
                }
            }

            return map;
        }

        private static string Normalize(string value)
        {
            return value
                .ToLowerInvariant()
                .Replace("_", "")
                .Replace(" ", "")
                .Replace("-", "")
                .Replace("\r", "")
                .Replace("\n", "")
                .Trim();
        }
    }
}
































//    public static Dictionary<string, string> MatchColumns(
//        List<string> excelColumns,
//        List<string> dbColumns)
//    {
//        var map = new Dictionary<string, string>();

//        foreach (var excelCol in excelColumns)
//        {
//            var normalizedExcel = Normalize(excelCol);

//            var match = dbColumns.FirstOrDefault(db =>
//                Normalize(db).Contains(normalizedExcel) ||
//                normalizedExcel.Contains(Normalize(db)));

//            if (match != null)
//                map[excelCol] = match;
//        }

//        return map;
//    }

//    private static string Normalize(string value) =>
//        value.Replace("_", "")
//             .Replace(" ", "")
//             .ToLower();
//}
//        public static Dictionary<string, string> MatchColumns(List<string> excelColumns, List<string> dbColumns)
//        {
//            var map = new Dictionary<string, string>();

//            // Pre-normalize DB columns
//            var normalizedDbColumns = dbColumns.ToDictionary(
//                db => Normalize(db),
//                db => db,
//                StringComparer.OrdinalIgnoreCase);

//            foreach (var excelCol in excelColumns)
//            {
//                var normalizedExcel = Normalize(excelCol);

//                // 1️⃣ Alias-based match (PRIORITY)
//                if (ColumnAliases.TryGetValue(normalizedExcel, out var logicalName))
//                {
//                    if (normalizedDbColumns.TryGetValue(logicalName, out var dbMatch))
//                    {
//                        map[excelCol] = dbMatch;
//                        continue;
//                    }
//                }

//                // 2️⃣ Fallback fuzzy match (existing behavior)
//                var match = normalizedDbColumns
//                    .FirstOrDefault(db =>
//                        db.Key.Contains(normalizedExcel) ||
//                        normalizedExcel.Contains(db.Key));

//                if (!string.IsNullOrEmpty(match.Value))
//                {
//                    map[excelCol] = match.Value;
//                }
//            }

//            return map;
//        }


//        private static string Normalize(string value)
//        {
//            return value
//                .ToLowerInvariant()
//                .Replace("_", "")
//                .Replace("-", "")
//                .Replace(" ", "")
//                .Replace("\r", "")
//                .Replace("\n", "")
//                .Trim();
//        }


//        private static readonly Dictionary<string, string> ColumnAliases =
//    new(StringComparer.OrdinalIgnoreCase)
//{
//    // Company Name aliases
//    { "ucm_comp_name_sdt120", "companyname" },
//    { "ucmcompanynamesdt120", "companyname" },
//    { "companyname", "companyname" },
//    { "company", "companyname" }
//};


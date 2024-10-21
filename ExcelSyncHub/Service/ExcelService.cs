using System.Data;
using System.Drawing;
using System.Text;
using DBUtilityHub.Data;
using Spire.Xls;
using ExcelSyncHub.Model.DTO;
using DBUtilityHub.Models;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using ExcelSyncHub.Models.DTO;
using Newtonsoft.Json;
using ExcelSyncHub.Service.IService;
using Spire.Xls.Core;
using System.Net;
using InfluxDB.Client.Api.Domain;
using System.IO.Packaging;
using Microsoft.Data.SqlClient;
using Npgsql;


namespace ExcelSyncHub.Service
{
    public class ExcelService : IExcelService
    {

        private readonly ApplicationDbContext _context;
        private readonly IHttpContextAccessor _httpContextAccessor;

        public ExcelService(ApplicationDbContext context, IHttpContextAccessor httpContextAccessor)
        {
            _context = context;
            _httpContextAccessor = httpContextAccessor;
        }

        //Generate Excel
        public byte[] GenerateExcelFile(List<ColumnMetaDataDTO> columns, int? parentId , string UserName, string Password, string DataBase, string HostName)
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add the first worksheet with detailed column information
            worksheet.Name = "DataDictionary";

            // Set protection options for the first sheet (read-only)
            worksheet.Protect("your_password", SheetProtectionType.All);
            worksheet.DefaultRowHeight = 15;
            var headerColor = worksheet.Range["A2:O2"];
            headerColor.Style.FillPattern = ExcelPatternType.Solid;
            headerColor.Style.Font.Color = Color.White;
            headerColor.Style.KnownColor = ExcelColors.Blue;
            headerColor.Style.Font.IsBold = true;
            worksheet.DefaultColumnWidth = 15;
            // Add column headers for the first sheet
            worksheet.Range["A2"].Text = "SI.No";
            worksheet.Range["B2"].Text = "Data Item";
            worksheet.Range["C2"].Text = "Data Type";
            //worksheet.Range["D2"].Text = "Length";
            worksheet.Range["D2"].Text = "MinLength";
            worksheet.Range["E2"].Text = "MaxLength";
            worksheet.Range["F2"].Text = "MinRange";
            worksheet.Range["G2"].Text = "MaxRange";
            worksheet.Range["H2"].Text = "DateMinValue";
            worksheet.Range["I2"].Text = "DateMaxValue";
            worksheet.Range["J2"].Text = "Description";
            worksheet.Range["K2"].Text = "Blank value Allowed";
            worksheet.Range["L2"].Text = "Default Value";
            worksheet.Range["M2"].Text = "Unique Value";
            worksheet.Range["N2"].Text = "Option1";
            worksheet.Range["O2"].Text = "Option2";

            // Populate the first sheet with column details
            columns.ForEach(column =>
            {
                int i = columns.IndexOf(column); // Get the index of the current column
                int row = i + 3;
                var range = worksheet.Range[row, 1, row, 15];
                range.Style.Color = (row % 2 == 0) ? Color.LightBlue : Color.Azure;

                SetCellText(worksheet, i + 3, 1, (i + 1).ToString());
                SetCellText(worksheet, i + 3, 2, column.ColumnName);
                SetCellText(worksheet, i + 3, 3, column.Datatype);
                SetCellText(worksheet, i + 3, 4, column.MinLength?.ToString() ?? string.Empty);
                SetCellText(worksheet, i + 3, 5, (column.MaxLength > 0) ? column.MaxLength.ToString() : string.Empty);
                SetCellText(worksheet, i + 3, 6, column.MinRange?.ToString() ?? string.Empty);
                SetCellText(worksheet, i + 3, 7, (column.MaxRange > 0) ? column.MaxRange.ToString() : string.Empty);
                SetCellText(worksheet, i + 3, 8, column.DateMinValue ?? string.Empty);
                SetCellText(worksheet, i + 3, 9, column.DateMaxValue ?? string.Empty);
                SetCellText(worksheet, i + 3, 10, column.Description ?? string.Empty);
                worksheet.Range[i + 3, 11].Text = column.IsNullable.ToString();
                SetCellText(worksheet, i + 3, 12, ((column.Datatype.ToLower() == "boolean") || (column.Datatype.ToLower() == "bit")) ? (string.IsNullOrEmpty(column.DefaultValue) ? string.Empty : ((column.DefaultValue.ToLower() == "true") ? column.True : column.False)) : column.DefaultValue);
                worksheet.Range[i + 3, 13].Text = column.IsPrimaryKey.ToString();
                SetCellText(worksheet, i + 3, 14, column.True ?? string.Empty);
                SetCellText(worksheet, i + 3, 15, column.False ?? string.Empty);

                worksheet.Range["A1"].Text = column.EntityId.ToString();
            });

            worksheet.HideRow(1);

            // Add static content in the last row (vertically)
            var lastRowIndex = worksheet.Rows.Length;
            int startRow = lastRowIndex + 2;
            int endRow = startRow + 7;

            for (int i = startRow; i < endRow; i += 2)
            {
                Color rowColor = (i % 4 == 0) ? Color.LightGray : Color.LightSkyBlue;
                worksheet.Range[i, 1, i, 5].Style.Color = rowColor;
            }
            worksheet.Range[lastRowIndex + 2, 1].Text = "";
            worksheet.Range[lastRowIndex + 3, 1].Text = "Note:";
            worksheet.Range[lastRowIndex + 4, 1].Text = "1. Don't add or delete any columns";
            worksheet.Range[lastRowIndex + 5, 1].Text = "2. Don't add any extra sheets";
            worksheet.Range[lastRowIndex + 6, 1].Text = "3. Follow the length if mentioned";
            if (parentId.HasValue)
            {
                worksheet.Range[lastRowIndex + 7, 1].Text = "4. This is Exported Data ExcelFile";
            }
            var staticContentRange = worksheet.Range[lastRowIndex + 2, 1, lastRowIndex + 7, 5];
            staticContentRange.Style.FillPattern = ExcelPatternType.Solid;
            staticContentRange.Style.KnownColor = ExcelColors.Yellow;
            // Add the second worksheet for column names
            Worksheet columnNamesWorksheet;

            if (!parentId.HasValue)
            {
                columnNamesWorksheet = workbook.Worksheets.Add("Fill data");
            }
            else
            {
                columnNamesWorksheet = workbook.Worksheets.Add("Error Records");
            }

            // After adding content to the columns
            // Set a default column width for the "Fill data" worksheet
            columnNamesWorksheet.DefaultColumnWidth = 15; // Set the width in characters (adjust as needed)
            columnNamesWorksheet.DefaultRowHeight = 15;
            int lastColumnIndex = columns.Count + 1;

            for (int i = 0; i < columns.Count; i++)
            {
                var column = columns[i];
                columnNamesWorksheet.Range[2, i + 1].Text = column.ColumnName;
                int entityId = column.EntityId;
                columnNamesWorksheet.Range["A1"].Text = entityId.ToString();
                columnNamesWorksheet.Range[2, 1, 2, lastColumnIndex - 1].Style.Color = Color.Blue;
                columnNamesWorksheet.Range[2, 1, 2, lastColumnIndex - 1].Style.Font.IsBold = true;
                columnNamesWorksheet.Range[2, 1, 2, lastColumnIndex - 1].Style.Font.Color = Color.White;
            }
            columnNamesWorksheet.HideRow(1);

            if (parentId.HasValue)
            {
                InsertDataIntoExcel(columnNamesWorksheet, columns, parentId);
            }

            string[] sheetsToRemove = { "Sheet2", "Sheet3" }; // Names of sheets to be removed
            foreach (var sheetName in sheetsToRemove)
            {
                Worksheet sheetToRemove = workbook.Worksheets[sheetName];
                if (sheetToRemove != null)
                {
                    workbook.Worksheets.Remove(sheetToRemove);
                }
            }

            AddDataValidation(columnNamesWorksheet, columns, parentId);


            Worksheet foreignKeyWorksheet = workbook.Worksheets.Add("ForeignKeyData");
            foreignKeyWorksheet.DefaultColumnWidth = 15;
            foreignKeyWorksheet.DefaultRowHeight = 15;

            // Add headers for foreign key data
            foreignKeyWorksheet.Range["A1"].Text = "Foreign Key Column";
            foreignKeyWorksheet.Range["B1"].Text = "Referenced Table";
            foreignKeyWorksheet.Range["C1"].Text = "Referenced Column";
            foreignKeyWorksheet.Range["A1:C1"].Style.Font.IsBold = true;
            foreignKeyWorksheet.Range["A1:C1"].Style.Color = Color.Blue;
            foreignKeyWorksheet.Range["A1:C1"].Style.Font.Color = Color.White;

            // Example: Populate the sheet with foreign key data (this data would likely come from another data source)
            // You need to replace this part with actual foreign key data
            //var foreignKeyData = new List<(string ForeignKeyColumn, string ReferencedTable, string ReferencedColumn)>
            //    {
            //        ("CustomerId", "Customer", "Id"),
            //        ("OrderId", "Order", "Id"),
            //        ("ProductId", "Product", "Id")
            //    };

            var foreignKeyData = FetchForeignKeyDataFromColumnMetaData(UserName, Password, DataBase, HostName);


            for (int i = 0; i < foreignKeyData.Count; i++)
            {
                var data = foreignKeyData[i];
                int row = i + 2; // Start populating from the second row
                foreignKeyWorksheet.Range[row, 1].Text = data.ForeignKeyColumn;
                foreignKeyWorksheet.Range[row, 2].Text = data.ReferencedTable;
                foreignKeyWorksheet.Range[row, 3].Text = data.ReferencedColumn;
            }

            using (MemoryStream memoryStream = new MemoryStream())
            {
                workbook.SaveToStream(memoryStream, FileFormat.Version2013);
                return memoryStream.ToArray();
            }

        }


private List<(string ForeignKeyColumn, string ReferencedTable, string ReferencedColumn)> FetchForeignKeyDataFromColumnMetaData(string UserName, string Password, string DataBase, string HostName)
    {
        var foreignKeyData = new List<(string ForeignKeyColumn, string ReferencedTable, string ReferencedColumn)>();

        // PostgreSQL connection string
        var connectionString = $"Host={HostName};Database={DataBase};Username={UserName};Password={Password};SSL Mode=Prefer;Trust Server Certificate=True;";

        using (var connection = new NpgsqlConnection(connectionString))
        {
            connection.Open();
            string query = @"
            SELECT 
                cm.""EntityId"" AS ForeignKeyColumn,
                rt.""ReferenceColumnID"" AS ReferencedColumn,
                rt.""ReferenceEntityID"" AS ReferencedEntityID
            FROM 
                ""ColumnMetaDataEntity"" cm
            JOIN 
                ""ColumnMetaDataEntity"" rt ON cm.""ReferenceEntityID"" = rt.""EntityId""
            WHERE 
                cm.""ReferenceEntityID"" IS NOT NULL";

            using (var command = new NpgsqlCommand(query, connection))
            {
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var foreignKeyColumn = reader["ForeignKeyColumn"].ToString();
                        var referencedTable = reader["ReferencedColumn"].ToString(); // You may want to get this from another query or table if needed.
                        var referencedColumn = reader["ReferencedEntityID"].ToString();

                        foreignKeyData.Add((foreignKeyColumn, referencedTable, referencedColumn));
                    }
                }
            }
        }

        return foreignKeyData;
    }




    // Helper method to set cell text
    private static void SetCellText(IWorksheet worksheet, int row, int col, string value)
        {
            worksheet.Range[row, col].Text = string.IsNullOrEmpty(value) ? string.Empty : value;
        }

        //HighLight Formula
        private void HighlightDuplicates(Worksheet sheet, int columnNumber, int startRow, int endRow)
        {
            string columnLetter = GetExcelColumnName(columnNumber);

            string range = $"{columnLetter}{startRow}:{columnLetter}{endRow}";
            ConditionalFormatWrapper format = sheet.Range[range].ConditionalFormats.AddCondition();
            format.FormatType = ConditionalFormatType.DuplicateValues;
            format.BackColor = Color.IndianRed;
        }

        // Validation formulas
        private void AddDataValidation(Worksheet columnNamesWorksheet, List<ColumnMetaDataDTO> columns, int? parentId)
        {
            int startRow = 2; // The first row where you want validation
            int endRow = 100000;  // Adjust the last row as needed
            //Export Excel
            if (parentId.HasValue)
            {
                int columnCount = columnNamesWorksheet.Columns.Length - 2;
                char letter = 'A';
                char lastletter = (char)('A' + columnCount - 1);
                // Protect the worksheet with a password
                columnNamesWorksheet.Protect("password");

                for (int col = 1; col <= columnCount; col++)
                {
                    string dataType = columns[col - 1]?.Datatype ?? string.Empty;
                    int length = columns[col - 1]?.Length ?? 0;
                    bool isPrimaryKey = columns[col - 1]?.IsPrimaryKey ?? false;
                    string truevalue = columns[col - 1]?.True ?? string.Empty;
                    string falsevalue = columns[col - 1]?.False ?? string.Empty;
                    bool isNullable = columns[col - 1]?.IsNullable ?? false;
                    int minRange = columns[col - 1]?.MinRange ?? 0;
                    int maxRange = columns[col - 1]?.MaxRange ?? 0;
                    int minLength = columns[col - 1]?.MinLength ?? 0;
                    int maxLength = columns[col - 1]?.MaxLength ?? 0;
                    string dateMinValue = columns[col - 1]?.DateMinValue ?? string.Empty;
                    string dateMaxValue = columns[col - 1]?.DateMaxValue ?? string.Empty;

                    // Specify the range for data validation
                    CellRange range = columnNamesWorksheet.Range[startRow, col, endRow, col];
                    Validation validation = range.DataValidation;


                    //Protect the worksheet with password
                    //columnNamesWorksheet.Protect("123456", SheetProtectionType.All);
                    List<string> integerTypes = new List<string> { "int", "integer", "number", "numeric", "smallint" /* add more as needed */ };
                    List<string> stringTypes = new List<string> { "string", "varchar", "nvarchar", "text", "character varying"/* add more as needed */ };
                    List<string> timestampTypes = new List<string> { "timestamp", "datetime", "timestamp without time zone", "date", "timestamp with time zone" /* add more as needed */ };



                    if (stringTypes.Contains(dataType, StringComparer.OrdinalIgnoreCase))
                    {
                        // Text validation with min and max length
                        validation.CompareOperator = ValidationComparisonOperator.Between;
                        if ((minLength == 0) && (maxLength == 0))
                        {
                            // Handle the case when both minimum and maximum length are 0
                            validation.Formula1 = "0";
                            validation.Formula2 = "0";
                        }
                        else if ((!string.IsNullOrEmpty(minLength.ToString()) || minLength == 0) && (string.IsNullOrEmpty(maxLength.ToString()) || maxLength == 0))
                        {
                            // Minimum length provided, no maximum length
                            validation.Formula1 = minLength.ToString();
                            validation.AllowType = CellDataType.TextLength;
                            validation.InputTitle = "Input Data";
                            validation.InputMessage = $"Enter a value with a minimum length of {validation.Formula1} characters.";
                            validation.ErrorTitle = "Error";
                            validation.ErrorMessage = $"The value must have a minimum length of {validation.Formula1} characters.";
                        }
                        else if ((string.IsNullOrEmpty(minLength.ToString()) || minLength == 0) && (!string.IsNullOrEmpty(maxLength.ToString()) || maxLength == 0))
                        {
                            validation.Formula2 = maxLength.ToString();
                            validation.AllowType = CellDataType.TextLength;
                            validation.InputTitle = "Input Data";
                            validation.InputMessage = $"Type text with a maximum length of {validation.Formula2} characters.";
                            validation.ErrorTitle = "Error";
                            validation.ErrorMessage = "The entered value exceeds the allowed length.";
                        }
                        else
                        {
                            // Both minimum and maximum length provided
                            validation.Formula1 = minLength.ToString();
                            validation.Formula2 = maxLength.ToString();
                            validation.AllowType = CellDataType.TextLength;
                            validation.InputTitle = "Input Data";
                            validation.InputMessage = $"Type text with a length between {validation.Formula1} and {validation.Formula2} characters.";
                            validation.ErrorTitle = "Error";
                            validation.ErrorMessage = "Entered value should be within the specified length range.";
                        }
                        if (isPrimaryKey)
                        {
                            HighlightDuplicates(columnNamesWorksheet, col, startRow, endRow);
                            validation.InputMessage = $"Enter the unique value.";
                            validation.ErrorTitle = "Error";
                            validation.ErrorMessage = "Entered Values must be unique";
                        }
                    }
                    else if (integerTypes.Contains(dataType, StringComparer.OrdinalIgnoreCase))
                    {
                        validation.CompareOperator = ValidationComparisonOperator.Between;
                        if (isPrimaryKey)
                        {
                            HighlightDuplicates(columnNamesWorksheet, col, startRow, endRow);

                            validation.CompareOperator = ValidationComparisonOperator.Between;
                            if ((minRange == 0) && (maxRange == 0))
                            {
                                validation.Formula1 = "1";
                                validation.Formula2 = int.MaxValue.ToString();
                                validation.AllowType = CellDataType.Integer;
                                validation.InputTitle = "Input Data";
                                validation.InputMessage = "Enter an integer from 1 ";
                                validation.ErrorTitle = "Error";
                                validation.ErrorMessage = "The value should be an greater than or equal to 1 ";
                            }
                            else if ((!string.IsNullOrEmpty(minRange.ToString()) || minRange == 0) && (string.IsNullOrEmpty(maxRange.ToString()) || maxRange == 0))
                            {
                                // Minimum value provided, no maximum value
                                validation.Formula1 = minRange.ToString();
                                validation.Formula2 = int.MaxValue.ToString();
                                validation.AllowType = CellDataType.Integer;
                                validation.InputTitle = "Input Data";
                                validation.InputMessage = $"Enter a value with a minimum value of {validation.Formula1}.";
                                validation.ErrorTitle = "Error";
                                validation.ErrorMessage = $"The value must be at least {validation.Formula1}.";
                            }
                            else if ((string.IsNullOrEmpty(minRange.ToString()) || minRange == 0) && (!string.IsNullOrEmpty(maxRange.ToString()) || maxRange == 0))
                            {
                                validation.Formula1 = "1";
                                validation.Formula2 = maxRange.ToString();
                                validation.AllowType = CellDataType.Integer;
                                validation.InputTitle = "Input Data";
                                validation.InputMessage = $"Enter an integer value between 1 to {validation.Formula2}.";
                                validation.ErrorTitle = "Error";
                                validation.ErrorMessage = "The entered value exceeds the allowed range.";
                            }
                            else
                            {
                                // Both minimum and maximum values provided
                                validation.Formula1 = minRange.ToString();
                                validation.Formula2 = maxRange.ToString();
                                validation.AllowType = CellDataType.Integer;
                                validation.InputTitle = "Input Data";
                                validation.InputMessage = $"Enter an integer between {validation.Formula1} and {validation.Formula2}.";
                                validation.ErrorTitle = "Error";
                                validation.ErrorMessage = "The value should be within the specified range.";
                            }
                        }
                        else if ((minRange == 0) && (maxRange == 0))
                        {
                            // Handle the case when both minimum and maximum length are 0
                            validation.CompareOperator = ValidationComparisonOperator.Between;
                            validation.Formula1 = int.MinValue.ToString();
                            validation.Formula2 = int.MaxValue.ToString();
                            validation.AllowType = CellDataType.Integer;
                            validation.InputTitle = "Input Data";
                            validation.InputMessage = "Enter an integer.";
                            validation.ErrorTitle = "Error";
                            validation.ErrorMessage = "The value should be an integer ";
                        }
                        else if ((!string.IsNullOrEmpty(minRange.ToString()) || minRange == 0) && (string.IsNullOrEmpty(maxRange.ToString()) || maxRange == 0))
                        {
                            // Minimum value provided, no maximum value
                            validation.Formula1 = minRange.ToString();
                            validation.Formula2 = int.MaxValue.ToString();
                            validation.AllowType = CellDataType.Integer;
                            validation.InputTitle = "Input Data";
                            validation.InputMessage = $"Enter a value with a minimum value of {validation.Formula1}.";
                            validation.ErrorTitle = "Error";
                            validation.ErrorMessage = $"The value must be at least {validation.Formula1}.";
                        }
                        else if ((string.IsNullOrEmpty(minRange.ToString()) || minRange == 0) && (!string.IsNullOrEmpty(maxRange.ToString()) || maxRange == 0))
                        {
                            validation.Formula1 = int.MinValue.ToString();
                            validation.Formula2 = maxRange.ToString();
                            validation.AllowType = CellDataType.Integer;
                            validation.InputTitle = "Input Data";
                            validation.InputMessage = $"Enter an integer value less than or equal to {validation.Formula2}.";
                            validation.ErrorTitle = "Error";
                            validation.ErrorMessage = "The entered value exceeds the allowed range.";
                        }
                        else
                        {
                            // Both minimum and maximum values provided
                            validation.Formula1 = minRange.ToString();
                            validation.Formula2 = maxRange.ToString();
                            validation.AllowType = CellDataType.Integer;
                            validation.InputTitle = "Input Data";
                            validation.InputMessage = $"Enter an integer between {validation.Formula1} and {validation.Formula2}.";
                            validation.ErrorTitle = "Error";
                            validation.ErrorMessage = "The value should be within the specified range.";
                        }

                    }
                    else if (dataType.Equals("Date", StringComparison.OrdinalIgnoreCase) || timestampTypes.Contains(dataType, StringComparer.OrdinalIgnoreCase))
                    {
                        // Date validation
                        validation.CompareOperator = ValidationComparisonOperator.Between;

                        if (string.IsNullOrEmpty(dateMinValue) && string.IsNullOrEmpty(dateMaxValue))
                        {
                            // No minimum and maximum date values provided
                            validation.Formula1 = "1757-01-01";
                            validation.Formula2 = "9999-01-01";
                        }
                        else if (!string.IsNullOrEmpty(dateMinValue) && string.IsNullOrEmpty(dateMaxValue))
                        {
                            // Minimum date value provided, no maximum date value
                            validation.Formula1 = dateMinValue;
                            validation.Formula2 = "9999-01-01";
                        }
                        else if (string.IsNullOrEmpty(dateMinValue) && !string.IsNullOrEmpty(dateMaxValue))
                        {
                            // No minimum date value, maximum date value provided
                            validation.Formula1 = "1757-01-01";
                            validation.Formula2 = dateMaxValue;
                        }
                        else
                        {
                            // Both minimum and maximum date values provided
                            validation.Formula1 = dateMinValue;
                            validation.Formula2 = dateMaxValue;
                        }

                        validation.AllowType = CellDataType.Date;
                        validation.InputTitle = "Input Data";
                        validation.InputMessage = $"Type a date between {validation.Formula1} and {validation.Formula2} in this cell.";
                        validation.ErrorTitle = "Error";
                        validation.ErrorMessage = "Enter a valid date with correct format (yyyy-MM-dd).";

                        // Ensure the date format is not avoided
                        var cellRange = range.Worksheet.Range[range.Row, range.Column];
                        cellRange.NumberFormat = "yyyy-MM-dd";
                    }
                    else if ((dataType.Equals("boolean", StringComparison.OrdinalIgnoreCase)) || (dataType.Equals("bit", StringComparison.OrdinalIgnoreCase)))
                    {
                        if (string.IsNullOrEmpty(truevalue) && string.IsNullOrEmpty(falsevalue))
                        {
                            // No specific values provided, allow "true" and "false"
                            validation.Values = new string[] { "true", "false" };
                            validation.ErrorTitle = "Error";
                            validation.InputTitle = "Input Data";
                            validation.ErrorMessage = "Select values from dropdown";
                            validation.InputMessage = "Select values from dropdown";
                        }
                        else
                        {
                            // Specific values provided, enforce dropdown validation
                            validation.Values = new string[] { truevalue, falsevalue };
                            validation.ErrorTitle = "Error";
                            validation.InputTitle = "Input Data";
                            validation.ErrorMessage = "Select values from dropdown";
                            validation.InputMessage = "Select values from dropdown";
                        }
                    }
                    else if (dataType.Equals("char", StringComparison.OrdinalIgnoreCase))
                    {
                        // Character validation for a single character
                        validation.CompareOperator = ValidationComparisonOperator.Between;
                        validation.Formula1 = "1";
                        validation.Formula2 = "1";
                        validation.AllowType = CellDataType.TextLength;
                        validation.InputTitle = "Input Data";
                        validation.InputMessage = "Type a single character.";
                        validation.ErrorTitle = "Error";
                        validation.ErrorMessage = "Enter a valid character.";
                    }
                    else if (dataType.Equals("bytea", StringComparison.OrdinalIgnoreCase))
                    {
                        // Byte validation
                        // Modify the validation code for bytea data
                        validation.CompareOperator = ValidationComparisonOperator.Between;
                        validation.Formula1 = "1"; // Set a minimum length of 1
                        validation.Formula2 = "1000000"; // Set a maximum length as needed
                        validation.AllowType = CellDataType.TextLength;
                        validation.InputTitle = "Input Data";
                        validation.InputMessage = "Type a byte array with a length between 1 and 1000000 characters.";
                        validation.ErrorTitle = "Error";
                        validation.ErrorMessage = "Invalid byte array length";
                        // Include byte validation
                        bool isValidByteA = IsValidByteA(columns[col - 1].DefaultValue, 1, 1000000); // Modify the length limits as needed
                        if (!isValidByteA)
                        {
                            // Data does not meet byte validation criteria
                            validation.ErrorMessage = "Invalid byte array format or length.";
                        }
                    }
                }
                for (int i = 3; i <= 65537; i++)
                {
                    string startindex = letter + i.ToString();
                    string endindex = lastletter + i.ToString();
                    CellRange lockrange = columnNamesWorksheet.Range[startindex + ":" + endindex];
                    lockrange.Style.Locked = false;
                }
            }
            //Normal Excel
            else
            {
                int columnCount = columnNamesWorksheet.Columns.Length;
                char letter = 'A';
                char lastletter = (char)('A' + columnCount - 1);
                // Protect the worksheet with a password
                columnNamesWorksheet.Protect("password");

                for (int col = 1; col <= columnCount; col++)
                {
                    // Get the data type for the current column
                    string dataType = columns[col - 1]?.Datatype ?? string.Empty;
                    int length = columns[col - 1]?.Length ?? 0;
                    bool isPrimaryKey = columns[col - 1]?.IsPrimaryKey ?? false;
                    string truevalue = columns[col - 1]?.True ?? string.Empty;
                    string falsevalue = columns[col - 1]?.False ?? string.Empty;
                    bool isNullable = columns[col - 1]?.IsNullable ?? false;
                    int minRange = columns[col - 1]?.MinRange ?? 0;
                    int maxRange = columns[col - 1]?.MaxRange ?? 0;
                    int minLength = columns[col - 1]?.MinLength ?? 0;
                    int maxLength = columns[col - 1]?.MaxLength ?? 0;
                    string dateMinValue = columns[col - 1]?.DateMinValue ?? string.Empty;
                    string dateMaxValue = columns[col - 1]?.DateMaxValue ?? string.Empty;

                    // Specify the range for data validation
                    CellRange range = columnNamesWorksheet.Range[startRow, col, endRow, col];
                    Validation validation = range.DataValidation;
                    ////Protect the worksheet with password
                    //columnNamesWorksheet.Protect("123456", SheetProtectionType.All);
                    // Check if the current column has a data type of "ListOfValues"
                    List<string> integerTypes = new List<string> { "int", "integer", "number", "numeric", "smallint" /* add more as needed */ };
                    List<string> stringTypes = new List<string> { "string", "varchar", "nvarchar", "text", "character varying"/* add more as needed */ };
                    List<string> timestampTypes = new List<string> { "timestamp", "datetime", "timestamp without time zone", "date" , "timestamp with time zone" /* add more as needed */ };

                    if (stringTypes.Contains(dataType, StringComparer.OrdinalIgnoreCase))
                    {
                        // Text validation with min and max length
                        validation.CompareOperator = ValidationComparisonOperator.Between;
                        if ((minLength == 0) && (maxLength == 0))
                        {
                            // Handle the case when both minimum and maximum length are 0
                            validation.Formula1 = "0";
                            validation.Formula2 = "0";
                        }
                        else if ((!string.IsNullOrEmpty(minLength.ToString()) || minLength == 0) && (string.IsNullOrEmpty(maxLength.ToString()) || maxLength == 0))
                        {
                            // Minimum length provided, no maximum length
                            validation.Formula1 = minLength.ToString();
                            validation.AllowType = CellDataType.TextLength;
                            validation.InputTitle = "Input Data";
                            validation.InputMessage = $"Enter a value with a minimum length of {validation.Formula1} characters.";
                            validation.ErrorTitle = "Error";
                            validation.ErrorMessage = $"The value must have a minimum length of {validation.Formula1} characters.";
                        }
                        else if ((string.IsNullOrEmpty(minLength.ToString()) || minLength == 0) && (!string.IsNullOrEmpty(maxLength.ToString()) || maxLength == 0))
                        {
                            validation.Formula2 = maxLength.ToString();
                            validation.AllowType = CellDataType.TextLength;
                            validation.InputTitle = "Input Data";
                            validation.InputMessage = $"Type text with a maximum length of {validation.Formula2} characters.";
                            validation.ErrorTitle = "Error";
                            validation.ErrorMessage = "The entered value exceeds the allowed length.";
                        }
                        else
                        {
                            // Both minimum and maximum length provided
                            validation.Formula1 = minLength.ToString();
                            validation.Formula2 = maxLength.ToString();
                            validation.AllowType = CellDataType.TextLength;
                            validation.InputTitle = "Input Data";
                            validation.InputMessage = $"Type text with a length between {validation.Formula1} and {validation.Formula2} characters.";
                            validation.ErrorTitle = "Error";
                            validation.ErrorMessage = "Entered value should be within the specified length range.";
                        }
                        if (isPrimaryKey)
                        {
                            HighlightDuplicates(columnNamesWorksheet, col, startRow, endRow);
                            validation.InputMessage = $"Enter the unique value.";
                            validation.ErrorTitle = "Error";
                            validation.ErrorMessage = "Entered Values must be unique";
                        }
                    }
                    else if (integerTypes.Contains(dataType, StringComparer.OrdinalIgnoreCase))
                    {
                        validation.CompareOperator = ValidationComparisonOperator.Between;
                        if (isPrimaryKey)
                        {
                            HighlightDuplicates(columnNamesWorksheet, col, startRow, endRow);

                            validation.CompareOperator = ValidationComparisonOperator.Between;
                            if ((minRange == 0) && (maxRange == 0))
                            {
                                validation.Formula1 = "1";
                                validation.Formula2 = int.MaxValue.ToString();
                                validation.AllowType = CellDataType.Integer;
                                validation.InputTitle = "Input Data";
                                validation.InputMessage = "Enter an integer from 1 ";
                                validation.ErrorTitle = "Error";
                                validation.ErrorMessage = "The value should be an greater than or equal to 1 ";
                            }
                            else if ((!string.IsNullOrEmpty(minRange.ToString()) || minRange == 0) && (string.IsNullOrEmpty(maxRange.ToString()) || maxRange == 0))
                            {
                                // Minimum value provided, no maximum value
                                validation.Formula1 = minRange.ToString();
                                validation.Formula2 = int.MaxValue.ToString();
                                validation.AllowType = CellDataType.Integer;
                                validation.InputTitle = "Input Data";
                                validation.InputMessage = $"Enter a value with a minimum value of {validation.Formula1}.";
                                validation.ErrorTitle = "Error";
                                validation.ErrorMessage = $"The value must be at least {validation.Formula1}.";
                            }
                            else if ((string.IsNullOrEmpty(minRange.ToString()) || minRange == 0) && (!string.IsNullOrEmpty(maxRange.ToString()) || maxRange == 0))
                            {
                                validation.Formula1 = "1";
                                validation.Formula2 = maxRange.ToString();
                                validation.AllowType = CellDataType.Integer;
                                validation.InputTitle = "Input Data";
                                validation.InputMessage = $"Enter an integer value between 1 to {validation.Formula2}.";
                                validation.ErrorTitle = "Error";
                                validation.ErrorMessage = "The entered value exceeds the allowed range.";
                            }
                            else
                            {
                                // Both minimum and maximum values provided
                                validation.Formula1 = minRange.ToString();
                                validation.Formula2 = maxRange.ToString();
                                validation.AllowType = CellDataType.Integer;
                                validation.InputTitle = "Input Data";
                                validation.InputMessage = $"Enter an integer between {validation.Formula1} and {validation.Formula2}.";
                                validation.ErrorTitle = "Error";
                                validation.ErrorMessage = "The value should be within the specified range.";
                            }
                        }
                        else if ((minRange == 0) && (maxRange == 0))
                        {
                            // Handle the case when both minimum and maximum length are 0
                            validation.CompareOperator = ValidationComparisonOperator.Between;
                            validation.Formula1 = int.MinValue.ToString();
                            validation.Formula2 = int.MaxValue.ToString();
                            validation.AllowType = CellDataType.Integer;
                            validation.InputTitle = "Input Data";
                            validation.InputMessage = "Enter an integer.";
                            validation.ErrorTitle = "Error";
                            validation.ErrorMessage = "The value should be an integer ";
                        }
                        else if ((!string.IsNullOrEmpty(minRange.ToString()) || minRange == 0) && (string.IsNullOrEmpty(maxRange.ToString()) || maxRange == 0))
                        {
                            // Minimum value provided, no maximum value
                            validation.Formula1 = minRange.ToString();
                            validation.Formula2 = int.MaxValue.ToString();
                            validation.AllowType = CellDataType.Integer;
                            validation.InputTitle = "Input Data";
                            validation.InputMessage = $"Enter a value with a minimum value of {validation.Formula1}.";
                            validation.ErrorTitle = "Error";
                            validation.ErrorMessage = $"The value must be at least {validation.Formula1}.";
                        }
                        else if ((string.IsNullOrEmpty(minRange.ToString()) || minRange == 0) && (!string.IsNullOrEmpty(maxRange.ToString()) || maxRange == 0))
                        {
                            validation.Formula1 = int.MinValue.ToString();
                            validation.Formula2 = maxRange.ToString();
                            validation.AllowType = CellDataType.Integer;
                            validation.InputTitle = "Input Data";
                            validation.InputMessage = $"Enter an integer value less than or equal to {validation.Formula2}.";
                            validation.ErrorTitle = "Error";
                            validation.ErrorMessage = "The entered value exceeds the allowed range.";
                        }
                        else
                        {
                            // Both minimum and maximum values provided
                            validation.Formula1 = minRange.ToString();
                            validation.Formula2 = maxRange.ToString();
                            validation.AllowType = CellDataType.Integer;
                            validation.InputTitle = "Input Data";
                            validation.InputMessage = $"Enter an integer between {validation.Formula1} and {validation.Formula2}.";
                            validation.ErrorTitle = "Error";
                            validation.ErrorMessage = "The value should be within the specified range.";
                        }

                    }
                    else if (dataType.Equals("Date", StringComparison.OrdinalIgnoreCase) || timestampTypes.Contains(dataType, StringComparer.OrdinalIgnoreCase))
                    {
                        // Date validation
                        validation.CompareOperator = ValidationComparisonOperator.Between;

                        if (string.IsNullOrEmpty(dateMinValue) && string.IsNullOrEmpty(dateMaxValue))
                        {
                            // No minimum and maximum date values provided
                            validation.Formula1 = "1757-01-01";
                            validation.Formula2 = "9999-01-01";
                        }
                        else if (!string.IsNullOrEmpty(dateMinValue) && string.IsNullOrEmpty(dateMaxValue))
                        {
                            // Minimum date value provided, no maximum date value
                            validation.Formula1 = dateMinValue;
                            validation.Formula2 = "9999-01-01";
                        }
                        else if (string.IsNullOrEmpty(dateMinValue) && !string.IsNullOrEmpty(dateMaxValue))
                        {
                            // No minimum date value, maximum date value provided
                            validation.Formula1 = "1757-01-01";
                            validation.Formula2 = dateMaxValue;
                        }
                        else
                        {
                            // Both minimum and maximum date values provided
                            validation.Formula1 = dateMinValue;
                            validation.Formula2 = dateMaxValue;
                        }

                        validation.AllowType = CellDataType.Date;
                        validation.InputTitle = "Input Data";
                        validation.InputMessage = $"Type a date between {validation.Formula1} and {validation.Formula2} in this cell.";
                        validation.ErrorTitle = "Error";
                        validation.ErrorMessage = "Enter a valid date with correct format (yyyy-MM-dd).";

                        // Ensure the date format is not avoided
                        var cellRange = range.Worksheet.Range[range.Row, range.Column];
                        cellRange.NumberFormat = "yyyy-MM-dd";
                    }
                    else if ((dataType.Equals("boolean", StringComparison.OrdinalIgnoreCase)) || (dataType.Equals("bit", StringComparison.OrdinalIgnoreCase)))
                    {
                        if (string.IsNullOrEmpty(truevalue) && string.IsNullOrEmpty(falsevalue))
                        {
                            // No specific values provided, allow "true" and "false"
                            validation.Values = new string[] { "true", "false" };
                            validation.ErrorTitle = "Error";
                            validation.InputTitle = "Input Data";
                            validation.ErrorMessage = "Select values from dropdown";
                            validation.InputMessage = "Select values from dropdown";
                        }
                        else
                        {
                            // Specific values provided, enforce dropdown validation
                            validation.Values = new string[] { truevalue, falsevalue };
                            validation.ErrorTitle = "Error";
                            validation.InputTitle = "Input Data";
                            validation.ErrorMessage = "Select values from dropdown";
                            validation.InputMessage = "Select values from dropdown";
                        }
                    }
                    else if (dataType.Equals("char", StringComparison.OrdinalIgnoreCase))
                    {
                        // Character validation for a single character
                        validation.CompareOperator = ValidationComparisonOperator.Between;
                        validation.Formula1 = "1";
                        validation.Formula2 = "1";
                        validation.AllowType = CellDataType.TextLength;
                        validation.InputTitle = "Input Data";
                        validation.InputMessage = "Type a single character.";
                        validation.ErrorTitle = "Error";
                        validation.ErrorMessage = "Enter a valid character.";
                    }
                    else if (dataType.Equals("bytea", StringComparison.OrdinalIgnoreCase))
                    {
                        // Byte validation
                        // Modify the validation code for bytea data
                        validation.CompareOperator = ValidationComparisonOperator.Between;
                        validation.Formula1 = "1"; // Set a minimum length of 1
                        validation.Formula2 = "1000000"; // Set a maximum length as needed
                        validation.AllowType = CellDataType.TextLength;
                        validation.InputTitle = "Input Data";
                        validation.InputMessage = "Type a byte array with a length between 1 and 1000000 characters.";
                        validation.ErrorTitle = "Error";
                        validation.ErrorMessage = "Invalid byte array length";
                        // Include byte validation
                        bool isValidByteA = IsValidByteA(columns[col - 1].DefaultValue, 1, 1000000); // Modify the length limits as needed
                        if (!isValidByteA)
                        {
                            // Data does not meet byte validation criteria
                            validation.ErrorMessage = "Invalid byte array format or length.";
                        }
                    }

                }
                for (int i = 3; i <= 65537; i++)
                {
                    string startindex = letter + i.ToString();
                    string endindex = lastletter + i.ToString();
                    CellRange lockrange = columnNamesWorksheet.Range[startindex + ":" + endindex];
                    lockrange.Style.Locked = false;
                }
            }
        }

        //Get Column names
        public string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        // Export Error Datas
        private async Task InsertDataIntoExcel(Worksheet columnNamesWorksheet, List<ColumnMetaDataDTO> columns, int? parentId)
        {
            try
            {
                var logChilds = await GetAllLogChildsByParentIDAsync(parentId.Value);
                int rowIndex = 3;


                    // Set the column name as "ErrorMessage" for the last column after processing all rows
                    columnNamesWorksheet.Range[rowIndex - 1, columns.Count + 1].Text = "ErrorMessage";
                    columnNamesWorksheet.Range[rowIndex - 1, columns.Count + 2].Text = "ErrorRowNumber";

                foreach (var logChild in logChilds)
                {
                    string errorMessage = logChild.ErrorMessage;
                    string errorrowNumber = logChild.ErrorRowNumber;

                    string[] errorRowNumbers = errorrowNumber.Split(',');

                    string[] rows = logChild.Filedata.Split(';');

                    for (int i = 1; i < rows.Length; i++)
                    {
                        if (string.IsNullOrWhiteSpace(rows[i]))
                        {
                            continue;
                        }
                        string cleanedRow = rows[i].TrimStart(';').Trim();
                        string[] values = cleanedRow.Split(',');

                        // Display each ErrorRowNumber on a separate row
                        for (int j = 0; j < values.Length; j++)
                        {
                            columnNamesWorksheet.Range[rowIndex, j + 1].Text = values[j];
                        }

                        columnNamesWorksheet.Range[rowIndex, values.Length + 1].Text = errorMessage;
                        columnNamesWorksheet.Range[rowIndex, values.Length + 2].Text = errorRowNumbers[i - 1]; // Use (i - 1) to get the corresponding ErrorRowNumber
                        rowIndex++;
                    }

                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        //Get Tbale name by Entity ID
        public int GetEntityIdByEntityName(string entityName)
        {
            // Assuming you have a list of EntityListMetadataModel instances
            var entityListMetadataModels = GetEntityListMetadataModels(); // Implement this method to fetch your metadata models

            // Use LINQ to find the entity Id
            int entityId = entityListMetadataModels
                .Where(model => model.EntityName == entityName)
                .Select(model => model.Id)
                .FirstOrDefault();

            if (entityId != 0) // Check if a valid entity Id was found
            {
                return entityId;
            }
            else
            {
                // Handle the case where the entity name is not found
                throw new Exception("Entity not found");
            }
        }

        // Get Entity names from meta table
        private List<TableMetaDataEntity> GetEntityListMetadataModels()
        {
            var sample = _context.TableMetaDataEntity.ToList();

            return sample;
        }

        // Read excel file
        public DataTable ReadExcelFromFormFile(IFormFile excelFile)
        {
            using (Stream stream = excelFile.OpenReadStream())
            {
                using (var package = new ExcelPackage(stream))
                {
                    DataTable dataTable = new DataTable();

                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {

                        var firstCell = worksheet.Cells[2, col];

                        if (string.IsNullOrWhiteSpace(firstCell.Text))
                        {
                            continue;
                        }
                        dataTable.Columns.Add(firstCell.Text);
                    }


                    dataTable.Columns.Add("RowNumber", typeof(int)); // Add "RowNumber" column


                    for (int rowNumber = 3; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                    {
                        var dataRow = dataTable.NewRow();

                        // Set the "RowNumber" value for each row

                        int colIndex = 0;

                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {

                            // Check if this column should be included

                            if (dataTable.Columns.Contains(worksheet.Cells[2, col].Text))
                            {

                                dataRow[colIndex] = worksheet.Cells[rowNumber, col].Text;

                                colIndex++;

                            }
                        }

                        dataTable.Rows.Add(dataRow);
                    }

                    bool allRowsAreNull = dataTable.AsEnumerable()

                    .All(row => row.ItemArray.All(field => field is DBNull || string.IsNullOrWhiteSpace(field.ToString())));

                    if (allRowsAreNull)
                    {
                        return null;
                    }

                    dataTable = dataTable.AsEnumerable()

                        .Where(row => !row.ItemArray.All(field => field is DBNull || string.IsNullOrWhiteSpace(field.ToString())))

                        .CopyToDataTable();

                    dataTable = dataTable.AsEnumerable().Select((row, index) =>
                    {
                        row.SetField("RowNumber", index + 3);
                        return row;
                    }).CopyToDataTable();
                    return dataTable;
                }
            }
        }

        // Read excel file
        public List<Dictionary<string, string>> ReadDataFromExcel(Stream excelFileStream, int rowCount)
        {
            using (var package = new ExcelPackage(excelFileStream))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                rowCount = rowCount + 2;

                int colCount = worksheet.Dimension.Columns;

                var data = new List<Dictionary<string, string>>();

                var columnNames = new List<string>();

                var skipColumns = new List<bool>();

                for (int col = 1; col <= colCount; col++)
                {
                    var columnName = worksheet.Cells[2, col].Value?.ToString();

                    columnNames.Add(columnName);

                    skipColumns.Add(string.IsNullOrWhiteSpace(columnName));
                }

                // Read data rows

                for (int row = 3; row <= rowCount; row++)
                {
                    var rowData = new Dictionary<string, string>();

                    for (int col = 1; col <= colCount; col++)
                    {
                        if (!skipColumns[col - 1])
                        {

                            var columnName = columnNames[col - 1];

                            var cellValue = worksheet.Cells[row, col].Value?.ToString();

                            rowData[columnName] = cellValue;
                        }
                    }

                    // Include the row number as "RowNumber" in the dictionary

                    rowData["RowNumber"] = row.ToString();

                    data.Add(rowData);
                }
                return data;
            }
        }

        public bool IsValidByteA(string data, int minLength, int maxLength)
        {
            // Check if the input is a valid hexadecimal string
            if (!IsHexString(data))
            {
                return false;
            }

            // Check if the length is within acceptable limits
            if (data.Length < minLength || data.Length > maxLength)
            {
                return false;
            }
            // Add more specific checks if needed
            return true;
        }

        public bool IsHexString(string input)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(input, @"\A\b[0-9a-fA-F]+\b\Z");
        }

        public IEnumerable<ColumnMetaDataDTO> GetColumnsForEntity(string entityName)
        {
            var entity = _context.TableMetaDataEntity.FirstOrDefault(e => e.EntityName == entityName);
            if (entity == null)
            {
                // Entity not found, return a 404 Not Found response
                return null;
            }
            var columnsDTO = _context.ColumnMetaDataEntity
                .Where(column => column.EntityId == entity.Id)
                .Select(column => new ColumnMetaDataDTO
                {
                    Id = column.Id,
                    ColumnName = column.ColumnName,
                    Datatype = column.Datatype,
                    Length = column.Length,
                    Description = column.Description,
                    IsNullable = column.IsNullable,
                    DefaultValue = column.DefaultValue,
                    IsPrimaryKey = column.IsPrimaryKey,
                    True = column.True,
                    False = column.False,
                    MinLength = column.MinLength,
                    MaxLength = column.MaxLength,
                    MinRange = column.MinRange,
                    MaxRange = column.MaxRange,
                    DateMinValue = column.DateMinValue,
                    DateMaxValue = column.DateMaxValue,
                    IsForeignKey = column.IsForeignKey,
                    ReferenceColumnID = column.ReferenceColumnID,
                    ReferenceEntityID = column.ReferenceEntityID,
                }).ToList();
            if (columnsDTO.Count == 0)
            {
                // No columns found, return a 404 Not Found response with an error message
                return null;
            }
            return columnsDTO;
        }

        //not Nul Validation
        public async Task<ValidationResultData> ValidateNotNull(DataTable excelData, List<ColumnMetaDataDTO> columnsDTO)
        {
            List<string> badRows = new List<string>();
            List<string> errorColumnNames = new List<string>();
            DataTable validRowsDataTable = excelData.Clone(); // Create a DataTable to store valid rows

            for (int row = 0; row < excelData.Rows.Count; row++)
            {
                bool rowValidationFailed = false;

                string badRow = string.Join(",", excelData.Rows[row].ItemArray);

                if (excelData.Columns.Contains("ErrorMessage"))
                {
                    for (int col = 0; col < excelData.Columns.Count - 3; col++)
                    {
                        string cellData = excelData.Rows[row][col].ToString();
                        ColumnMetaDataDTO columnDTO = columnsDTO[col];

                        if (columnDTO.IsNullable == true && string.IsNullOrEmpty(cellData))
                        {
                            rowValidationFailed = true;
                            badRows.Add(badRow);
                            if (!errorColumnNames.Contains(columnDTO.ColumnName))
                            {
                                errorColumnNames.Add(columnDTO.ColumnName);
                            }

                            break;
                        }
                    }
                }
                else
                {
                    for (int col = 0; col < excelData.Columns.Count - 2; col++)
                    {
                        string cellData = excelData.Rows[row][col].ToString();
                        ColumnMetaDataDTO columnDTO = columnsDTO[col];

                        if(columnDTO.IsNullable == false && string.IsNullOrEmpty(cellData))
                        //if (columnDTO.IsNullable == true && string.IsNullOrEmpty(cellData))
                        {
                            rowValidationFailed = true;
                            badRows.Add(badRow);
                            if (!errorColumnNames.Contains(columnDTO.ColumnName))
                            {
                                errorColumnNames.Add(columnDTO.ColumnName);
                            }

                            break;
                        }
                    }

                }

                if (!rowValidationFailed)
                {
                    validRowsDataTable.Rows.Add(excelData.Rows[row].ItemArray);
                }
            }

            // Return both results
            return new ValidationResultData { BadRows = badRows, SuccessData = validRowsDataTable, errorcolumns = errorColumnNames, Column_Name = string.Empty };
        }

        //Primary Key validation
        public async Task<ValidationResultData> ValidatePrimaryKeyAsync(ValidationResultData validationResult, List<ColumnMetaDataDTO> columnsDTO, string tableName,DBConnectionDTO connectionDTO)
        {
            List<string> badRows = new List<string>();
            string columnName = string.Empty;
            DataTable validRowsDataTable = validationResult.SuccessData;
            DataTable successdata = validRowsDataTable.Clone();
            List<int> primaryKeyColumns = new List<int>();
            if (validRowsDataTable.Columns.Contains("ErrorMessage"))
            {
                for (int col = 0; col < validRowsDataTable.Columns.Count - 3; col++)
                {
                    ColumnMetaDataDTO columnDTO = columnsDTO[col];
                    if (columnDTO.IsPrimaryKey)
                    {
                        primaryKeyColumns.Add(col);
                        columnName = columnDTO.ColumnName; // Set the primary key column name
                    }
                }
            }
            else
            {
                for (int col = 0; col < validRowsDataTable.Columns.Count - 2; col++)
                {
                    ColumnMetaDataDTO columnDTO = columnsDTO[col];
                    if (columnDTO.IsPrimaryKey)
                    {
                        primaryKeyColumns.Add(col);
                        columnName = columnDTO.ColumnName; // Set the primary key column name
                    }
                }
            }

            if (!String.IsNullOrEmpty(columnName))
            {
                HashSet<string> seenValues = new HashSet<string>();
                var ids = await GetAllIdsFromDynamicTable(connectionDTO, tableName);
                List<string> idsAsStringList = ids.Select(item => item.ToString()).Cast<string>().ToList();

                for (int row = 0; row < validRowsDataTable.Rows.Count; row++)
                {
                    bool rowValidationFailed = false;
                    //  var badRowData = new List<string>();

                    for (int col = 0; col < primaryKeyColumns.Count; col++)
                    {
                        int primaryKeyColumnIndex = primaryKeyColumns[col];
                        string cellData = validRowsDataTable.Rows[row][primaryKeyColumnIndex].ToString();

                        if (seenValues.Contains(cellData))
                        {
                            // Set the flag to indicate validation failure for this row
                            rowValidationFailed = true;
                            columnName = columnsDTO[primaryKeyColumnIndex].ColumnName;
                            badRows.Add(string.Join(",", validRowsDataTable.Rows[row].ItemArray)); // Store the row data
                            break; // Exit the loop as soon as a validation failure is encountered
                        }

                        if (idsAsStringList.Contains(cellData))
                        {
                            rowValidationFailed = true;
                            columnName = columnsDTO[primaryKeyColumnIndex].ColumnName;
                            badRows.Add(string.Join(",", validRowsDataTable.Rows[row].ItemArray)); // Store the row data
                            break;
                        }

                        // Store the value for duplicate checking
                        seenValues.Add(cellData);
                    }

                    // If row validation succeeded, add the entire row to the successdata DataTable
                    if (!rowValidationFailed)
                    {
                        successdata.Rows.Add(validRowsDataTable.Rows[row].ItemArray);
                    }
                }

            }
            // Return both bad rows and success data using the custom class
            return new ValidationResultData { BadRows = badRows, SuccessData = successdata, Column_Name = columnName };
        }

        //Convert result type
        public async Task<ValidationResult> resultparams(ValidationResultData validationResult, string comma_separated_string)
        {
            string errorMessages = string.Empty;
            string values = string.Join(",", validationResult.BadRows.Select(row => row.Split(',').Last()));
            validationResult.BadRows.Insert(0, comma_separated_string);
            List<string> modifiedRows = validationResult.BadRows.Select(row => row.Substring(0, row.LastIndexOf(','))).ToList();
            validationResult.BadRows = modifiedRows;
            string delimiter = ";"; // Specify the delimiter you want
            string delimiter1 = ","; // Specify the delimiter you want   //chng
            string baddatas = string.Join(delimiter, validationResult.BadRows);
            string badcolumns = string.Join(delimiter1, validationResult.errorcolumns);
            errorMessages = "Null value found in column" + " " + badcolumns;
            // Return both results
            return new ValidationResult { ErrorRowNumber = values, Filedatas = baddatas, errorMessages = errorMessages };
        }

        //Convert primary key validation result to result type
        public async Task<ValidationResult> resultparamsforprimary(ValidationResultData validationResult, string comma_separated_string, string tableName)
        {

            var badRowsPrimaryKey = validationResult.BadRows;


            string columnName = validationResult.Column_Name;

            badRowsPrimaryKey = badRowsPrimaryKey.Where(x => x != "").ToList();
            string values = string.Join(",", badRowsPrimaryKey.Select(row => row.Split(',').Last()));

            badRowsPrimaryKey.Insert(0, comma_separated_string);

            List<string> modifiedRows = badRowsPrimaryKey.Select(row =>
            {
                int lastCommaIndex = row.LastIndexOf(',');
                if (lastCommaIndex >= 0)
                {
                    return row.Substring(0, lastCommaIndex);
                }
                else
                {
                    return row; // No comma found, keep the original string
                }
            }).Where(row => !string.IsNullOrEmpty(row)).ToList();
            badRowsPrimaryKey = modifiedRows;
            string delimiter = ";"; // Specify the delimiter you want
            string baddatas = string.Join(delimiter, badRowsPrimaryKey);
            string errorMessages = "Duplicate key value violates unique constraints in column " + columnName + " " + "in" + " " + tableName;

            // Return both results
            return new ValidationResult { ErrorRowNumber = values, Filedatas = baddatas, errorMessages = errorMessages };
        }

        //Convert range validation result to result type
        public async Task<ValidationResult> resultparamsforrange(ValidationResultData validationResult, string comma_separated_string, string tableName)
        {
            var badRowsPrimaryKey = validationResult.BadRows;

            string columnName = validationResult.Column_Name;

            badRowsPrimaryKey = badRowsPrimaryKey.Where(x => x != "").ToList();
            string values = string.Join(",", badRowsPrimaryKey.Select(row => row.Split(',').Last()));

            badRowsPrimaryKey.Insert(0, comma_separated_string);

            List<string> modifiedRows = badRowsPrimaryKey.Select(row =>
            {
                int lastCommaIndex = row.LastIndexOf(',');
                if (lastCommaIndex >= 0)
                {
                    return row.Substring(0, lastCommaIndex);
                }
                else
                {
                    return row; // No comma found, keep the original string
                }
            }).Where(row => !string.IsNullOrEmpty(row)).ToList();
            badRowsPrimaryKey = modifiedRows;
            string delimiter = ";"; // Specify the delimiter you want
            string baddatas = string.Join(delimiter, badRowsPrimaryKey);
            string errorMessages = "Incorrect Range Value on " + columnName + " " + "in" + " " + tableName;

            // Return both results
            return new ValidationResult { ErrorRowNumber = values, Filedatas = baddatas, errorMessages = errorMessages };
        }

        // Insert Log
        public async Task<LogDTO> Createlog(string tableName, List<string> filedata, string fileName, int successdata, List<string> errorMessage, int total_count, List<string> ErrorRowNumber)
        {
            var storeentity = await _context.TableMetaDataEntity.FirstOrDefaultAsync(x => x.EntityName.ToLower() == tableName.ToLower());

            LogParent logParent = new LogParent();

            logParent.FileName = fileName;

            logParent.User_Id = 1;

            logParent.Entity_Id = storeentity.Id;

            logParent.Timestamp = DateTime.UtcNow;

            logParent.PassCount = successdata;

            logParent.RecordCount = total_count;

            logParent.FailCount = total_count - successdata;

            // Insert the LogParent record

            _context.LogParents.Add(logParent);

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

            List<LogChild> logChildren = new List<LogChild>();

            for (int i = 0; i < errorMessage.Count; i++)
            {
                LogChild logChild = new LogChild();

                logChild.ParentID = logParent.ID; // Set the ParentId

                logChild.ErrorMessage = errorMessage[i];

                if (filedata.Count > 0)
                {
                    logChild.Filedata = filedata[i];
                }
                else
                {
                    logChild.Filedata = ""; // Set the filedata as needed
                }

                if (ErrorRowNumber.Count > 0)
                {
                    logChild.ErrorRowNumber = ErrorRowNumber[i];
                }
                else
                {
                    logChild.ErrorRowNumber = ""; // Set the filedata as needed
                }
                if (ErrorRowNumber.Count > 0)
                {
                    logChild.ErrorRowNumber = ErrorRowNumber[i];
                }
                else
                {
                    logChild.ErrorRowNumber = ""; // Set the filedata as needed
                }

                // Insert the LogChild record
                _context.LogChilds.Add(logChild);

                logChildren.Add(logChild);
            }
            try
            {
                await _context.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                // Log or handle the exception
                Console.WriteLine("Error: " + ex.Message);
            }
            LogDTO logDTO = new LogDTO()
            {
                LogParentDTOs = logParent,
                ChildrenDTOs = logChildren
            };
            return logDTO;
        }

        //Insert uploaded data into Client db
        public async Task InsertDataFromDataTableToPostgreSQL(DataTable data, string tableName, List<string> columns, IFormFile file, DBConnectionDTO connectionDTO)
        {
            try
            {
                var columnProperties = GetColumnsForEntity(tableName).ToList();
                List<ColumnMetaDataDTO> booleanColumns = columnProperties.Where(c => c.Datatype.ToLower() == "boolean" || c.Datatype.ToLower() == "bit").ToList();
                List<ColumnMetaDataDTO> jsonColumns = columnProperties.Where(c => c.Datatype.ToLower() == "jsonb").ToList();
                List<Dictionary<string, string>> convertedDataList = new List<Dictionary<string, string>>();

                foreach (DataRow row in data.Rows)
                {
                    Dictionary<string, string> convertedData = new Dictionary<string, string>();

                    foreach (DataColumn column in data.Columns)
                    {
                        string columnName = column.ColumnName;
                        string cellValue = row[column].ToString();

                        ColumnMetaDataDTO columnProperty = columnProperties.FirstOrDefault(col => col.ColumnName == columnName);

                        if (columnProperty != null)
                        {
                            if (jsonColumns.Any(jsonCol => jsonCol.ColumnName == columnName) && !IsJson(cellValue))
                            {
                                // Convert the value to JSON format for JSONB columns
                                var jsonObject = new { JsonColumnName = cellValue };
                                convertedData[columnName] = JsonConvert.SerializeObject(jsonObject);
                            }
                            else
                            {
                                // Use the column name from ColumnProperties as the key and the cell value as the value
                                convertedData[columnName] = cellValue;
                            }
                        }
                    }

                    convertedDataList.Add(convertedData);
                }

                // 'convertedDataList' is now a list of dictionaries, each representing a row in the desired format.
                var storeEntity = await _context.TableMetaDataEntity.FirstOrDefaultAsync(x => x.EntityName.ToLower() == tableName.ToLower());

                tableName = storeEntity.EntityName;

                string encodedPassword = Uri.EscapeDataString(connectionDTO.Password);

                // Convert the parameters to strings
                string connectionDTOString = JsonConvert.SerializeObject(new DBConnectionDTO
                {
                    Provider = connectionDTO.Provider,
                    HostName = connectionDTO.HostName,
                    DataBase = connectionDTO.DataBase,
                    UserName = connectionDTO.UserName,
                    Password = encodedPassword 
                });

                System.Diagnostics.Debug.WriteLine($"Serialized DTO: {connectionDTOString}");

                // Convert the parameters to strings
                string convertedDataListString = JsonConvert.SerializeObject(convertedDataList);
                string booleanColumnsString = JsonConvert.SerializeObject(booleanColumns);

                // Deserialize the JSON string back into a list of dictionaries
                List<Dictionary<string, string>> deserializedDataList = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(convertedDataListString);


                // Create an HttpClient instance
                using (var httpClient = new HttpClient())
                {
                    // Set the base address for the external API
                    httpClient.BaseAddress = new Uri("https://localhost:7246");

                    
                    // Create query parameters
                    var queryString = new Dictionary<string, string>
            {
                { "ConnectionDTO", connectionDTOString },
                { "ConvertedDataList", convertedDataListString },
                { "BooleanColumns", booleanColumnsString },
                { "TableName", tableName }
            };

                    // Call the external API with the query parameters
                    var response = await httpClient.PostAsync($"/EntityMigrate/InsertData?{ToQueryString(queryString)}", null);

                    // Check the response status
                    if (response.IsSuccessStatusCode)
                    {
                        // Process successful response
                        var result = await response.Content.ReadAsStringAsync();
                        // Handle the result as needed
                    }
                    else
                    {
                        // Handle error response
                        // You can inspect the response status code and content for more information
                        var errorMessage = await response.Content.ReadAsStringAsync();
                        // Handle the error message
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle the exception, log, or rethrow as needed
                throw new Exception($"Error in InsertDataFromDataTableToPostgreSQL: {ex.Message}");
            }
        }


        private bool IsJson(string str)
        {
            try
            {
                JsonConvert.DeserializeObject(str);
                return true;
            }
            catch
            {
                return false;
            }
        }

        // Helper method to convert dictionary to query string
        private string ToQueryString(Dictionary<string, string> parameters)
        {
            return string.Join("&", parameters.Select(x => $"{WebUtility.UrlEncode(x.Key)}={WebUtility.UrlEncode(x.Value)}"));
        }

        public int GetEntityIdByEntityNamefromui(string entityName)
        {
            // Assuming you have a list of EntityListMetadataModel instances
            List<TableMetaDataEntity> entityListMetadataModels = GetEntityListMetadataModelforlist(); // Implement this method to fetch your metadata models

            // Use LINQ to find the entity Id
            int entityId = entityListMetadataModels
                .Where(model => model.EntityName == entityName)
                .Select(model => model.Id)
                .FirstOrDefault();

            if (entityId != 0) // Check if a valid entity Id was found
            {
                return entityId;
            }
            else
            {
                // Handle the case where the entity name is not found
                throw new Exception("Entity not found");//return null
            }
        }

        public List<TableMetaDataEntity> GetEntityListMetadataModelforlist()
        {
            {
                List<TableMetaDataEntity> entityListMetadataModels = _context.TableMetaDataEntity.ToList();
                return entityListMetadataModels;
            }
        }

        public int? GetEntityIdFromTemplate(IFormFile file, int sheetIndex)
        {
            using (var package = new ExcelPackage(file.OpenReadStream()))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetIndex];
                int entityId;
                if (int.TryParse(worksheet.Cells[1, 1].Text, out entityId))
                {
                    return entityId;
                }
                return null;
            }
        }
        public string GetPrimaryKeyColumnForEntity(string entityName)
        {
            var entity = _context.TableMetaDataEntity.FirstOrDefault(e => e.EntityName == entityName);
            if (entity == null)
            {
                // Entity not found, return null or throw an exception
                return null;
            }
            var primaryKeyColumn = _context.ColumnMetaDataEntity
                .Where(column => column.EntityId == entity.Id && column.IsPrimaryKey)
                .Select(column => column.ColumnName)
                .FirstOrDefault();
            return primaryKeyColumn;
        }

        public async Task<List<dynamic>> GetAllIdsFromDynamicTable(DBConnectionDTO connectionDTO, string tableName)
        {
            try
            {
                // Assuming you have a valid base URL for the other project
                string otherApiBaseUrl = "https://localhost:7246/";

                // Create HttpClient instance
                using (var httpClient = new HttpClient())
                {
                    // Set the base address of the other API
                    httpClient.BaseAddress = new Uri(otherApiBaseUrl);

                    string encodedPassword = Uri.EscapeDataString(connectionDTO.Password);
                   string encryptedHostName = Uri.EscapeDataString(connectionDTO.HostName);


                    // Call the other API to get table details
                    var response = await httpClient.GetAsync($"EntityMigrate/GetPrimaryColumnData?Provider={connectionDTO.Provider}&HostName={connectionDTO.HostName}&DataBase={connectionDTO.DataBase}&UserName={connectionDTO.UserName}&Password={encodedPassword}&tableName={tableName}");

                    // Check the response status
                    if (response.IsSuccessStatusCode)
                    {
                        var contentAsString = await response.Content.ReadAsStringAsync();
                        var apiResponse = JsonConvert.DeserializeObject<Models.DTO.APIResponse>(contentAsString);

                        if (apiResponse.IsSuccess)
                        {
                            // Assuming Result is a List<dynamic>
                            var primaryColumnData = JsonConvert.DeserializeObject<List<dynamic>>(Convert.ToString(apiResponse.Result));
                            return primaryColumnData;
                        }
                        else
                        {
                            return null;
                        }
                    }
                    else
                    {
                        // Handle the error, log, or throw a new exception as needed
                        var errorMessage = await response.Content.ReadAsStringAsync();
                        throw new Exception($"Error calling the GetPrimaryColumnData API. Details: {errorMessage}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Log the exception or handle it appropriately
                throw new Exception($"Error calling the GetPrimaryColumnData API. Details: {ex.Message}");
            }
        }

        public bool TableExists(string tableName)
        {
            try
            {
                //Implement IsTableExists APi
                return true;
            }
            catch (Exception ex)
            {
                // Log or print detailed information about the exception
                Console.WriteLine($"Error checking table existence. ConnectionString: TableName: {tableName}, Exception: {ex}");

                // Rethrow the exception
                throw;
            }
        }

        public (int EntityId, string EntityColumnName) GetAllEntityColumnData(int checklistEntityValue)
        {
            List<ColumnMetaDataEntity> allData = _context.ColumnMetaDataEntity
                .AsEnumerable()  // Bring data into memory
                .Where(c => c.Id == checklistEntityValue)
                .ToList();

            if (allData.Count > 0)
            {
                // Return the EntityId and EntityColumnName of the first item in the list
                return (allData[0].EntityId, allData[0].ColumnName);
            }
            else
            {
                // Return default values or handle as needed
                return (0, string.Empty);
            }
        }

        public List<dynamic> GetTableDataByChecklistEntityValue(DBConnectionDTO connectionDTO, int checklistEntityValue)
        {
            // Get the EntityId and EntityColumnName using the provided method
            var (entityId, entityColumnName) = GetAllEntityColumnData(checklistEntityValue);

            if (entityId == 0 || string.IsNullOrEmpty(entityColumnName))
            {
                // Handle the case where the EntityId or EntityColumnName is not found
                throw new InvalidOperationException("No entity metadata found");
            }

            // Use Entity Framework Core to get the table name
            var entityListMetadata = _context.TableMetaDataEntity.FirstOrDefault(entity => entity.Id == entityId);

            if (entityListMetadata == null)
            {
                // Log or print debug information
                throw new InvalidOperationException("No entity metadata found");
            }

            string tableName = entityListMetadata.EntityName;

            try
            {
                // Create an HttpClient instance
                using (var httpClient = new HttpClient())
                {
                    // Set the base address for the API
                    httpClient.BaseAddress = new Uri("https://localhost:7246");

                    // Call the GetTabledata API endpoint
                    var response = httpClient.GetAsync($"/EntityMigrate/GetPrimaryColumnData?Provider={connectionDTO.Provider}&HostName={connectionDTO.HostName}&DataBase={connectionDTO.DataBase}&UserName={connectionDTO.UserName}&Password={connectionDTO.Password}&tableName={tableName}").Result;

                    // Check the response status
                    if (response.IsSuccessStatusCode)
                    {
                        // Read the response content as a string
                        var contentAsString = response.Content.ReadAsStringAsync().Result;

                        // Deserialize the string content to a list of dynamic objects
                        var tableData = JsonConvert.DeserializeObject<List<dynamic>>(contentAsString);

                        // Continue with the rest of your method logic...
                        return tableData;
                    }
                    else
                    {
                        // Handle the error, log, or throw a new exception as needed
                        var errorMessage = response.Content.ReadAsStringAsync().Result;
                        throw new Exception($"Error fetching table data. Details: {errorMessage}");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error fetching table data: {ex.Message}");
            }
        }

        // range Validation
        public async Task<ValidationResultData> ValidateRange(ValidationResultData validationResult, List<ColumnMetaDataDTO> columnsDTO, string tableName)
        {
            List<string> badRows = new List<string>();
            List<string> errorColumnNames = new List<string>();
            var excelData = validationResult.SuccessData;
            DataTable validRowsDataTable = excelData.Clone(); // Create a DataTable to store valid rows
            for (int row = 0; row < excelData.Rows.Count; row++)
            {
                bool rowValidationFailed = false;

                string badRow = string.Join(",", excelData.Rows[row].ItemArray);

                if (excelData.Columns.Contains("ErrorMessage"))
                {
                    for (int col = 0; col < excelData.Columns.Count - 3; col++)
                    {
                        string cellData = excelData.Rows[row][col].ToString();
                        ColumnMetaDataDTO columnDTO = columnsDTO[col];

                        if (double.TryParse(cellData, out double numericValue))
                        {
                            if (numericValue < columnDTO.MinRange)
                            {
                                // Your logic for when numericValue is less than MinLength
                                rowValidationFailed = true;
                                badRows.Add(badRow);
                                if (!errorColumnNames.Contains(columnDTO.ColumnName))
                                {
                                    errorColumnNames.Add(columnDTO.ColumnName);
                                }
                                break;
                            }

                            if (numericValue > columnDTO.MaxRange)
                            {
                                // Your logic for when numericValue is greater than MaxLength
                                rowValidationFailed = true;
                                badRows.Add(badRow);
                                if (!errorColumnNames.Contains(columnDTO.ColumnName))
                                {
                                    errorColumnNames.Add(columnDTO.ColumnName);
                                }
                                break;
                            }
                        }
                    }
                }
                else
                {
                    for (int col = 0; col < excelData.Columns.Count - 2; col++)
                    {
                        string cellData = excelData.Rows[row][col].ToString();
                        ColumnMetaDataDTO columnDTO = columnsDTO[col];

                        if (double.TryParse(cellData, out double numericValue))
                        {
                            if (numericValue < columnDTO.MinRange)
                            {
                                // Your logic for when numericValue is less than MinLength
                                rowValidationFailed = true;
                                badRows.Add(badRow);
                                if (!errorColumnNames.Contains(columnDTO.ColumnName))
                                {
                                    errorColumnNames.Add(columnDTO.ColumnName);
                                }
                                break;
                            }

                            if (numericValue > columnDTO.MaxRange)
                            {
                                // Your logic for when numericValue is greater than MaxLength
                                rowValidationFailed = true;
                                badRows.Add(badRow);
                                if (!errorColumnNames.Contains(columnDTO.ColumnName))
                                {
                                    errorColumnNames.Add(columnDTO.ColumnName);
                                }
                                break;
                            }
                        }
                    }
                }

                if (!rowValidationFailed)
                {
                    validRowsDataTable.Rows.Add(excelData.Rows[row].ItemArray);
                }
            }
            // Return both results
            return new ValidationResultData { BadRows = badRows, SuccessData = validRowsDataTable, errorcolumns = errorColumnNames, Column_Name = string.Empty };
        }

        //Length validation
        public async Task<ValidationResultData> ValidateLength(ValidationResultData validationResult, List<ColumnMetaDataDTO> columnsDTO, string tableName)
        {
            List<string> badRows = new List<string>();
            List<string> errorColumnNames = new List<string>();
            var excelData = validationResult.SuccessData;
            DataTable validRowsDataTable = excelData.Clone(); // Create a DataTable to store valid rows
            for (int row = 0; row < excelData.Rows.Count; row++)
            {
                bool rowValidationFailed = false;

                string badRow = string.Join(",", excelData.Rows[row].ItemArray);

                if (excelData.Columns.Contains("ErrorMessage"))
                {
                    for (int col = 0; col < excelData.Columns.Count - 3; col++)
                    {
                        string cellData = excelData.Rows[row][col].ToString();
                        ColumnMetaDataDTO columnDTO = columnsDTO[col];

                            if (columnDTO.MinLength > cellData.Length)
                            {
                                // Your logic for when numericValue is less than MinLength
                                rowValidationFailed = true;
                                badRows.Add(badRow);
                                if (!errorColumnNames.Contains(columnDTO.ColumnName))
                                {
                                    errorColumnNames.Add(columnDTO.ColumnName);
                                }
                                break;
                            }

                            if (columnDTO.MaxLength < cellData.Length)
                            {
                                // Your logic for when numericValue is greater than MaxLength
                                rowValidationFailed = true;
                                badRows.Add(badRow);
                                if (!errorColumnNames.Contains(columnDTO.ColumnName))
                                {
                                    errorColumnNames.Add(columnDTO.ColumnName);
                                }
                                break;
                            }
                        }
                }
                else
                {
                    for (int col = 0; col < excelData.Columns.Count - 2; col++)
                    {
                        string cellData = excelData.Rows[row][col].ToString();
                        ColumnMetaDataDTO columnDTO = columnsDTO[col];

                            if (columnDTO.MinLength > cellData.Length)
                            {
                                // Your logic for when numericValue is less than MinLength
                                rowValidationFailed = true;
                                badRows.Add(badRow);
                                if (!errorColumnNames.Contains(columnDTO.ColumnName))
                                {
                                    errorColumnNames.Add(columnDTO.ColumnName);
                                }
                                break;
                            }

                            if (columnDTO.MaxLength < cellData.Length)
                            {
                                // Your logic for when numericValue is greater than MaxLength
                                rowValidationFailed = true;
                                badRows.Add(badRow);
                                if (!errorColumnNames.Contains(columnDTO.ColumnName))
                                {
                                    errorColumnNames.Add(columnDTO.ColumnName);
                                }
                                break;
                            }
                    }
                }

                if (!rowValidationFailed)
                {
                    validRowsDataTable.Rows.Add(excelData.Rows[row].ItemArray);
                }
            }
            // Return both results
            return new ValidationResultData { BadRows = badRows, SuccessData = validRowsDataTable, errorcolumns = errorColumnNames, Column_Name = string.Empty };
        }

        //Convert range validation result to result type
        public async Task<ValidationResult> resultparamsforlength(ValidationResultData validationResult, string comma_separated_string, string tableName)
        {
            var badRowsPrimaryKey = validationResult.BadRows;

            string columnName = validationResult.Column_Name;

            badRowsPrimaryKey = badRowsPrimaryKey.Where(x => x != "").ToList();
            string values = string.Join(",", badRowsPrimaryKey.Select(row => row.Split(',').Last()));

            badRowsPrimaryKey.Insert(0, comma_separated_string);

            List<string> modifiedRows = badRowsPrimaryKey.Select(row =>
            {
                int lastCommaIndex = row.LastIndexOf(',');
                if (lastCommaIndex >= 0)
                {
                    return row.Substring(0, lastCommaIndex);
                }
                else
                {
                    return row; // No comma found, keep the original string
                }
            }).Where(row => !string.IsNullOrEmpty(row)).ToList();
            badRowsPrimaryKey = modifiedRows;
            string delimiter = ";"; // Specify the delimiter you want
            string baddatas = string.Join(delimiter, badRowsPrimaryKey);
            string errorMessages = "Incorrect Length Value on " + columnName + " " + "in" + " " + tableName;

            // Return both results
            return new ValidationResult { ErrorRowNumber = values, Filedatas = baddatas, errorMessages = errorMessages };
        }
        public async Task<List<LogChild>> GetAllLogChildsByParentIDAsync(int parentID)
        {
            try
            {
                return await _context.LogChilds
                    .Where(c => c.ParentID == parentID)
                    .ToListAsync();
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}

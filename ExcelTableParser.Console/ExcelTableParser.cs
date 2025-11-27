using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

namespace ExcelTableParser.Console
{
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.Linq;
    using System.Reflection;

    // -------------------------------------------------
    // NEW: Column Attribute
    // -------------------------------------------------

    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnAttribute : Attribute
    {
        public string Name { get; }
        public ColumnAttribute(string name) => Name = name;
    }

    // -------------------------------------------------
    // Error + Result Models
    // -------------------------------------------------

    public class ExcelTableResult<T>
    {
        public List<T> Items { get; set; } = new();
        public List<ExcelTableError> Errors { get; set; } = new();
    }

    public class ExcelTableError
    {
        public int RowNumber { get; set; }
        public string ColumnName { get; set; }
        public string PropertyName { get; set; }
        public string RawValue { get; set; }
        public string Message { get; set; }
    }

    // -------------------------------------------------
    // Main Parser
    // -------------------------------------------------

public static class ExcelTableParser
    {
        public static ExcelTableResult<T> ParseTable<T>(
            Stream excelStream,
            string sheetName,
            string tableName = null,
            IServiceProvider serviceProvider = null,
            Func<T, int, IEnumerable<string>> customValidator = null)
            where T : new()
        {
            var result = new ExcelTableResult<T>();

            using var doc = SpreadsheetDocument.Open(excelStream, false);
            var wbPart = doc.WorkbookPart;

            // --------------------- Locate Sheet ---------------------
            var sheet = wbPart.Workbook.Descendants<Sheet>()
                        .FirstOrDefault(s => s.Name == sheetName)
                        ?? throw new Exception($"Sheet '{sheetName}' not found.");

            var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);

            // --------------------- Locate Table ---------------------
            var tablePart = wsPart.TableDefinitionParts
                .FirstOrDefault(tp => tableName == null || tp.Table.DisplayName == tableName)
                ?? throw new Exception($"Table '{tableName}' not found.");

            var table = tablePart.Table;
            string range = table.Reference;

            ParseRange(range, out int colStart, out int colEnd, out int rowStart, out int rowEnd);

            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();

            // --------------------- Read Headers ---------------------
            Row headerRow = sheetData.Elements<Row>()
                                     .First(r => r.RowIndex == (uint)rowStart);

            List<string> headers = new();
            for (int c = colStart; c <= colEnd; c++)
            {
                string cellRef = GetColName(c) + rowStart;
                Cell cell = headerRow.Elements<Cell>().FirstOrDefault(x => x.CellReference == cellRef);
                headers.Add(ReadCellValue(wbPart, cell).Trim());
            }

            // --------------------- Attribute-Based Column Mapping ---------------------
            var props = typeof(T).GetProperties();
            var columnMap = new Dictionary<string, PropertyInfo>(StringComparer.OrdinalIgnoreCase);

            foreach (var prop in props)
            {
                var attr = prop.GetCustomAttribute<ColumnAttribute>();
                string colName = attr?.Name ?? prop.Name;
                columnMap[colName] = prop;
            }

            var columnToProp = new PropertyInfo[headers.Count];

            for (int i = 0; i < headers.Count; i++)
            {
                if (columnMap.TryGetValue(headers[i], out var prop))
                    columnToProp[i] = prop;
            }

            // --------------------- Parse Data Rows ---------------------
            for (int r = rowStart + 1; r <= rowEnd; r++)
            {
                Row row = sheetData.Elements<Row>().FirstOrDefault(x => x.RowIndex == (uint)r);
                if (row == null) continue;

                T obj = new();

                for (int c = colStart; c <= colEnd; c++)
                {
                    int idx = c - colStart;
                    var prop = columnToProp[idx];
                    if (prop == null) continue;

                    string columnHeader = headers[idx];

                    string cellRef = GetColName(c) + r;
                    Cell cell = row.Elements<Cell>().FirstOrDefault(x => x.CellReference == cellRef);
                    string rawValue = ReadCellValue(wbPart, cell);

                    object converted = ConvertSafe(rawValue, prop.PropertyType, out string error);

                    if (error != null)
                    {
                        result.Errors.Add(new ExcelTableError
                        {
                            RowNumber = r,
                            ColumnName = columnHeader,
                            PropertyName = prop.Name,
                            RawValue = rawValue,
                            Message = error
                        });
                        continue;
                    }

                    prop.SetValue(obj, converted);
                }

                // --------------------- DataAnnotations + IValidatableObject ---------------------
                foreach (var ve in RunValidation(obj, serviceProvider))
                {
                    result.Errors.Add(new ExcelTableError
                    {
                        RowNumber = r,
                        PropertyName = ve.MemberNames.FirstOrDefault() ?? "",
                        Message = ve.ErrorMessage
                    });
                }

                // --------------------- Custom Row Validator ---------------------
                if (customValidator != null)
                {
                    foreach (var msg in customValidator(obj, r))
                    {
                        result.Errors.Add(new ExcelTableError
                        {
                            RowNumber = r,
                            Message = msg
                        });
                    }
                }

                result.Items.Add(obj);
            }

            return result;
        }

        // --------------------------- Validation ---------------------------
        private static List<ValidationResult> RunValidation(object obj, IServiceProvider provider)
        {
            var ctx = new ValidationContext(obj, provider, null);
            var results = new List<ValidationResult>();
            Validator.TryValidateObject(obj, ctx, results, validateAllProperties: true);
            return results;
        }

        // --------------------------- Conversion ---------------------------
        private static object ConvertSafe(string value, Type targetType, out string error)
        {
            error = null;

            if (string.IsNullOrWhiteSpace(value))
                return targetType.IsValueType ? Activator.CreateInstance(targetType) : null;

            Type t = Nullable.GetUnderlyingType(targetType) ?? targetType;

            try
            {
                if (t.IsEnum)
                    return Enum.Parse(t, value, true);

                return Convert.ChangeType(value, t);
            }
            catch
            {
                error = $"Cannot convert '{value}' to {targetType.Name}";
                return targetType.IsValueType ? Activator.CreateInstance(targetType) : null;
            }
        }

        // --------------------------- Range Helpers ---------------------------
        private static void ParseRange(string range, out int colStart, out int colEnd, out int rowStart, out int rowEnd)
        {
            var p = range.Split(':');
            Extract(p[0], out colStart, out rowStart);
            Extract(p[1], out colEnd, out rowEnd);

            static void Extract(string cellRef, out int col, out int row)
            {
                string letters = new(cellRef.TakeWhile(char.IsLetter).ToArray());
                string digits = new(cellRef.SkipWhile(char.IsLetter).ToArray());
                col = ColToNum(letters);
                row = int.Parse(digits);
            }
        }

        private static int ColToNum(string col)
        {
            int sum = 0;
            foreach (char c in col.ToUpper())
                sum = sum * 26 + (c - 'A' + 1);
            return sum;
        }

        private static string GetColName(int index)
        {
            string name = "";
            while (index > 0)
            {
                int rem = (index - 1) % 26;
                name = (char)('A' + rem) + name;
                index = (index - 1) / 26;
            }
            return name;
        }

        private static string ReadCellValue(WorkbookPart wbPart, Cell cell)
        {
            if (cell == null) return "";
            string value = cell.InnerText;

            if (cell.DataType?.Value == CellValues.SharedString)
            {
                var sst = wbPart.SharedStringTablePart.SharedStringTable;
                return sst.ChildElements[int.Parse(value)].InnerText;
            }

            return value;
        }
    }
}
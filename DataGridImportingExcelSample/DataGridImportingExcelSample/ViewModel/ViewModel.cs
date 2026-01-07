using System.Data;
using System.Text;

namespace DataGridImportingExcelSample
{
    /// <summary>
    /// Provides services to import tabular data from user-selected files.
    /// </summary>
    public class ViewModel
    {
        /// <summary>
        /// Prompts the user to pick a file and imports its content into a DataTable.
        /// Supports .xlsx and .csv files.
        /// </summary>
        public async Task<DataTable?> ImportFromFilePickerAsync()
        {
            var customFileTypes = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
            {
                { DevicePlatform.iOS, new[] { "public.comma-separated-values-text", "org.openxmlformats.spreadsheetml.sheet" } },
                { DevicePlatform.MacCatalyst, new[] { "public.comma-separated-values-text", "org.openxmlformats.spreadsheetml.sheet" } },
                { DevicePlatform.Android, new[] { "text/csv", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } },
                { DevicePlatform.WinUI, new[] { ".csv", ".xlsx" } },
            });

            var result = await FilePicker.Default.PickAsync(new PickOptions
            {
                PickerTitle = "Select an Excel (.xlsx) or CSV (.csv) file",
                FileTypes = customFileTypes
            });

            if (result is null)
            {
                return null;
            }

            var extension = Path.GetExtension(result.FileName);
            await using var pickedStream = await result.OpenReadAsync();

            if (extension.Equals(".csv", StringComparison.OrdinalIgnoreCase))
            {
                return ImportFromCsv(pickedStream);
            }

            if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) || extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
            {
                return ImportFromExcel(pickedStream);
            }

            throw new NotSupportedException($"Unsupported file type: {extension}. Only .xlsx and .csv are supported.");
        }

        /// <summary>
        /// Import data from an Excel .xlsx file stream using Syncfusion XlsIO.
        /// </summary>
        private static DataTable ImportFromExcel(Stream excelStream)
        {
            // Use Syncfusion XlsIO's built-in export to avoid platform-specific value issues
            using var memory = new MemoryStream();
            excelStream.CopyTo(memory);
            memory.Position = 0;
            using var excelEngine = new Syncfusion.XlsIO.ExcelEngine();
            var xlsApp = excelEngine.Excel;
            xlsApp.DefaultVersion = Syncfusion.XlsIO.ExcelVersion.Xlsx;
            var workbook = xlsApp.Workbooks.Open(memory);
            var sheet = workbook.Worksheets[0];
            var dataTable = sheet.ExportDataTable(sheet.UsedRange, Syncfusion.XlsIO.ExcelExportDataTableOptions.ColumnNames);

            // remove completely empty rows
            for (int i = dataTable.Rows.Count - 1; i >= 0; i--)
            {
                bool empty = true;
                foreach (var obj in dataTable.Rows[i].ItemArray)
                {
                    if (obj is not null && obj is not DBNull && !(obj is string s && string.IsNullOrWhiteSpace(s)))
                    {
                        empty = false; break;
                    }
                }

                if (empty)
                {
                    dataTable.Rows.RemoveAt(i);
                }
            }

            return dataTable;
        }

        /// <summary>
        /// Import data from a CSV file stream.
        /// Assumes first row contains column headers and comma delimiter with quote support.
        /// </summary>
        private static DataTable ImportFromCsv(Stream csvStream)
        {
            using var reader = new StreamReader(csvStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, leaveOpen: false);
            var headerLine = reader.ReadLine();
            if (headerLine is null)
            {
                return new DataTable();
            }

            var headers = ParseCsvLine(headerLine);
            var table = new DataTable();
            foreach (var h in headers)
            {
                var name = string.IsNullOrWhiteSpace(h) ? "Column" + (table.Columns.Count + 1) : h.Trim();
                if (table.Columns.Contains(name))
                {
                    int index = 1;
                    var baseName = name;
                    while (table.Columns.Contains(name))
                    {
                        name = $"{baseName}_{index++}";
                    }
                }

                table.Columns.Add(name, typeof(string));
            }

            string? line;
            while ((line = reader.ReadLine()) is not null)
            {
                var cells = ParseCsvLine(line);
                var row = table.NewRow();
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    row[i] = i < cells.Count ? cells[i] : string.Empty;
                }

                table.Rows.Add(row);
            }

            return table;
        }

        /// <summary>
        /// Parse a CSV line with support for quoted fields and escaped quotes ("").
        /// </summary>
        private static List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();
            if (line.Length == 0)
            {
                result.Add(string.Empty);
                return result;
            }

            var stringBuilder = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"')
                        {
                            stringBuilder.Append('"');
                            i++;
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        stringBuilder.Append(c);
                    }
                }
                else
                {
                    if (c == ',')
                    {
                        result.Add(stringBuilder.ToString());
                        stringBuilder.Clear();
                    }
                    else if (c == '"')
                    {
                        inQuotes = true;
                    }
                    else
                    {
                        stringBuilder.Append(c);
                    }
                }
            }

            result.Add(stringBuilder.ToString());
            return result;
        }
    }
}

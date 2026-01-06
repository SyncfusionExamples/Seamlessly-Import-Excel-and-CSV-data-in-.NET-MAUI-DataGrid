using Syncfusion.Maui.DataGrid;
using System.Data;
using System.Globalization;

namespace DataGridImportingExcelSample
{
    public partial class MainPage : ContentPage
    {
        private MainViewModel ViewModel { get; } = new();
        public MainPage()
        {
            InitializeComponent();
            BindingContext = ViewModel;
        }

        /// <summary>
        /// Import data from an Excel or CSV file selected by the user.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <exception cref="NotSupportedException"></exception>
        private async void OnImportClicked(object sender, EventArgs e)
        {
            var result = await FilePicker.Default.PickAsync(new PickOptions
            {
                PickerTitle = "Select a data file",
                FileTypes = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
                {
                    // iOS/macOS use UTType identifiers
                    { DevicePlatform.iOS, new[] { "org.openxmlformats.spreadsheetml.sheet", "public.comma-separated-values-text" } },
                    { DevicePlatform.MacCatalyst, new[] { "org.openxmlformats.spreadsheetml.sheet", "public.comma-separated-values-text" } },
                    // Android uses MIME types
                    { DevicePlatform.Android, new[] { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "text/csv" } },
                    // Windows uses file extensions
                    { DevicePlatform.WinUI, new[] { ".xlsx", ".csv" } }
                })
            });

            if (result is null)
            {
                return;
            }

            try
            {
                using var stream = await result.OpenReadAsync();
                var ext = Path.GetExtension(result.FileName).ToLowerInvariant();
                DataTable table = ext switch
                {
                    ".xlsx" => ImportFromExcel(stream),
                    ".csv" => ImportFromCsv(stream),
                    _ => throw new NotSupportedException("Please choose a .xlsx or .csv file.")
                };

                DataGrid.Columns.Clear();
                DataGrid.ItemsSource = table.DefaultView;
                ViewModel.Items = table;
                ViewModel.DynamicItems = null;
                await DisplayAlertAsync("Success", $"Imported {table.Rows.Count} rows and {table.Columns.Count} columns.", "OK");
            }

            catch (Exception ex)
            {
                await DisplayAlertAsync("Import Failed", ex.Message, "OK");
            }
        }

        /// <summary>
        /// Import data from an Excel .xlsx file stream.
        /// </summary>
        /// <param name="excelStream"></param>
        /// <returns></returns>
        private static DataTable ImportFromExcel(Stream excelStream)
        {
            // Use Syncfusion XlsIO'value built-in export to avoid platform-specific value issues
            using var mem = new MemoryStream();
            excelStream.CopyTo(mem);
            mem.Position = 0;
            using var excelEngine = new Syncfusion.XlsIO.ExcelEngine();
            var xlsApp = excelEngine.Excel;
            xlsApp.DefaultVersion = Syncfusion.XlsIO.ExcelVersion.Xlsx;
            var workbook = xlsApp.Workbooks.Open(mem);
            var sheet = workbook.Worksheets[0];
            var dt = sheet.ExportDataTable(sheet.UsedRange, Syncfusion.XlsIO.ExcelExportDataTableOptions.ColumnNames);
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                bool empty = true;
                foreach (var obj in dt.Rows[i].ItemArray)
                {
                    if (obj is not null && obj is not DBNull && !(obj is string s && string.IsNullOrWhiteSpace(s)))
                    {
                        empty = false; break;
                    }
                }

                if (empty)
                {
                    dt.Rows.RemoveAt(i);
                }
            }

            return dt;
        }

        /// <summary>
        /// Import data from a CSV file stream.
        /// </summary>
        /// <param name="csvStream"></param>
        /// <returns></returns>
        private static DataTable ImportFromCsv(Stream csvStream)
        {
            var dt = new DataTable("Imported");
            using var reader = new StreamReader(csvStream);
            var headerLine = reader.ReadLine();
            if (headerLine is null) return dt;
            var headers = SplitCsvLine(headerLine)
                          .Select((h, i) => string.IsNullOrWhiteSpace(h) ? $"Column{i + 1}" : h.Trim())
                          .ToList();

            foreach (var h in headers)
            {
                dt.Columns.Add(h, typeof(object));
            }

            string? line;
            while ((line = reader.ReadLine()) != null)
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                var parts = SplitCsvLine(line);
                var dr = dt.NewRow();
                bool isEmpty = true;
                for (int i = 0; i < headers.Count; i++)
                {
                    var raw = i < parts.Count ? parts[i].Trim() : string.Empty;
                    var value = InferCsvValue(raw);
                    if (value is not DBNull && !(value is string s && string.IsNullOrWhiteSpace(s)))
                    {
                        isEmpty = false;
                    }

                    dr[i] = value;
                }

                if (!isEmpty)
                {
                    dt.Rows.Add(dr);
                }
            }

            return dt;

            static List<string> SplitCsvLine(string line)
            {
                var list = new List<string>();
                bool inQuotes = false;
                var current = new System.Text.StringBuilder();
                for (int i = 0; i < line.Length; i++)
                {
                    char ch = line[i];
                    if (ch == '"')
                    {
                        if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                        {
                            current.Append('"'); i++;
                        }
                        else
                        {
                            inQuotes = !inQuotes;
                        }
                    }
                    else if (ch == ',' && !inQuotes)
                    {
                        list.Add(current.ToString()); current.Clear();
                    }
                    else
                    {
                        current.Append(ch);
                    }
                }

                list.Add(current.ToString());
                return list;
            }

            static object InferCsvValue(string value)
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    return DBNull.Value;
                }
                if (bool.TryParse(value, out var row))
                {
                    return row;
                }
                if (long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var lower))
                {
                    return row;
                }
                if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var obj))
                {
                    return obj;
                }
                if (DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out var date))
                {
                    return date;
                }
                if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out date))
                {
                    return date;
                }

                return value;
            }
        }
    }
}

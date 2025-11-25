using System.Data;
using System.Globalization;
using XlsIO = Syncfusion.XlsIO;

namespace DataGridImportExcelSample
{
    public partial class MainPage : ContentPage
    {
        private MainViewModel ViewModel { get; } = new();

        public MainPage()
        {
            InitializeComponent();
            BindingContext = ViewModel;
        }

        private async void OnImportClicked(object sender, EventArgs e)
        {
            var result = await FilePicker.Default.PickAsync(new PickOptions
            {
                PickerTitle = "Select a data file",
                FileTypes = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
            {
                { DevicePlatform.iOS, new[] { "com.microsoft.excel.xlsx", "public.comma-separated-values-text" } },
                { DevicePlatform.MacCatalyst, new[] { "org.openxmlformats.spreadsheetml.sheet", "public.comma-separated-values-text" } },
                { DevicePlatform.Android, new[] { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "text/csv" } },
                { DevicePlatform.WinUI, new[] { ".xlsx", ".csv" } }
            })
            });

            if (result is null)
                return;

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

                // Update VM and force the grid to regenerate columns
                ViewModel.Items = table;
                DataGrid.Columns.Clear();
                DataGrid.ItemsSource = null;              // force refresh
                DataGrid.ItemsSource = ViewModel.Items;   // rebind

                await DisplayAlertAsync("Success", $"Imported {table.Rows.Count} rows and {table.Columns.Count} columns.", "OK");
            }
            catch (Exception ex)
            {
                await DisplayAlertAsync("Import Failed", ex.Message, "OK");
            }
        }

        // ---------- Excel -> DataTable ----------
        private static DataTable ImportFromExcel(Stream excelStream)
        {
            var dt = new DataTable("Imported");

            using var excelEngine = new XlsIO.ExcelEngine();
            XlsIO.IApplication xlsApp = excelEngine.Excel;
            xlsApp.DefaultVersion = XlsIO.ExcelVersion.Xlsx;

            var workbook = xlsApp.Workbooks.Open(excelStream);
            var sheet = workbook.Worksheets[0];

            int firstRow = sheet.UsedRange.Row;
            int firstCol = sheet.UsedRange.Column;
            int lastRow = sheet.UsedRange.LastRow;
            int lastCol = sheet.UsedRange.LastColumn;

            if (lastRow < firstRow || lastCol < firstCol)
                return dt;

            // Headers
            var headers = new List<string>();
            for (int c = firstCol; c <= lastCol; c++)
            {
                var header = sheet.Range[firstRow, c].Text?.Trim();
                if (string.IsNullOrWhiteSpace(header))
                    header = $"Column{c - firstCol + 1}";
                headers.Add(header);
                dt.Columns.Add(header, typeof(object));
            }

            // Helper: best-typed value
            static object? GetCellValue(XlsIO.IRange cell)
            {
                try
                {
                    var dt = cell.DateTime;
                    if (dt != default) return dt;
                }
                catch { }

                var num = cell.Number;
                if (!double.IsNaN(num)) return num;

                var text = cell.Text?.Trim();
                if (string.IsNullOrEmpty(text)) return DBNull.Value;

                if (bool.TryParse(text, out var b)) return b;
                if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var d)) return d;
                if (DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out var dt1)) return dt1;
                if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var dt2)) return dt2;

                return text;
            }

            // Rows
            for (int r = firstRow + 1; r <= lastRow; r++)
            {
                var dr = dt.NewRow();
                bool isEmpty = true;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    var value = GetCellValue(sheet.Range[r, c]);

                    if (value is not DBNull && value is not null && !(value is string s && string.IsNullOrWhiteSpace(s)))
                        isEmpty = false;

                    dr[c - firstCol] = value ?? DBNull.Value;
                }

                if (!isEmpty)
                    dt.Rows.Add(dr);
            }

            return dt;
        }

        // ---------- CSV -> DataTable ----------
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
                dt.Columns.Add(h, typeof(object));

            string? line;
            while ((line = reader.ReadLine()) != null)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;

                var parts = SplitCsvLine(line);
                var dr = dt.NewRow();
                bool isEmpty = true;

                for (int i = 0; i < headers.Count; i++)
                {
                    var raw = i < parts.Count ? parts[i].Trim() : string.Empty;
                    var value = InferCsvValue(raw);

                    if (value is not DBNull && !(value is string s && string.IsNullOrWhiteSpace(s)))
                        isEmpty = false;

                    dr[i] = value;
                }

                if (!isEmpty)
                    dt.Rows.Add(dr);
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
                        else inQuotes = !inQuotes;
                    }
                    else if (ch == ',' && !inQuotes)
                    {
                        list.Add(current.ToString()); current.Clear();
                    }
                    else current.Append(ch);
                }

                list.Add(current.ToString());
                return list;
            }

            static object InferCsvValue(string s)
            {
                if (string.IsNullOrWhiteSpace(s)) return DBNull.Value;
                if (bool.TryParse(s, out var b)) return b;
                if (long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var l)) return l;
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d)) return d;
                if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out var dt)) return dt;
                if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out dt)) return dt;
                return s;
            }
        }
    }
}
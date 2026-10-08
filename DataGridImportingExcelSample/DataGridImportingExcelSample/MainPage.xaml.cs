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
        private readonly ServiceModel _service = new();
        private async void OnImportClicked(object sender, EventArgs e)
        {
            try
            {
                var table = await _service.ImportFromFilePickerAsync();
                if (table is null)
                {
                    return;
                }

                DataGrid.Columns.Clear();
                DataGrid.ItemsSource = table.DefaultView;
                ViewModel.Items = table;
                await DisplayAlertAsync("Success", $"Imported {table.Rows.Count} rows and {table.Columns.Count} columns.", "OK");
            }
            catch (Exception ex)
            {
                await DisplayAlertAsync("Import Failed", ex.Message, "OK");
            }
        }
    }
}

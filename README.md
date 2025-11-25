# Importing-Excel-Data-into-SfDataGrid-in-.NET-MAUI
This demo shows how to showcase Importing Excel Data into SfDataGrid in .NET MAUI

***

## **Introduction**

Excel files are everywhere—sales reports, inventory lists, customer records. But static spreadsheets don’t belong in modern apps. Imagine transforming those rows and columns into a sleek, interactive grid that works seamlessly across mobile and desktop. With **.NET MAUI** and **Syncfusion SfDataGrid**, you can do exactly that!

In this guide, you’ll learn how to **import Excel data into SfDataGrid using Syncfusion XlsIO**, so your app can deliver dynamic, real-time data visualization without the hassle.

***

## **Why Import Excel Data into SfDataGrid?**

*   **Cross-platform support:** Works on Android, iOS, Windows, and macOS.
*   **Rich UI features:** Sorting, filtering, and responsive columns.
*   **Enterprise-ready:** High performance with virtualization and smooth scrolling.
*   **Seamless integration:** Bind data directly to `ObservableCollection` or `DataTable` for instant updates.

***

## **Architecture Overview**

The process involves:

1.  **Excel File Reading:** Use Syncfusion.XlsIO to read `.xlsx` or `.xls`.
2.  **Data Mapping:** Convert worksheet data into a `DataTable`.
3.  **Binding to SfDataGrid:** Set the `DataTable` as `ItemsSource`.

<!---->

    [Excel File (.xlsx)]
           ↓
    [Syncfusion XlsIO Library] → [DataTable]
           ↓
    [SfDataGrid (UI Control)]

***

## **Prerequisites**

*   .NET MAUI workload installed
*   NuGet packages:
    *   `Syncfusion.Maui.DataGrid`
    *   `Syncfusion.XlsIO.NetCore`

***

## **Build the ViewModel**

Add a ViewModel that implements `INotifyPropertyChanged` for dynamic updates:

```csharp
public class MainViewModel : INotifyPropertyChanged
{
    private DataTable? _items;
    public DataTable? Items
    {
        get => _items;
        set
        {
            if (!ReferenceEquals(_items, value))
            {
                _items = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Items)));
            }
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
}
```

***

## **Build the UI**

Create a simple layout with an **Import button** and **SfDataGrid**:

```xml
<Grid Padding="16" RowSpacing="12">
    <Grid.RowDefinitions>
        <RowDefinition Height="Auto" />
        <RowDefinition Height="*" />
    </Grid.RowDefinitions>

    <Button Grid.Row="0"
            Text="Import from File"
            Clicked="OnImportClicked" />

    <syncfusion:SfDataGrid Grid.Row="1"
                           x:Name="DataGrid"
                           ColumnWidthMode="Fill"
                           HeaderRowHeight="40"
                           RowHeight="40" />
</Grid>
```

***

## **Implement Excel Import Logic**

Use FilePicker to select Excel/CSV and Syncfusion XlsIO to read data:

```csharp
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

    if (result is null) return;

    try
    {
        using var stream = await result.OpenReadAsync();
        var ext = Path.GetExtension(result.FileName).ToLowerInvariant();

        DataTable table = ext switch
        {
            ".xlsx" => ImportFromExcel(stream),
            ".csv"  => ImportFromCsv(stream),
            _       => throw new NotSupportedException("Please choose a .xlsx or .csv file.")
        };

        ViewModel.Items = table;
        DataGrid.Columns.Clear();
        DataGrid.ItemsSource = null;              // force refresh
        DataGrid.ItemsSource = ViewModel.Items;   // rebind

        await DisplayAlertAsync("Success",
            $"Imported {table.Rows.Count} rows and {table.Columns.Count} columns.",
            "OK");
    }
    catch (Exception ex)
    {
        await DisplayAlertAsync("Import Failed", ex.Message, "OK");
    }
}
```

***

### ✅ What this does:

*   Opens a file picker for Excel or CSV.
*   Reads the file into a `DataTable`.
*   Updates the ViewModel and refreshes the SfDataGrid.
*   Displays a success or error message.

***

## **Run the Sample**

*   Click **Import from File**.
*   Select an Excel file.
*   Data loads dynamically into the grid.

***

![Import Excel Data in DataGrid](ImportExcelData.gif)

## **Conclusion**

 Thanks for reading! In this blog, we’ve seen how to showcase bulk editing in [.NET MAUI DataGrid](https://www.syncfusion.com/maui-controls/maui-datagrid). Check out our Release Notes[https://www.syncfusion.com/products/release-history] and [What’s New pages](https://www.syncfusion.com/products/whatsnew) to see the other updates in this release and leave your feedback in the comments section below. 
 For current Syncfusion customers, the newest version of Essential Studio is available from the [license and downloads page](https://www.syncfusion.com/Account/Login?ReturnUrl=%2faccount%2fdownloads). If you are not yet a customer, you can try our 30-day free [trial](https://www.syncfusion.com/downloads) to check out these new features. 
 For questions, you can contact us through our support [forums](https://www.syncfusion.com/forums), [feedback portal](https://www.syncfusion.com/feedback), or support [portal](https://support.syncfusion.com/). We are always happy to assist you!
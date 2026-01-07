using System.ComponentModel;
using System.Data;
using System.Collections.ObjectModel;

namespace DataGridImportingExcelSample
{
    /// <summary>
    /// Main view model for data binding.
    /// </summary>
    public class Model : INotifyPropertyChanged
    {
        private DataTable? _items;
        private ObservableCollection<Dictionary<string, object?>>? _dynamicItems;

        /// <summary>
        /// Items as DataTable for platforms that support DataTable binding.
        /// </summary>
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

        /// <summary>
        /// Dynamic collection for platforms that do not support DataTable binding.
        /// </summary>
        public ObservableCollection<Dictionary<string, object?>>? DynamicItems
        {
            get => _dynamicItems;
            set
            {
                if (!ReferenceEquals(_dynamicItems, value))
                {
                    _dynamicItems = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DynamicItems)));
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }
}
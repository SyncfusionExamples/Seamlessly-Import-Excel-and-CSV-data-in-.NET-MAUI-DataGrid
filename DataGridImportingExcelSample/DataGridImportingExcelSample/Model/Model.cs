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

        public event PropertyChangedEventHandler? PropertyChanged;
    }
}
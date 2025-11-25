using System.ComponentModel;
using System.Data;

namespace DataGridImportExcelSample
{
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

}
using System;
using System.Windows;
using System.Diagnostics;
using System.ComponentModel;

namespace DataVirtualization
{
    public interface SpotRow
    {
        long ID { get; }
        string Leeftijd { get; }
        string Datum { get; }
        string Formaat { get; }
        string Genre { get; }
        string Afzender { get; }
        string Tag { get; }
        string Omvang { get; }
        string Titel { get; }
        string Modulus { get; }
    }

    public sealed class VirtualListItem<T> : INotifyPropertyChanged
    {
        private SpotRow tData;

        private VirtualList<T> _list;
        private int _listVersion;
        private int _index;
        private bool _isLoaded;
        
        internal VirtualListItem(VirtualList<T> list, int index)
        {
            _list = list;
            _listVersion = list.Version;
            _index = index;
        }

        internal VirtualListItem(VirtualList<T> list, int index, T _data) : this(list, index)
        {
            tData = (SpotRow)_data;
            _isLoaded = true;
        }

        internal VirtualList<T> List
        {
            get { return (_list.Version == _listVersion) ? _list : null; }
        }

        internal int Index
        {
            get { return _index; }
        }

        public bool IsLoaded()
        {
            return _isLoaded; 
        }

        public SpotRow Data
        {
            get { return tData; }
        }

        internal void SetData(T _data)
        {
            tData = (SpotRow)_data;
            _isLoaded = true;
        }

        public void Reload()
        {
           OnPropertyChanged(new PropertyChangedEventArgs("Data"));
        }

        public void Load()
        {
            if (_isLoaded)
                return;

            List.Load(Index);
        }

        #region INotifyPropertyChanged Members

        private void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, e);
        }

        event PropertyChangedEventHandler PropertyChanged;
        event PropertyChangedEventHandler INotifyPropertyChanged.PropertyChanged
        {
            add { PropertyChanged += value; }
            remove { PropertyChanged -= value; }
        }

        #endregion
    }
}

using System;
using System.Threading;
using System.ComponentModel;
using System.Diagnostics;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;

namespace DataVirtualization
{
    public partial class VirtualList<T> : IDisposable, IList, IList<VirtualListItem<T>>, INotifyPropertyChanged, INotifyCollectionChanged
    {
        public const int DefaultPageSize = 250;
        private static readonly NotifyCollectionChangedEventArgs _collectionReset = new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset);

        int _version;
        int _pageSize;

        VirtualListItem<T>[] _list;
        IVirtualListLoader<T> _loader;
        
        readonly SynchronizationContext _synchronizationContext;

        public VirtualList(IVirtualListLoader<T> loader, int pageSize, SynchronizationContext synchronizationContext)
        {
            if (loader == null)
                throw new ArgumentNullException("loader");
            if (pageSize <= 0)
                throw new ArgumentOutOfRangeException("pageSize");

            _synchronizationContext = synchronizationContext;

            _version++;
            _loader = loader;
            _pageSize = pageSize;

            LoadPage(0);
        }

        public void Dispose()
        {
            //Dispose(true);
            GC.SuppressFinalize(this);
        }

        //protected virtual void Dispose(bool disposing)
        //{
        //    //if (disposing)
        //        //_pageRequests.Clear();
        //}

        public void Refresh()
        {
            //ThrowIfDeferred();
            //_list = null;
            //SetCurrent(null, -1);
            //LoadPage(0);
        }

        public void Clear()
        {
            _list = null;
            //SetCurrent(null, -1);
            LoadPage(0);
        }

        //public QueuedBackgroundWorkerState LoadingState
        //{
        //    get { return _pageRequests.State; }
        //}

        //public Exception LastLoadingError
        //{
        //    get { return _pageRequests.LastError; }
        //}

        //void OnPageRequestsStateChanged(object sender, EventArgs e)
        //{
        //    if (LoadingStateChanged != null)
        //        LoadingStateChanged(this, EventArgs.Empty);
        //}

        //public event EventHandler LoadingStateChanged;

        //public void RetryLoading()
        //{
        //    if (LoadingState == QueuedBackgroundWorkerState.StoppedByError)
        //        _pageRequests.Retry();
        //}

        private void PopulatePageData(int startIndex, IList<T> pageData, int overallCount)
        {
            bool flagRefresh = false;
            if (_list == null || _list.Length != (overallCount ))
            {
                _list = new VirtualListItem<T>[overallCount];
                flagRefresh = true;
            }

            int l = _list.Length; 

            for (int i = 0; i < pageData.Count; i++)
            {
                int index = startIndex + i;

                if (index < l)
                {
                    if (_list[index] == null)
                    {
                        _list[index] = new VirtualListItem<T>(this, index, pageData[i]);
                    }
                    else
                    {
                        _list[index].SetData(pageData[i]);
                    }
                }
            }

            if (flagRefresh)
            {
                //FireCollectionReset(null);               
                _synchronizationContext.Post(FireCollectionReset, null);
            }
        }

        private void FireCollectionReset(object arg)
        {
            //SetCurrent(null, -1);
            OnCollectionReset();
        }

        internal void Load(int index)
        {
            int startIndex = index - (index % _pageSize);
            LoadRange(startIndex, _pageSize);
        }

        private void LoadPage(int pageIndex)
        {
            int startIndex = pageIndex * _pageSize;
            LoadRange(startIndex, _pageSize);
        }

        private void LoadRange(int startIndex, int count)
        {
            int overallCount;
            IList<T> result = _loader.LoadRange(startIndex, count, null, out overallCount);
            PopulatePageData(startIndex, result, overallCount);
        }

        //internal void LoadAsync(int index)
        //{
        //    int pageIndex = index / _pageSize;
        //    _pageRequests.Add(pageIndex);
        //}

        internal int Version
        {
            get { return _version; }
        }

        public bool Contains(VirtualListItem<T> item)
        {
            return IndexOf(item) != -1;
        }

        public int IndexOf(VirtualListItem<T> item)
        {
            return item == null || item.List != this ? -1 : item.Index;
        }

        public int IndexOf(object value)
        {
            return IndexOf((VirtualListItem<T>)value);
        }

        object IList.this[int index]
        {
            get
            {
                return GetItem(index);
            }
            set
            {
                throw new NotSupportedException();
            }
        }

        public void CopyTo(VirtualListItem<T>[] array, int arrayIndex)
        {
            if (array == null)
                throw new ArgumentNullException("array");
            if (arrayIndex < 0)
                throw new ArgumentOutOfRangeException("arrayIndex");
            if (arrayIndex >= array.Length)
                throw new ArgumentException("arrayIndex is greater or equal than the array length");
            if (arrayIndex + Count > array.Length)
                throw new ArgumentException("Number of elements in list is greater than available space");
            foreach (var item in this)
                array[arrayIndex++] = item;
        }

        public void CopyTo(T[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public void CopyTo(Array array, int index)
        {
            throw new NotImplementedException();
        }

        public int Count
        {
            get { return _list == null ? 0 : _list.Length; }
        }

        VirtualListItem<T> IList<VirtualListItem<T>>.this[int index]
        {
            get
            {
                return GetItem(index);
            }
            set
            {
                throw new NotSupportedException();
            }
        }

        VirtualListItem<T> GetItem(int index)
        {
            if (_list[index] == null)
                _list[index] = new VirtualListItem<T>(this, index);
            return _list[index];
        }

        #region IList<VirtualListItem<T>> Members

        void IList<VirtualListItem<T>>.Insert(int index, VirtualListItem<T> item)
        {
            throw new NotSupportedException();
        }

        void IList<VirtualListItem<T>>.RemoveAt(int index) 
        {
            throw new NotSupportedException();
        }

        void IList.RemoveAt(int index)
        {
            throw new NotSupportedException();
        }


        public void Insert(int index, T item)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region ICollection<VirtualListItem<T>> Members

        void ICollection<VirtualListItem<T>>.Add(VirtualListItem<T> item)
        {
            throw new NotSupportedException();
        }

        public int Add(object value)
        {
            throw new NotImplementedException();
        }
 
        void ICollection<VirtualListItem<T>>.Clear()
        {
            throw new NotSupportedException();
        }

        void IList.Clear()
        {
            throw new NotSupportedException();
        }

        bool ICollection<VirtualListItem<T>>.IsReadOnly
        {
            get { return true; }
        }

        bool IList.IsReadOnly
        {
            get { return true; }
        }

        bool IList.IsFixedSize 
        {
            get { return true; }
        }

        bool ICollection<VirtualListItem<T>>.Remove(VirtualListItem<T> item)
        {
            throw new NotSupportedException();
        }

        public void Remove(object value)
        {
            throw new NotImplementedException();
        }

        public void Insert(int index, object value)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region IEnumerable<VirtualListItem<T>> Members

        IEnumerator<VirtualListItem<T>> IEnumerable<VirtualListItem<T>>.GetEnumerator()
        {
            for (int i = 0; i < Count; i++)
            {
                yield return GetItem(i);
            }
        }

        #endregion

        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            for (int i = 0; i < Count; i++)
            {
                yield return GetItem(i);
            }
        }

        #endregion

        #region INotifyPropertyChanged Members

        protected virtual void OnPropertyChanged(PropertyChangedEventArgs e)
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

        #region INotifyCollectionChanged Members

        protected virtual void OnCollectionChanged(NotifyCollectionChangedEventArgs e)
        {
            if (CollectionChanged != null)
                CollectionChanged(this, e);
        }

        void OnCollectionReset()
        {
            OnCollectionChanged(_collectionReset);
        }

        event NotifyCollectionChangedEventHandler CollectionChanged;
        event NotifyCollectionChangedEventHandler INotifyCollectionChanged.CollectionChanged
        {
            add { CollectionChanged += value; }
            remove { CollectionChanged -= value; }
        }

        #endregion

        public bool Contains(T item)
        {
            throw new NotImplementedException();
        }

        public bool Contains(object value)
        {
            throw new NotImplementedException();
        }

        public object SyncRoot
        {
            get { return this; }
        }

        public bool IsSynchronized
        {
            get { return false; }
        }
    }
}

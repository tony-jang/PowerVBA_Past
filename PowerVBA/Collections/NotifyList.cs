using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowerVBA.Collections
{
    
    public class NotifyList<T> : List<T>, ICloneable
    {
        public NotifyList()
        { }
        public NotifyList(IEnumerable<T> collection) : base(collection)
        { }

        public new void Insert(int index, T item)
        {
            OnPreviewListchanged(item, ChangeAction.Add);
            base.Insert(index, item);
            OnListChanged(item, ChangeAction.Add);
        }

        public new void Add(T item)
        {
            OnPreviewListchanged(item, ChangeAction.Add);
            base.Add(item);
            OnListChanged(item, ChangeAction.Add);
        }
        public new void Remove(T item)
        {
            OnPreviewListchanged(item, ChangeAction.Remove);
            base.Remove(item);
            OnListChanged(item, ChangeAction.Remove);
        }

        List<EventHandler<ChangeEventArgs<T>>> delegates = new List<EventHandler<ChangeEventArgs<T>>>();


        public event EventHandler<ChangeEventArgs<T>> ListChanged;
        public event EventHandler<ChangeEventArgs<T>> PreviewListChanged;

        protected void OnListChanged(T item, ChangeAction action)
        {
            if (ListChanged != null) ListChanged(this, new ChangeEventArgs<T>(item, action));
        }

        protected void OnPreviewListchanged(T item, ChangeAction action)
        {
            if (PreviewListChanged != null) PreviewListChanged(this, new ChangeEventArgs<T>(item, action));
        }

        

        public object Clone()
        {
            NotifyList<T> list = new NotifyList<T>();

            foreach (T itm in this)
            {
                list.Add(itm);
            }

            return list;
        }
    }
}

using PowerVBA.Core.Class;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace PowerVBA.UserControls
{
    public class ImageTreeViewItem : TreeViewItem
    {
        public static DependencyProperty ImageProperty =
            DependencyProperty.Register(nameof(Image), typeof(ImageSource), typeof(ImageTreeViewItem));

        public ImageSource Image
        {
            get { return (ImageSource)GetValue(ImageProperty); }
            set { SetValue(ImageProperty, value); }
        }

        public ImageTreeViewItem()
        {
            
        }
        public ImageTreeViewItem(ImageSource img, string header, iItemData data = null)
        {
            Image = img;
            Header = header;
            this.Tag = Tag;
            _data = data;
        }

        iItemData _data { get; set; }
        public iItemData data
        {
            get { return _data; }
        }
    }




    
}

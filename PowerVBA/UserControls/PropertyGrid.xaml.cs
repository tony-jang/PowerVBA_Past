using PowerVBA.Collections;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PowerVBA.UserControls
{
    /// <summary>
    /// PropertyGrid.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class PropertyGrid : UserControl
    {
        public PropertyGrid()
        {
            InitializeComponent();
            Data.ListChanged += DataChanged;
        }

        private void DataChanged(object sender, ChangeEventArgs<PropertyGridItem> e)
        {
            switch (e.Action)
            {
                case ChangeAction.Add:
                    propGrid.Children.Add(e.Item);
                    break;
                case ChangeAction.Remove:
                    propGrid.Children.Remove(e.Item);
                    break;
            }
        }

        public void AddIntItem(string DicItem, ref int Value)
        {
            var ppitm = new PropertyGridItem(DicItem, ref Value, false);
            Data.Add(ppitm);
        }

        public void AddStrItem(string DicItem, ref string Value)
        {
            var ppitm = new PropertyGridItem(DicItem, ref Value, false);
            Data.Add(ppitm);
        }

        public void AddBoolItem(string DicItem, ref bool Value)
        {
            var ppitm = new PropertyGridItem(DicItem, ref Value, false);
            Data.Add(ppitm);
        }


        public enum Test
        {
            a,b
        }

        public void AddEnumItem(string DicItem, ref int enumData, Type type)
        {
            var ppitm = new PropertyGridItem(DicItem, ref enumData, type, false);
            Data.Add(ppitm);
        }

        private NotifyList<PropertyGridItem> Data = new NotifyList<PropertyGridItem>();

       
    }
}

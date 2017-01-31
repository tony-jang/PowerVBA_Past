using System;
using System.Collections.Generic;
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
    /// PropertyGridItem.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class PropertyGridItem : UserControl
    {
        public enum PropType
        {
            @enum,
            @string,
            @bool,
            @int,
            int2,
            int3,
        }

        public PropType myPropType;

        public void initData()
        {
            
            Grid.SetColumn(pp_TxtB, 2);
            Grid.SetColumn(pp_CB, 2);
            Grid.SetColumn(pp_ComboB, 2);
            pp_TxtB.VerticalAlignment = VerticalAlignment.Center;
            pp_TxtB.FontSize = 14;
            pp_CB.VerticalAlignment = VerticalAlignment.Center;
            pp_ComboB.VerticalAlignment = VerticalAlignment.Center;
        }
        public PropertyGridItem(string name, ref int Data, bool ReadOnly) : base()
        {
            InitializeComponent(); initData();

            myPropType = PropType.@int;
            pp_TxtB.IsEnabled = !ReadOnly;
            pp_TxtB.Text = Data.ToString();
            ppGrid.Children.Add(pp_TxtB);
        }
        public PropertyGridItem(string name, ref short Data, bool ReadOnly) : base()
        {
            InitializeComponent(); initData();
            propName.Text = name;

            myPropType = PropType.@int2;
            pp_TxtB.IsEnabled = !ReadOnly;
            pp_TxtB.Text = Data.ToString();
            ppGrid.Children.Add(pp_TxtB);
        }
        public PropertyGridItem(string name, ref long Data, bool ReadOnly) : base()
        {
            InitializeComponent(); initData();
            propName.Text = name;

            myPropType = PropType.@int3;
            pp_TxtB.IsEnabled = !ReadOnly;
            pp_TxtB.Text = Data.ToString();
            ppGrid.Children.Add(pp_TxtB);
        }
        public PropertyGridItem(string name, ref string Data, bool ReadOnly) : base()
        {
            InitializeComponent(); initData();
            propName.Text = name;

            myPropType = PropType.@string;
            pp_TxtB.IsEnabled = !ReadOnly;
            pp_TxtB.Text = Data;
            ppGrid.Children.Add(pp_TxtB);
        }
        public PropertyGridItem(string name, ref bool Data, bool ReadOnly) : base()
        {
            InitializeComponent(); initData();
            propName.Text = name;

            myPropType = PropType.@bool;
            pp_CB.IsEnabled = !ReadOnly;
            pp_CB.IsChecked = Data;
            ppGrid.Children.Add(pp_CB);
            
        }
        public PropertyGridItem(string name, ref int Data, Type enumType, bool ReadOnly) : base()
        {
            InitializeComponent(); initData();
            propName.Text = name;

            myPropType = PropType.@enum;
            pp_ComboB.IsEnabled = !ReadOnly;

            foreach(string itm in Enum.GetNames(enumType))
                pp_ComboB.Items.Add(itm);
            pp_ComboB.SelectedIndex = Data;
            ppGrid.Children.Add(pp_ComboB);
        }

        
        TextBox pp_TxtB = new TextBox();
        CheckBox pp_CB = new CheckBox();
        ComboBox pp_ComboB = new ComboBox();
        

        

    }
}

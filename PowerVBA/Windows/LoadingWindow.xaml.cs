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
using System.Windows.Shapes;

namespace PowerVBA.Windows
{
    /// <summary>
    /// LoadingWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class LoadingWindow : ChromeWindow
    {

        public LoadingWindow(string msg, int min, int max)
        {
            InitializeComponent();
            infoTB.Text = msg;
            pb.Minimum = min;
            pb.Maximum = max;
        }
        


        public void ValueIncrease()
        {
            if (pb.Value + 1 > pb.Maximum) return;
            valueTB.Text = $"({pb.Value}/{pb.Maximum})";
            pb.Value++;
        }
        public void ValueDecrease()
        {
            if (pb.Value - 1 > pb.Minimum) return;
            valueTB.Text = $"({pb.Value}/{pb.Maximum})";
            pb.Value--;
        }



    }
}

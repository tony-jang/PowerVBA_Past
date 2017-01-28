using PowerVBA.Connector;
using PowerVBA.Windows;
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

namespace PowerVBA
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : PowerVBA.Windows.ChromeWindow
    {
        PresentationConnector pc;
        public MainWindow()
        {
            InitializeComponent();
            pc = new PresentationConnector(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\Macro.pptm");
            VBProjectConnector vbprojConn = new VBProjectConnector(pc.PowerPointPresentation.VBProject);

            foreach(string name in vbprojConn.GetAllProcedureNames())
            {
                MessageBox.Show(name);
            }
            this.Closing += ThisClosing;
        }

        private void ThisClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            pc.Dispose();
        }

        private void mainTabMenu_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (mainTabMenu.SelectedIndex == 0)
            {
                mainTabMenu.SelectedIndex = 1;
            }
        }
    }
}

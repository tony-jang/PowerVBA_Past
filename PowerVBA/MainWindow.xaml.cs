using PowerVBA.Core.Connector;
using PowerVBA.UserControls;
using PowerVBA.Windows;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
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
using ppt = Microsoft.Office.Interop.PowerPoint;
using static PowerVBA.Resources.ResourceImage;
using PowerVBA.Core.Class;

namespace PowerVBA
{

    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : PowerVBA.Windows.ChromeWindow
    {
        PresentationConnector pc;

        Thread thr;

        List<SlideData> slidedatas = new List<SlideData>();

        public MainWindow()
        {
            InitializeComponent();



            var pptItem = new ImageTreeViewItem(GetResourceImage("Component Icon/ppticon_s.png"), "Presentation (프레젠테이션)");
            ImageTreeViewItem slideItem;

            pc = new PresentationConnector(@"F:\장유탁 파일\PowerPoint Game\Buster Wars\U. Buster Wars 1.5.0.pptx", false, false);

            thr = new Thread(() =>
            {
                Stopwatch sw = new Stopwatch();
                sw.Start();
                //Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\Icon.pptx"
                //pc = new PresentationConnector(@"F:\장유탁 파일\Github Project\PowerVBA\PowerVBA\Icon.pptx",false,false);
                
                //VBProjectConnector vbprojConn = new VBProjectConnector(pc.PowerPointPresentation.VBProject);


                ppt.Slides slides = pc.PowerPointPresentation.Slides;
                int SlideNum = slides.Count;

                Dispatcher.Invoke(new Action(() =>
                {
                    pb.Maximum = SlideNum;
                    this.IsEnabled = false;
                }));


                foreach (ppt.Slide slide in pc.PowerPointPresentation.Slides)
                {
                    Dispatcher.Invoke(new Action(() => {
                        pb.Value++;
                        InfoTB.Text = $"슬라이드 정보를 읽어오는 중입니다. ({pb.Value}/{SlideNum})";

                        var sd = new SlideData(false, slide);

                        slidedatas.Add(sd);

                        var shapeItem = new ImageTreeViewItem(GetResourceImage("Component Icon/shapeicon_s.png"), "Shapes (도형 목록)", sd);
                        slideItem = new ImageTreeViewItem(GetResourceImage("Component Icon/slideicon_s.png"), "슬라이드 " + slide.SlideNumber);

                        slideItem.Items.Add(shapeItem);
                        pptItem.Items.Add(slideItem);

                        this.IsEnabled = true;
                        
                    }));
                }
                Dispatcher.Invoke(new Action(() => { pptComponent.Items.Add(pptItem); }));
            });

            thr.SetApartmentState(ApartmentState.STA);
            thr.Start();

            this.Closing += ThisClosing;
            pptComponent.SelectedItemChanged += PptComponent_SelectedItemChanged;
        }

        private void PptComponent_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue.GetType() == typeof(ImageTreeViewItem))
            {
                ImageTreeViewItem itm = (ImageTreeViewItem)e.NewValue;
                iItemData itmdata = itm.data;

                if (itmdata != null)
                    switch (itmdata.type)
                    {
                        case itemDataType.ShapeData:
                            MessageBox.Show("ShapeData");
                            break;
                        case itemDataType.SlideData:
                            SlideData slidedata = (SlideData)itmdata;

                            if (slidedata.IsLoaded) return;
                            pb.Minimum = 0; pb.Maximum = slidedata.Item.Shapes.Count;

                                Thread thr = new Thread(() =>
                                {
                                    foreach (ppt.Slide slide in pc.PowerPointPresentation.Slides)
                                    {

                                    }
                                        foreach (ImageTreeViewItem imageitm in pc.GetShapeItemBySlide(slidedata.Item.SlideNumber))
                                {
                                    Dispatcher.Invoke(new Action(() => {
                                        pb.Value++;
                                        InfoTB.Text = $"슬라이드의 도형 정보를 읽어오는 중입니다. ({pb.Value}/{pb.Maximum})";

                                        itm.Items.Add(imageitm);
                                    }));
                                }

                                    slidedata.IsLoaded = true;

                                });

                                thr.SetApartmentState(ApartmentState.STA);
                                thr.Start();
                           

                            break;
                        default:
                            MessageBox.Show(itmdata.GetType().ToString());
                            break;
                    }

            }
        }

        private void ThisClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                thr.Abort();
                pc.Dispose();
            }
            catch (Exception)
            { }

        }

        private void mainTabMenu_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (mainTabMenu.SelectedIndex == 0)
            {
                mainTabMenu.SelectedIndex = 1;
            }
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            foreach (ppt.Shape shape in pc.GetShapes(false))
            {
                //if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
                //{
                MessageBox.Show(shape.Type.ToString() + " :: " + shape.Name);
                //}

            }
        }
    }
}

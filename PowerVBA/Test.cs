using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using POWERPNT = Microsoft.Office.Interop.PowerPoint;
using VBA = Microsoft.Vbe.Interop;

namespace VBATest
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            return;
            //open(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) +
            //     "\\블록 RPG 과거 구현.pptm");

            Stopwatch sw = new Stopwatch();

            sw.Start();

            List<POWERPNT.Shape> shapes = new List<POWERPNT.Shape>();
            POWERPNT.Application pptApplication = new POWERPNT.Application();
            POWERPNT.Presentation pptPresentation = pptApplication.Presentations.Open(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) +
                 "\\도형+그룹.pptx", MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse); //,

            //MessageBox.Show(pptPresentation.ReadOnly.ToString());
            foreach (POWERPNT.Slide slide in pptPresentation.Slides)
            {
                var layout = slide.Layout;
                //MessageBox.Show(layout.ToString());
                GetAllShapes(slide.Shapes, ref shapes);

                
                

                POWERPNT.Sequence ms = slide.TimeLine.MainSequence;

                

                foreach(POWERPNT.Effect animation in ms)
                {
                    //MessageBox.Show(animation.Shape.Name);
                }
                
            }
            
            int ctr = 0;
            foreach (var shape in shapes)
            {
                
                try
                {
                    //MessageBox.Show(shape.Name);
                    shape.Delete();
                    //shape.Left = 500;
                }
                catch (Exception)
                {
                    ctr++;
                }
                
            }

            MessageBox.Show(shapes.Count + " :: " + sw.ElapsedMilliseconds + " :: " + ctr);

            
            //pptPresentation.Save();
            //pptPresentation.Close();
            //pptPresentation.Close();

            //pptApplication.Quit();
        }

        public void open(string FileName)
        {
            var powerpnt = new POWERPNT.Application();
            var presentation = powerpnt.Presentations.Open(FileName, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

            var project = presentation.VBProject;
            var projectName = project.Name;
            var procedureType = Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Proc;

            foreach (var component in project.VBComponents)
            {
                VBA.VBComponent vbComponent = component as VBA.VBComponent;

                if (vbComponent != null)
                {
                    string componentName = vbComponent.Name;
                    var componentCode = vbComponent.CodeModule;
                    int componentCodeLines = componentCode.CountOfLines;
                    
                    int line = 1;
                    while (line < componentCodeLines)
                    {

                        string procedureName = componentCode.get_ProcOfLine(line, out procedureType);

                        if (procedureName != string.Empty)
                        {

                            int procedureLines = componentCode.get_ProcCountLines(procedureName, procedureType);
                            int procedureStartLine = componentCode.get_ProcStartLine(procedureName, procedureType);
                            int codeStartLine = componentCode.get_ProcBodyLine(procedureName, procedureType);
                            string comments = "[No comments]";

                            if (codeStartLine != procedureStartLine)
                                comments = componentCode.get_Lines(line, codeStartLine - procedureStartLine);

                            int signatureLines = 1;
                            while (componentCode.get_Lines(codeStartLine, signatureLines).EndsWith("_"))
                                signatureLines++;


                            string signature = componentCode.get_Lines(codeStartLine, signatureLines);
                            signature = signature.Replace("\n", string.Empty);
                            signature = signature.Replace("\r", string.Empty);
                            signature = signature.Replace("_", string.Empty);


                            
                            for (int i = 0 + line; i <= procedureLines + line - 1; i++)
                            { 
                                string lineText = componentCode.get_Lines(i, 1);
                                MessageBox.Show(lineText);
                            }
                            line += procedureLines - 1;

                        }

                        line++;
                    }
                }
            }
            powerpnt.Quit();
        }

        public void GetAllShapes(POWERPNT.Shapes shapes, ref List<POWERPNT.Shape> allShapes)
        {
            foreach (POWERPNT.Shape shape in shapes)
            {
                allShapes.Add(shape);

                //shape.TextFrame.TextRange.Text = "!";
                //shape.Child == MsoTriState.msoTrue && 
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    if (shape.GroupItems.Count > 0)
                        GetChildShapes(shape.GroupItems, ref allShapes);
                }
            }
        }

        private static void GetChildShapes(POWERPNT.GroupShapes groupShapes, ref List<POWERPNT.Shape> allShapes)
        {
            foreach (POWERPNT.Shape shape in groupShapes)
            {
                allShapes.Add(shape);

                if (shape.Type == MsoShapeType.msoGroup)
                    GetChildShapes(shape.GroupItems, ref allShapes);
            }
        }

        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }
    }
}

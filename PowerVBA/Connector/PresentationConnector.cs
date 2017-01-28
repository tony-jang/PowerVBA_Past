using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ppt = Microsoft.Office.Interop.PowerPoint;
using VBA = Microsoft.Vbe.Interop;
using PowerVBA.Converter;
using System.Windows;

namespace PowerVBA.Connector
{
    public class PresentationConnector : IDisposable
    {
        ppt.Application pptApp;
        ppt.Presentation pptPresentation;
        
        /// <summary>
        /// PresentationConnector를 초기화시킵니다.
        /// </summary>
        /// <param name="FileLocation">파일 위치입니다.</param>
        /// <param name="OpenWithWindow">true일시 파워포인트 창이 같이 뜹니다.</param>
        public PresentationConnector(string FileLocation,bool Untitled = false, bool OpenWithWindow = true)
        {
            pptApp = new ppt.Application();
            pptPresentation = pptApp.Presentations.Open(FileLocation, MsoTriState.msoFalse, boolConverter.boolToState(Untitled), boolConverter.boolToState(OpenWithWindow));

            MessageBox.Show(pptPresentation.HasVBProject.ToString());
        }

        /// <summary>
        /// 슬라이드를 가져옵니다.
        /// </summary>
        /// <returns></returns>
        public ppt.Slides GetSlides()
        {
            return pptPresentation.Slides;
        }

        public ppt.Presentation PowerPointPresentation
        {
            get { return pptPresentation; }
        }
        public List<ppt.Shape> GetShapes(bool GetAllChildShapes)
        {
            List<ppt.Shape> shapes = new List<ppt.Shape>();
            foreach(ppt.Slide slide in GetSlides())
                shapes.AddRange(ShapesToList(slide.Shapes, GetAllChildShapes));
              
            return shapes;
        }

        public List<ppt.Shape> ShapesToList(ppt.Shapes pptshapes, bool GetAllShapes)
        {
            List<ppt.Shape> shapes = new List<ppt.Shape>();
            foreach(ppt.Shape shape in pptshapes)
            {
                shapes.Add(shape);
                if (GetAllShapes && shape.Type == MsoShapeType.msoGroup)
                {
                    shapes.AddRange(ShapesToList((ppt.Shapes)shape.GroupItems,true));
                }
            }

            return shapes;
        }

        public void GetAllShapes(ppt.Shapes shapes, ref List<ppt.Shape> allShapes)
        {
            foreach (ppt.Shape shape in shapes)
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

        private static void GetChildShapes(ppt.GroupShapes groupShapes, ref List<ppt.Shape> allShapes)
        {
            foreach (ppt.Shape shape in groupShapes)
            {
                allShapes.Add(shape);

                if (shape.Type == MsoShapeType.msoGroup)
                    GetChildShapes(shape.GroupItems, ref allShapes);
            }
        }


        /// <summary>
        /// 애니메이션 시퀀스를 가져옵니다.
        /// </summary>
        /// <param name="slidenumber">슬라이드 번호입니다.</param>
        /// <returns></returns>
        public ppt.Sequence AnimationTimeLine(int slidenumber)
        {
            return pptPresentation.Slides[slidenumber].TimeLine.MainSequence;
        }
        

        public void Dispose()
        {
            pptPresentation.Close();
            pptApp.Quit();
        }
    }
}

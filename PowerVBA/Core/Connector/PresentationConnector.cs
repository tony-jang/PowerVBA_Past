using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ppt = Microsoft.Office.Interop.PowerPoint;
using VBA = Microsoft.Vbe.Interop;
using PowerVBA.Core.Converter;
using System.Windows;
using PowerVBA.UserControls;
using System.Windows.Media.Imaging;
using static PowerVBA.Core.Converter.ShapeConverter;
using static PowerVBA.Resources.ResourceImage;
using PowerVBA.Core.Class;

namespace PowerVBA.Core.Connector
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

            //MessageBox.Show(pptPresentation.HasVBProject.ToString());
        }




        #region [ 아이템 반환 ]
        

        public List<ImageTreeViewItem> GetSlidesItem()
        {
            var listitm = new List<ImageTreeViewItem>();

            foreach(ppt.Slide slide in GetSlides())
            {
                ImageTreeViewItem itm = new ImageTreeViewItem();
                itm.Image = GetResourceImage("Component Icon/SlideIcon_s.png");
                itm.Header = $"Slide{slide.SlideNumber} (슬라이드)";
                itm.Tag = slide;

                listitm.Add(itm);
            }

            return listitm;
        }

        public List<ImageTreeViewItem> GetShapeItemBySlide(int SlideNumber)
        {
            if (SlideNumber > pptPresentation.Slides.Count) return null;
            BitmapImage shapeimage = GetResourceImage("Component Icon/ShapeIcon_s.png");
            var Itm = new List<ImageTreeViewItem>();

            foreach (ImageTreeViewItem shapeitm in GetShapeItem(ShapesToList(GetSlides()[SlideNumber].Shapes, false),true))
            {
                Itm.Add(shapeitm);
            }

            return Itm;
        }




        public List<ImageTreeViewItem> GetShapeItem(ppt.Shapes shapes, bool GetAllGroupItem)
        {
            return GetShapeItem(ShapesToList(shapes, false), GetAllGroupItem);
        }

        public List<ImageTreeViewItem> GetShapeItem(List<ppt.Shape> shapelist, bool GetAllGroupItem)
        {
            BitmapImage shapeimage = GetResourceImage("Component Icon/ShapeIcon_s.png");
            var itm = new List<ImageTreeViewItem>();

            foreach (ppt.Shape shape in shapelist)
            {
                var shapeitm = new ImageTreeViewItem(shapeimage, shape.Name + " (도형)", new ShapeData(false, shape));

                if (shape.Type == MsoShapeType.msoGroup)
                    foreach (ImageTreeViewItem inneritm in GetShapeItem(GroupShapesToShapes(shape.GroupItems), GetAllGroupItem))
                        shapeitm.Items.Add(inneritm);

                itm.Add(shapeitm);
            }
            return itm;
        }
        #endregion




        #region [ 프레젠테이션 속성 반환 ]


        /// <summary>
        /// 현재 프레젠테이션의 슬라이드들을 가져옵니다.
        /// </summary>
        /// <returns></returns>
        public ppt.Slides GetSlides()
        {
            return pptPresentation.Slides;
        }

        /// <summary>
        /// 현재 파워포인트 프레젠테이션을 반환해줍니다.
        /// </summary>
        public ppt.Presentation PowerPointPresentation
        {
            get { return pptPresentation; }
        }

        public List<ppt.Shape> GetShapes(bool GetAllChildShapes)
        {
            List<ppt.Shape> shapes = new List<ppt.Shape>();
            foreach (ppt.Slide slide in GetSlides())
                shapes.AddRange(ShapesToList(slide.Shapes, GetAllChildShapes));

            return shapes;
        }
        

        /// <summary>
        /// allShapes에 모든 <see cref="ppt.Shape"/>를 추가합니다.
        /// </summary>
        /// <param name="shapes"></param>
        /// <param name="allShapes"></param>
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


        #endregion




        public void Dispose()
        {
            pptPresentation.Close();
            pptApp.Quit();
        }
    }
}

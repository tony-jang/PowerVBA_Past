using Microsoft.Office.Core;
using PowerVBA.Core.Class;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ppt = Microsoft.Office.Interop.PowerPoint;
using VBA = Microsoft.Vbe.Interop;

namespace PowerVBA.Core.Converter
{
    static class ShapeConverter
    {
        public static List<ppt.Shape> GroupShapesToShapes(ppt.GroupShapes groupitms)
        {
            var itm = new List<ppt.Shape>();
            foreach (ppt.Shape shape in groupitms)
            {
                itm.Add(shape);
            }
            return itm;
        }
        /// <summary>
        /// Shapes를 <see cref="List{}"/>로 변환해줍니다.
        /// </summary>
        /// <param name="pptshapes"></param>
        /// <param name="GetAllShapes"></param>
        /// <returns></returns>
        public static List<ppt.Shape> ShapesToList(ppt.Shapes pptshapes, bool GetAllShapes)
        {
            List<ppt.Shape> shapes = new List<ppt.Shape>();
            foreach (ppt.Shape shape in pptshapes)
            {
                shapes.Add(shape);

                if (GetAllShapes && shape.Type == MsoShapeType.msoGroup)
                {
                    shapes.AddRange(ShapesToList((ppt.Shapes)shape.GroupItems, true));
                }
            }

            return shapes;
        }
        

        public static CustomShapeData ShapeToCustomShapeData(int slideNumber, ppt.Shape shape)
        {
            List<string> strings = new List<string>();
            ppt.Shape shpe = shape;
            do
            {
                strings.Add(shpe.Name);
                try
                { shpe = shpe.ParentGroup; }
                catch (Exception)
                { break; }
                
            } while (true);
            

            return new CustomShapeData(slideNumber, strings.ToArray().Reverse().ToArray());
        }
        public static ppt.Shape CustomShapeDataToShape(CustomShapeData shapedata, ppt.Presentation pptPresentation)
        {
            ppt.Shape BaseShape = pptPresentation.Slides[shapedata.SlideNumber].Shapes[shapedata.Indexes[0]];

            for (int i = 1; i<=shapedata.Indexes.Count() - 1; i++)
            {
                BaseShape = BaseShape.GroupItems[shapedata.Indexes[i]];
            }

            return BaseShape;
        }
    }
}

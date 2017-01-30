using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    }
}

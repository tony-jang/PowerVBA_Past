using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ppt = Microsoft.Office.Interop.PowerPoint;

namespace PowerVBA.Core.Class
{
    public struct ShapeData : iItemData
    {
        public bool IsLoaded { get; set; }
        public CustomShapeData Item { get; set; }

        public itemDataType type { get; }

        public ShapeData(bool loaded, CustomShapeData itm)
        {
            IsLoaded = loaded;
            Item = itm;
            type = itemDataType.ShapeData;
        }
    }

    public struct CustomShapeData
    {
        public int SlideNumber;
        public string[] Indexes;

        public CustomShapeData(int slidenumber, string[] indexes)
        {

            SlideNumber = slidenumber;
            Indexes = indexes;
        }

    }
}

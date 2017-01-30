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
        public ppt.Shape Item { get; set; }

        public itemDataType type { get; }

        public ShapeData(bool loaded, ppt.Shape itm)
        {
            IsLoaded = loaded;
            Item = itm;
            type = itemDataType.ShapeData;
        }
    }
}

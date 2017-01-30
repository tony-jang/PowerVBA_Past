using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ppt = Microsoft.Office.Interop.PowerPoint;

namespace PowerVBA.Core.Class
{


    public struct SlideData : iItemData
    {
        public bool IsLoaded { get; set; }

        public ppt.Slide Item { get; set; }

        public itemDataType type { get; }
                

        public SlideData(bool loaded, ppt.Slide itm)
        {
            IsLoaded = loaded;
            Item = itm;
            type = itemDataType.SlideData;
        }

    }
}

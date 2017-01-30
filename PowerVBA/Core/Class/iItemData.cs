using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Class
{


    public interface iItemData
    {
        bool IsLoaded { get; set; }
        itemDataType type { get; } 
    }

    public enum itemDataType
    {
        None = 0,
        ShapeData = 1,
        SlideData = 2
    }
}

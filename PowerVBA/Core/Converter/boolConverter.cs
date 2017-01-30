using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Converter
{
    static class boolConverter
    {
        public static bool StateTobool(MsoTriState state)
        {
            if (state == MsoTriState.msoFalse) return false;
            else if (state == MsoTriState.msoTrue) return true;

            return false;
        }

        public static MsoTriState boolToState(bool Bool)
        {
            if (Bool) return MsoTriState.msoTrue;
            else return MsoTriState.msoFalse;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ppt = Microsoft.Office.Interop.PowerPoint;
using VBA = Microsoft.Vbe.Interop;

namespace PowerVBA.Core.Connector.Code
{
    class VBACodeConnector
    {
        private int Slidenumber;
        private VBA.VBComponent Comp;

        private string code;
        public VBACodeConnector(int slidenumber, VBA.VBComponent comp)
        {
            Slidenumber = slidenumber;
            Comp = comp;
        }

        List<string> AllProcedureNames;
        public void Update()
        {
            List<string> list = new List<string>();
            var procedureType = VBA.vbext_ProcKind.vbext_pk_Proc;
            VBA.VBComponent comp = Comp;

            if (comp != null)
            {
                // 파일 이름
                string componentName = comp.Name;
                var compCode = comp.CodeModule;
                int codeLines = compCode.CountOfLines;

                int line = 1;

                while (line < codeLines)
                {
                    string procedureName = compCode.get_ProcOfLine(line, out procedureType);


                    if (procedureName != string.Empty)
                    {
                        int procedureLines = compCode.get_ProcCountLines(procedureName, procedureType);
                        int procedureStartLine = compCode.get_ProcStartLine(procedureName, procedureType);
                        int codeStartLine = compCode.get_ProcBodyLine(procedureName, procedureType);
                        list.Add(procedureName);
                        line += procedureLines - 1;
                    }


                    line++;
                }
            }

            AllProcedureNames = list;
        }

    }


}

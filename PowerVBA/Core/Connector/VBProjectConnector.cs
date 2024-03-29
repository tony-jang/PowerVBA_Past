﻿using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ppt = Microsoft.Office.Interop.PowerPoint;
using VBA = Microsoft.Vbe.Interop;
using PowerVBA.Core.Converter;
using System.Windows;
using PowerVBA.Core.Connector;
using System.Text.RegularExpressions;

namespace PowerVBA.Core.Connector
{
    class VBProjectConnector : IDisposable
    {
        private VBA.VBProject VBProj;
        private VBA.VBComponents VBComps;
        public VBProjectConnector(VBA.VBProject vbproj)
        {
            VBProj = vbproj;
            VBComps = vbproj.VBComponents;
            //VBA.VBComponent newStandardModule = VBProj.VBComponents.Add(VBA.vbext_ComponentType.vbext_ct_StdModule);

            //newStandardModule.CodeModule.Name = "ABC";
            Update();
            return;
        }

        private List<string> AllProcedureNames;
        public List<string> GetAllProcedureNames
        {
            get { Update(); return AllProcedureNames; }
        }
        private Dictionary<string, string[]> CodeDictionary;

        public void Update()
        {
            List<string> list = new List<string>();
            var procedureType = VBA.vbext_ProcKind.vbext_pk_Proc;
            foreach (VBA.VBComponent comp in VBComps)
            {
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
            }
            AllProcedureNames = list;
        }
        
        public bool FilterName(string name)
        {
            string pattern = @"^[_|a-zA-z가-힣ㅏ-ㅣㄱ-ㅎ][_|a-zA-Z가-힣ㅏ-ㅣㄱ-ㅎ1-9]*$";

            if (Regex.Match(name, pattern).Success) return true;

            return false;
        }


        public bool InsertProcedure(Accessor accessor, string name, string ReturnType = "")
        {
            VBA.VBComponent newStandardModule = VBProj.VBComponents.Add(VBA.vbext_ComponentType.vbext_ct_StdModule);

            newStandardModule.Name = "Test Module";

            string code = "";
            if (accessor == Accessor.Private) code = "Private";
            else if (accessor == Accessor.Public) code = "Public";

            code += " ";

            if (ReturnType == "") code += "Sub";
            else code += "Function";

            if (!FilterName(name)) return false;
            
            code += " " + name + "() ";

            if (ReturnType != "") code += $"As {ReturnType}";

            MessageBox.Show(code);

            

            return true;
        }


        public void Dispose()
        {

        }
    }
}

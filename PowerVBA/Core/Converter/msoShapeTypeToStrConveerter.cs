using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Converter
{
    static class msoShapeTypeToStrConverter
    {
        //bracket. parenthesis = 괄호 의미 단어
        public static string MsoShapeTypeToString(MsoShapeType type, bool WithFormat = true)
        {
            string str = "오브젝트";
            switch (type)
            {
                case MsoShapeType.msoAutoShape: str = "도형"; break;
                case MsoShapeType.msoCanvas: str = "캔버스"; break;
                case MsoShapeType.msoCallout: str = "호출"; break;
                case MsoShapeType.msoChart: str = "차트"; break;
                case MsoShapeType.msoDiagram: str = "다이어그램"; break;
                case MsoShapeType.msoComment: str = "코멘트"; break;
                case MsoShapeType.msoContentApp: str = "컨텐츠 앱"; break;
                case MsoShapeType.msoOLEControlObject: str = "OLE 컨트롤"; break;
                case MsoShapeType.msoFreeform: str = "자유형"; break;
                case MsoShapeType.msoGroup: str = "그룹"; break;
                case MsoShapeType.msoLine: str = "선"; break;
                case MsoShapeType.msoMedia: str = "미디어"; break;
                case MsoShapeType.msoPicture: str = "사진"; break;
                case MsoShapeType.msoPlaceholder: str = "위치 홀더"; break;
                case MsoShapeType.msoSmartArt: str = "스마트 아트"; break;
                case MsoShapeType.msoTable: str = "표"; break;
                case MsoShapeType.msoTextBox: str = "텍스트 박스"; break;
                case MsoShapeType.msoTextEffect: str = "텍스트 효과"; break;
                
            }
            if (WithFormat) return $"({str})";
            return str;
            
        }
    }
}

using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace PowerVBA.Resources
{
    static class ResourceImage
    {
        private static string BaseURL = "/PowerVBA;Component/Resources/";
        public static BitmapImage GetResourceImage(string filename)
        {
            return new BitmapImage(new Uri(BaseURL + filename, UriKind.Relative));
        }
        public static BitmapImage GetResourceIcon(MsoShapeType type)
        {
            string URLList = "Component Icon/";
            string ImageName = "";
            switch (type)
            {
                case MsoShapeType.msoAutoShape:
                    ImageName = "shapeicon_s";
                    break;
                case MsoShapeType.msoChart:
                    ImageName = "charticon_s";
                    break;
                case MsoShapeType.msoGroup:
                    ImageName = "groupicon_s";
                    break;
                case MsoShapeType.msoPicture:
                    ImageName = "pictureicon_s";
                    break;
                case MsoShapeType.msoCanvas:
                    ImageName = "canvasicon_s";
                    break;
                case MsoShapeType.msoDiagram:
                    ImageName = "diagramicon_s";
                    break;
                case MsoShapeType.msoTextEffect:
                    ImageName = "Effecticon_s";
                    break;
                case MsoShapeType.msoFreeform:
                    ImageName = "Freeformicon_s";
                    break;
                case MsoShapeType.msoLinkedOLEObject:
                case MsoShapeType.msoLinkedPicture:
                    ImageName = "hyperlinkicon_s";
                    break;
                case MsoShapeType.msoSmartArt:
                    ImageName = "smartarticon_s";
                    break;
                case MsoShapeType.msoTable:
                    ImageName = "tableicon_s";
                    break;
                case MsoShapeType.msoTextBox:
                    ImageName = "textboxicon_s";
                    break;
                case MsoShapeType.msoOLEControlObject:
                    ImageName = "olectrlobjicon_s";
                    break;
                case MsoShapeType.msoEmbeddedOLEObject:
                    ImageName = "objecticon_s";
                    break;
                case MsoShapeType.msoLine:
                    ImageName = "lineicon_s";
                    break;
                case MsoShapeType.msoMedia:
                    ImageName = "mediaicon_s";
                    break;
                case MsoShapeType.msoPlaceholder:
                    //TODO : Add PlaceHolder Icon
                    ImageName = "placeholdericon_s";
                    break;
                default:
                    ImageName = "objecticon_s";
                    break;
            }

            return GetResourceImage(URLList + ImageName + ".png");
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using ESRI.ArcGIS.ADF.BaseClasses;
using ESRI.ArcGIS.ADF.CATIDs;
using System.Drawing;

namespace Shape2KML
{
    //represents Style for all KML elements
    class KMLStyle
    {
        Color lineColor;

        public Color LineColor
        {
            get { return lineColor; }
            set { lineColor = value; }
        }
        //byte lineOpacity;
        Color fillColor;

        public Color FillColor
        {
            get { return fillColor; }
            set { fillColor = value; }
        }
       // byte fillOpacity;
        Color iconColor;

        public Color IconColor
        {
            get { return iconColor; }
            set { iconColor = value; }
        }
        //byte iconOpacity;
        string iconPath;

        public string IconPath
        {
            get { return iconPath; }
            set { iconPath = value; }
        }

    }
}

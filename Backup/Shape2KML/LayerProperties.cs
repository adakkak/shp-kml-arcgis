using System;
using System.Collections.Generic;
using System.Text;
using ESRI.ArcGIS.ADF.BaseClasses;
using ESRI.ArcGIS.ADF.CATIDs;
using System.Drawing;
using ESRI.ArcGIS.Carto;

namespace Shape2KML
{
    public enum AltitudeMode
    {
        clampToGround,
        relativeToGround,
        absolute
    }

    class LayerProperties
    {
            public string name;
            private AltitudeMode altitudeMode;

            public AltitudeMode AltitudeMode
            {
                get { return altitudeMode; }
                set { altitudeMode = value; }
            }
            private int multiplier = 1;

            public int Multiplier
            {
                get { return multiplier; }
                set { multiplier = value; }
            }
            private int altitude = 0;

            public int Altitude
            {
                get { return altitude; }
                set { altitude = value; }
            }
            private string field;

            public string Field
            {
                get { return field; }
                set { field = value; }
            }
            private string descField = "";

            public string DescField
            {
                get { return descField; }
                set { descField = value; }
            }
            private string nameField = "";

            public string NameField
            {
                get { return nameField; }
                set { nameField = value; }
            }
            //public bool tesselate;
            //public bool extrude;
            private Color color;

            public Color Color
            {
                get { return color; }
                set { color = value; }
            }
            private ILayer layeru;

            public ILayer Layeru
            {
                get { return layeru; }
                set { layeru = value; }
            }
            private bool tesselate = false;

            public bool Tesselate
            {
                get { return tesselate; }
                set { tesselate = value; }
            }
            private bool extrude = false;

            public bool Extrude
            {
                get { return extrude; }
                set { extrude = value; }
            }
          
            public override string ToString()
            {
                return Layeru.Name;
            }
    }
}

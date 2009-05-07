using System;
using System.Collections.Generic;
using System.Text;
using ESRI.ArcGIS.ADF.BaseClasses;
using ESRI.ArcGIS.ADF.CATIDs;
using Google.KML;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geometry;
using System.Diagnostics;
using System.IO;
using ESRI.ArcGIS.Geodatabase;
using System.Windows.Forms;

namespace Shape2KML
{
    //used to generate a KML Document based on the layers and style selected
    class KMLGenerator
    {
        //layerele care vor fi adaugate in KML
        List<LayerProperties> kmlLayers;

        internal List<LayerProperties> KmlLayers
        {
            get { return kmlLayers; }
            set { kmlLayers = value; }
        }
        geStyle defaultStyle;

        string pathu;
        internal string Path
        {
            get {return pathu;}
            set { pathu = value; }
        }

        internal geStyle DefaultStyle
        {
            get { return defaultStyle; }
            set { defaultStyle = value; }
        }
        geDocument kmlDocument;

        public geDocument KmlDocument
        {
            get { return kmlDocument; }
            set { kmlDocument = value; }
        }

                
        //Ia lista de Layere impreuna cu proprietatile lor si genereaza un fisier kml din ele.
        
        public KMLGenerator(List<LayerProperties> layers)
        {
            kmlLayers = layers;
            kmlDocument = new geDocument();
        }

        private void processLayerProperties(LayerProperties lp, geFolder folder)
        {
      
          if (lp.Layeru is IFeatureLayer)
            {
                if ((lp.Layeru as IFeatureLayer).FeatureClass.ShapeType == ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline)
                    processPolyLineLayer(lp,folder);
                if ((lp.Layeru as IFeatureLayer).FeatureClass.ShapeType == ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint)
                    processPointLayer(lp,folder);
                if ((lp.Layeru as IFeatureLayer).FeatureClass.ShapeType == ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon)
                    processPolygonLayer(lp,folder);

            }
            if (lp.Layeru is IRasterLayer)
                processRasterLayer(lp,folder);
        }

        private void processRasterLayer(LayerProperties layerProps, geFolder folder)
        {
            IRasterLayer currentLayer = layerProps.Layeru as IRasterLayer;
                        
            //get coordinates for image :D
            geAngle90 north;
            geAngle90 south;
            geAngle180 east;
            geAngle180 west;
            Double coordlat, coordlong;
            IPoint refPoint;
            
            refPoint = currentLayer.AreaOfInterest.LowerLeft;

            //coordinate system.
            IGeographicCoordinateSystem gcs;
            SpatialReferenceEnvironment sre = new SpatialReferenceEnvironment();

            //create coordinate system for WGS 84 (Google earth)
            gcs = sre.CreateGeographicCoordinateSystem((int)esriSRGeoCSType.esriSRGeoCS_WGS1984);

            //spatial reference 
            ISpatialReference pointToSpatialReference;
            pointToSpatialReference = gcs;

            //project point to google earth projection 
            refPoint.Project(pointToSpatialReference);

            //and get coordinates
            refPoint.QueryCoords(out coordlong, out coordlat);
            //left = west point and lower = south
            west= new geAngle180(coordlong);
            south = new geAngle90(coordlat);

            //and upper right
            refPoint = currentLayer.AreaOfInterest.UpperRight;

            //project point to google earth projection 
            refPoint.Project(pointToSpatialReference);
            //x.ToString("{0:0.00}");

            //and get coordinates
            refPoint.QueryCoords(out coordlong, out coordlat);

            //north and east
            north = new geAngle90(coordlat);
            east = new geAngle180(coordlong);

            //so I have all coordinates, now I create the overlay.
            geGroundOverlay.geLatLonBox latlonbox = new geGroundOverlay.geLatLonBox(north,south,east,west);
            geGroundOverlay overlay = new geGroundOverlay(latlonbox);

            //copy image to current folder...
            string name =  System.IO.Path.GetDirectoryName(Path);
            name = System.IO.Path.Combine(name, currentLayer.Name);
            File.Copy(currentLayer.FilePath, name,true);

           // overlay.Icon.Href = currentLayer.FilePath;
            overlay.Icon = new geIcon(System.IO.Path.GetFileName(name));
            overlay.Name = currentLayer.Name;
            overlay.Description = currentLayer.FilePath;
            overlay.StyleUrl = "#Shape2KMLGeneratedStyle";

            switch (layerProps.AltitudeMode)
            {
                case AltitudeMode.absolute:
                    overlay.AltitudeMode = geAltitudeModeEnum.absolute;
                    overlay.Altitude = layerProps.Altitude;
                    //use altitude from field , do that later...
                    break;
                case AltitudeMode.clampToGround:
                    overlay.AltitudeMode = geAltitudeModeEnum.clampToGround;
                    break;
                case AltitudeMode.relativeToGround:
                    overlay.AltitudeMode = geAltitudeModeEnum.relativeToGround;
                    overlay.Altitude = layerProps.Altitude;
                    //use altitude from field , do that later...
                    break;
                default:
                    break;
            }



            

            //and add overlay to folder. THAT "simple"
            folder.Features.Add(overlay);

        }

        private void processPolyLineLayer(LayerProperties layerProps, geFolder folder)
        {
            //I have a polyline layer, and a kml folder :)
            IFeatureClass clasa = (layerProps.Layeru as IFeatureLayer).FeatureClass;
            //get acces to features
            IFeatureCursor featurele = clasa.Search(null, true);
            int nrFeature = clasa.FeatureCount(null);

            //if I have any features
            Polyline curba;
            IFeature currentFeature;

            while ((currentFeature = featurele.NextFeature()) != null)
            {
                curba = currentFeature.Shape as Polyline;
                //coordinates and vertices
                double coordLat;
                double coordLong;
                IEnumVertex colection = curba.EnumVertices;
                IPoint lineVertex;

                //create coord system WGS 84 (Google earth)
                IGeographicCoordinateSystem gcs;
                SpatialReferenceEnvironment sre = new SpatialReferenceEnvironment();
                gcs = sre.CreateGeographicCoordinateSystem((int)esriSRGeoCSType.esriSRGeoCS_WGS1984);
                //create spatial reference
                ISpatialReference pointToSpatialReference;
                pointToSpatialReference = gcs;

                #region add points to polyline
                //create a placemark for the line
                gePlacemark pmLine = new gePlacemark();
                pmLine.StyleUrl = "#Shape2KMLGeneratedStyle";

                List<geCoordinates> lineCoords = new List<geCoordinates>();
                int index1, index2;
                //iterate points...

                geLineString line = null;

                while (!colection.IsLastInPart())
                {
                    colection.Next(out lineVertex, out index1, out index2);
                    //project point and get coordinates
                    lineVertex.Project(pointToSpatialReference);
                    lineVertex.QueryCoords(out coordLong, out coordLat);

                    try
                    {
                    switch (layerProps.AltitudeMode)
                    {
                        case AltitudeMode.absolute:
                            if (layerProps.Field == "")
                            {
                                //add point to line                  
                                lineCoords.Add(new geCoordinates(new geAngle90(coordLat), new geAngle180(coordLong), layerProps.Altitude));
                            }
                            else
                            {
                                int altitude;
                                //if altitude is integer, this should work   
                                    altitude = (int)currentFeature.get_Value(currentFeature.Fields.FindField(layerProps.Field));
                               
                                //add point to line                  
                                    lineCoords.Add(new geCoordinates(new geAngle90(coordLat), new geAngle180(coordLong), layerProps.Multiplier * altitude));
                            
                            }
                            break;

                        case AltitudeMode.clampToGround:
                            //add point to line                  
                            lineCoords.Add(new geCoordinates(new geAngle90(coordLat), new geAngle180(coordLong)));
                            break;

                        case AltitudeMode.relativeToGround:
                           if (layerProps.Field == "")
                            {
                              //add point to line                  
                            lineCoords.Add(new geCoordinates(new geAngle90(coordLat), new geAngle180(coordLong), layerProps.Altitude));
                        }
                        else
                        {
                            float altitude;
                             //if altitude is integer, this should work   
                            altitude = (float)currentFeature.get_Value(currentFeature.Fields.FindField(layerProps.Field));

                            //add point to line                  
                            lineCoords.Add(new geCoordinates(new geAngle90(coordLat), new geAngle180(coordLong), layerProps.Multiplier * altitude));

                        }
                               
                        break;

                        default:
                            break;
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Altitude field is not a number value");
                    break;
                }
                }
                //create line from list of coords
                line = new geLineString(lineCoords);
             
                switch (layerProps.AltitudeMode)
                {   
                    case AltitudeMode.absolute:
                        line.AltitudeMode = geAltitudeModeEnum.absolute;
                        break;
                    case AltitudeMode.clampToGround:
                        line.AltitudeMode = geAltitudeModeEnum.clampToGround;
                        break;
                    case AltitudeMode.relativeToGround:
                        line.AltitudeMode = geAltitudeModeEnum.relativeToGround;
                        break;
                    default:
                        break;
                }
                    

                if (layerProps.DescField != "")
                    pmLine.Description = currentFeature.get_Value(currentFeature.Fields.FindField(layerProps.DescField)).ToString();

                if (layerProps.NameField != "")
                    pmLine.Name = currentFeature.get_Value(currentFeature.Fields.FindField(layerProps.NameField)).ToString();
                                

                //and add it to document
                pmLine.Geometry = line;
                folder.Features.Add(pmLine);

                #endregion
            }

        }

        private void processPointLayer(LayerProperties layerProps, geFolder folder)
        {
            //I have a point layer, and a kml folder :)
            IFeatureClass clasa = (layerProps.Layeru as IFeatureLayer).FeatureClass;
            //get acces to features
            IFeatureCursor featurele = clasa.Search(null, true);

            //if I have any features
            Point punct;
            IFeature currentFeature;

            //create coord system WGS 84 (Google earth)
            IGeographicCoordinateSystem gcs;
            SpatialReferenceEnvironment sre = new SpatialReferenceEnvironment();
            gcs = sre.CreateGeographicCoordinateSystem((int)esriSRGeoCSType.esriSRGeoCS_WGS1984);
            //create spatial reference
            ISpatialReference pointToSpatialReference;
            pointToSpatialReference = gcs;

            #region add points
            while ((currentFeature = featurele.NextFeature()) != null)
            {
                punct = currentFeature.Shape as Point;
                //coordinates and vertices
                double coordLat;
                double coordLong;


                //create a placemark for the point
                gePlacemark pmPoint = new gePlacemark();
                pmPoint.StyleUrl = "#Shape2KMLGeneratedStyle";

                //project point and get coordinates
                punct.Project(pointToSpatialReference);
                punct.QueryCoords(out coordLong, out coordLat);
                //add point     

                geCoordinates coords;
                gePoint point;

                switch (layerProps.AltitudeMode)
                {
                    case AltitudeMode.absolute:
                       
                        if (layerProps.Field == "")
                        {
                            //add point                  
                            coords = new geCoordinates(new geAngle90(coordLat), new geAngle180(coordLong), layerProps.Altitude);
                        }
                        else
                        {
                            int altitude;
                            //if altitude is integer, this should work   
                            altitude = (int)currentFeature.get_Value(currentFeature.Fields.FindField(layerProps.Field));

                            //add point                 
                            coords = new geCoordinates(new geAngle90(coordLat), new geAngle180(coordLong), layerProps.Multiplier * altitude);

                        }
                        point = new gePoint(coords);
                        point.AltitudeMode = geAltitudeModeEnum.absolute;
                        break;
                    case AltitudeMode.clampToGround:
                        coords = new geCoordinates(new geAngle90(coordLat), new geAngle180(coordLong));
                        point = new gePoint(coords);
                        point.AltitudeMode = geAltitudeModeEnum.clampToGround;
                        break;
                    case AltitudeMode.relativeToGround:
                        if (layerProps.Field == "")
                        {
                            //add point                  
                            coords = new geCoordinates(new geAngle90(coordLat), new geAngle180(coordLong), layerProps.Altitude);
                        }
                        else
                        {
                            int altitude;
                            //if altitude is integer, this should work   
                            altitude = (int)currentFeature.get_Value(currentFeature.Fields.FindField(layerProps.Field));

                            //add point                 
                            coords = new geCoordinates(new geAngle90(coordLat), new geAngle180(coordLong), layerProps.Multiplier * altitude);

                        }
                        point = new gePoint(coords);
                        point.AltitudeMode = geAltitudeModeEnum.relativeToGround;
                        break;
                    default:
                        point = null;
                        break;
                }

                if (layerProps.DescField != "")
                    pmPoint.Description = currentFeature.get_Value(currentFeature.Fields.FindField(layerProps.DescField)).ToString();

                if (layerProps.NameField != "")
                    pmPoint.Name = currentFeature.get_Value(currentFeature.Fields.FindField(layerProps.NameField)).ToString();

                pmPoint.Geometry = point;
                folder.Features.Add(pmPoint);
            #endregion
            }
        }

        private void processPolygonLayer(LayerProperties layerProps, geFolder folder)
        {
            //I have a polygon layer, and a kml folder :)
            IFeatureClass clasa = (layerProps.Layeru as IFeatureLayer).FeatureClass;
            //get acces to features
            IFeatureCursor featurele = clasa.Search(null, true);
            int nrFeature = clasa.FeatureCount(null);

            //if I have any features
            Polygon poligon;
            IFeature currentFeature;

            while ((currentFeature = featurele.NextFeature()) != null)
            {
                poligon = currentFeature.Shape as Polygon;
                //coordinates and vertices
                double coordLat;
                double coordLong;
                IEnumVertex colection = poligon.EnumVertices;
                IPoint polyVertex;

                //create coord system WGS 84 (Google earth)
                IGeographicCoordinateSystem gcs;
                SpatialReferenceEnvironment sre = new SpatialReferenceEnvironment();
                gcs = sre.CreateGeographicCoordinateSystem((int)esriSRGeoCSType.esriSRGeoCS_WGS1984);
                //create spatial reference
                ISpatialReference pointToSpatialReference;
                pointToSpatialReference = gcs;

                #region add points to polygon
                //create a placemark for the line
                gePlacemark pmPolygon = new gePlacemark();
                pmPolygon.StyleUrl = "#Shape2KMLGeneratedStyle";

                List<geCoordinates> polyCoords = new List<geCoordinates>();
                int index1, index2;
                //iterate points...

              

                while (!colection.IsLastInPart())
                {
                    //create polygon from vertices
                    colection.Next(out polyVertex, out index1, out index2);
                    //project point and get coordinates
                    polyVertex.Project(pointToSpatialReference);
                    polyVertex.QueryCoords(out coordLong, out coordLat);
                    //add point to line                  

                    try
                    {
                        //create points for polygon based on altitude mode.
                        switch (layerProps.AltitudeMode)
                        {
                            case AltitudeMode.absolute:
                                if (layerProps.Field == "")
                                {
                                    //add point to line                  
                                     polyCoords.Add(new geCoordinates(new geAngle90(coordLat), new geAngle180(coordLong),layerProps.Altitude));
                                }
                                else
                                {
                                    int altitude;
                                    //if altitude is integer, this should work   
                                    altitude = (int)currentFeature.get_Value(currentFeature.Fields.FindField(layerProps.Field));

                                    //add point to line                  
                                    polyCoords.Add(new geCoordinates(new geAngle90(coordLat), new geAngle180(coordLong),layerProps.Multiplier * altitude));
                                }
                                break;

                            case AltitudeMode.clampToGround:
                                //add point to line                  
                                polyCoords.Add(new geCoordinates(new geAngle90(coordLat), new geAngle180(coordLong)));
                                break;

                            case AltitudeMode.relativeToGround:
                                if (layerProps.Field == "")
                                {
                                    //add point to line                  
                                    polyCoords.Add(new geCoordinates(new geAngle90(coordLat), new geAngle180(coordLong), layerProps.Altitude));
                                }
                                else
                                {
                                    float altitude;
                                    //if altitude is integer, this should work   
                                    altitude = (float)currentFeature.get_Value(currentFeature.Fields.FindField(layerProps.Field));

                                    //add point to line                  
                                    polyCoords.Add(new geCoordinates(new geAngle90(coordLat), new geAngle180(coordLong), layerProps.Multiplier * altitude));
                               
                                }

                                break;

                            default:
                                break;
                        }
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Altitude field is not a number value");
                        break;
                    }

                }

                //create line from list of coords
                geOuterBoundaryIs outer = new geOuterBoundaryIs(new geLinearRing(polyCoords));
                gePolygon poly = new gePolygon(outer);
                //and add it to document

                switch (layerProps.AltitudeMode)
                {
                    //set altitude mode...
                    case AltitudeMode.absolute:
                        poly.AltitudeMode = geAltitudeModeEnum.absolute;
                        break;
                    case AltitudeMode.clampToGround:
                        poly.AltitudeMode = geAltitudeModeEnum.clampToGround;
                        break;
                    case AltitudeMode.relativeToGround:
                        poly.AltitudeMode = geAltitudeModeEnum.relativeToGround;
                        break;
                    default:
                        break;
                }
                
                if (layerProps.DescField != "")
                    pmPolygon.Description = currentFeature.get_Value(currentFeature.Fields.FindField(layerProps.DescField)).ToString();

                if (layerProps.NameField != "")
                    pmPolygon.Name = currentFeature.get_Value(currentFeature.Fields.FindField(layerProps.NameField)).ToString();

                pmPolygon.Geometry = poly;
                folder.Features.Add(pmPolygon);

                #endregion
            }

 

        }

        public void Generate(string path,ProgressBar pBar)
        {
            //generates document at path. shows progress using progressbar pBar

            this.Path = path;   
            kmlDocument.Name = System.IO.Path.GetFileName(path);;
            kmlDocument.StyleSelectors.Add(defaultStyle);

            //create a step in the progressbar for each layer
            int nrLayers = kmlLayers.Count;
            pBar.Minimum = 0;
            pBar.Maximum = nrLayers+1;
            pBar.Step = 1;
            //should generate a kml document at path...
            //first, iterate layers. Each layer makes a folder
            foreach (LayerProperties lp in kmlLayers)
            {
                //create a folder for the layer
                geFolder folder = new geFolder();
                folder.Name = lp.name;
                folder.Open = true;
                folder.Description = lp.Layeru.Name;

                processLayerProperties(lp,folder);

                kmlDocument.Features.Add(folder);
                pBar.PerformStep();
            }

            //no more layers, get document :)
            geKML kml = new geKML(kmlDocument);

            File.WriteAllBytes(path, kml.ToKML());
            pBar.PerformStep();
            //ProcessStartInfo psInfNotepad = new ProcessStartInfo("notepad.exe", path);
            //Process.Start(psInfNotepad);

        }
        

        public void GenerateAndOpen(string path)
        {
        }


    }
}

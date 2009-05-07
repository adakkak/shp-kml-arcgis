using System;
using System.Drawing;
using System.Runtime.InteropServices;
using ESRI.ArcGIS.ADF.BaseClasses;
using ESRI.ArcGIS.ADF.CATIDs;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.ArcMapUI;
using System.Windows.Forms;

namespace Shape2KML
{
    /// <summary>
    /// Summary description for ShapeToKML.
    /// </summary>
    [Guid("6753169b-83ab-4a5e-aefd-dd1b07f50172")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("Shape2KML.ShapeToKML")]
    public sealed class ShapeToKML : BaseCommand
    {
        #region COM Registration Function(s)
        [ComRegisterFunction()]
        [ComVisible(false)]
        static void RegisterFunction(Type registerType)
        {
            // Required for ArcGIS Component Category Registrar support
            ArcGISCategoryRegistration(registerType);

            //
            // TODO: Add any COM registration code here
            //
        }

        [ComUnregisterFunction()]
        [ComVisible(false)]
        static void UnregisterFunction(Type registerType)
        {
            // Required for ArcGIS Component Category Registrar support
            ArcGISCategoryUnregistration(registerType);

            //
            // TODO: Add any COM unregistration code here
            //
        }

        #region ArcGIS Component Category Registrar generated code
        /// <summary>
        /// Required method for ArcGIS Component Category registration -
        /// Do not modify the contents of this method with the code editor.
        /// </summary>
        private static void ArcGISCategoryRegistration(Type registerType)
        {
            string regKey = string.Format("HKEY_CLASSES_ROOT\\CLSID\\{{{0}}}", registerType.GUID);
            MxCommands.Register(regKey);

        }
        /// <summary>
        /// Required method for ArcGIS Component Category unregistration -
        /// Do not modify the contents of this method with the code editor.
        /// </summary>
        private static void ArcGISCategoryUnregistration(Type registerType)
        {
            string regKey = string.Format("HKEY_CLASSES_ROOT\\CLSID\\{{{0}}}", registerType.GUID);
            MxCommands.Unregister(regKey);

        }

        #endregion
        #endregion

        private IApplication m_application;

        private static Form frmMain;

        public static Form FrmMain
        {
            get { return ShapeToKML.frmMain; }
            set { ShapeToKML.frmMain = value; }
        }

        public ShapeToKML()
        {
            //
            // TODO: Define values for the public properties
            //
            base.m_category = "shape_to_kml";
            base.m_caption = "Shape_2_kml";
            base.m_message = "Convert_Shapefiles to KML format";
            base.m_name = "Shape_2_kml";
            base.m_toolTip = "Convert_Shapefile_to_kml";

            try
            {
                //
                // TODO: change bitmap name if necessary
                //
                string bitmapResourceName = GetType().Name + ".bmp";
                base.m_bitmap = new Bitmap(GetType(), bitmapResourceName);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap");
            }
        }

        #region Overriden Class Methods

        /// <summary>
        /// Occurs when this command is created
        /// </summary>
        /// <param name="hook">Instance of the application</param>
        public override void OnCreate(object hook)
        {
            //get a reference to the app.
            if (hook == null)
                return;

            m_application = hook as IApplication;

            //Disable if it is not ArcMap
            if (hook is IApplication)
                base.m_enabled = true;
            else
                base.m_enabled = false;

            // TODO:  Add other initialization code
        }

        /// <summary>
        /// Occurs when this command is clicked
        /// </summary>
        public override void OnClick()
        {
            // TODO: Add Command1.OnClick implementation
            // MessageBox.Show("dadadada");
            if (FrmMain== null)
            FrmMain = new ShapeToKMLForm(m_application as IApplication);
            //f.TopLevel = true;
            FrmMain.TopMost = true;
            FrmMain.Show();
        }

        #endregion
    }

   }

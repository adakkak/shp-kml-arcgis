using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using ESRI.ArcGIS.ADF.BaseClasses;
using ESRI.ArcGIS.ADF.CATIDs;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Carto;
using System.IO;
using Google.KML;

namespace Shape2KML
{
    public partial class ShapeToKMLForm : Form
    {
        IApplication m_application;
        List<LayerProperties> layers;

        public ShapeToKMLForm(IApplication application)
        {
            layers = new List<LayerProperties>();
            InitializeComponent();
            this.m_application = application;

            chkLayers.Items.Clear();
            IMxDocument activeDoc = application.Document as IMxDocument;

            IMap activeMap = activeDoc.FocusMap;
            layers.Clear();

            for (int i = 0; i < activeMap.LayerCount; i++)
            {
                //if it's a raster or a feature layer. Other supported layers will be added later
                if ((activeMap.get_Layer(i) is IFeatureLayer) || (activeMap.get_Layer(i) is IRasterLayer))
                {
                    //chkLayers.Items.Add(activeMap.get_Layer(i).Name, false);
                    LayerProperties lp = new LayerProperties();
                    lp.Layeru = activeMap.get_Layer(i);
                    lp.Tesselate = false;
                    lp.AltitudeMode = AltitudeMode.clampToGround;
                    lp.Extrude = false;
                    lp.Color = Color.Black;
                    lp.name = lp.Layeru.Name;
                    chkLayers.Items.Add(lp, false);
                    
                }
            }
            
            
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            //choose file
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Title = "Choose export file";
            sfd.Filter = "kml files (*.kml)|*.kml|All files (*.*)|*.*";
            sfd.FileName = "Export.kml";
            sfd.InitialDirectory = "C:\\";
            sfd.AddExtension = true;
            sfd.DefaultExt = ".kml";

            //find file that was chosen
            DialogResult result = sfd.ShowDialog();

            if (result == DialogResult.OK)
            {
                //ok clicked...
                String fileName = sfd.FileName;
                txtFileName.Text = fileName;

            }
        }

        private void chkOnTop_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = chkOnTop.Checked;
        }

        private void chkLayers_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // MessageBox.Show("check");
            //check which item changed....
            if ((chkLayers.GetItemChecked(e.Index)))
            {
                //an item was unchecked
                //deselect select all
                chkSelectAll.Checked = false;
                //remove unselected layer from combo
                for (int j = 0; j < cmbLayers.Items.Count; j++)
                {
                    if (cmbLayers.Items[j] == chkLayers.Items[e.Index])
                    {
                        //also remove from layers
                        layers.Remove(chkLayers.Items[e.Index] as LayerProperties);
                        cmbLayers.Items.RemoveAt(j);
                        //add style menu if needed.
                        checkLayers();
                    }

                }
            }
            else
            {
                //an item was checked.
                //iterate cmbLayers
                for (int j = 0; j < cmbLayers.Items.Count; j++)
                {
                    //don't readd it
                    if (cmbLayers.Items[j] == chkLayers.Items[e.Index])
                        return;
                }
                //add layer 
                cmbLayers.Items.Add(chkLayers.Items[e.Index]);
                //add style menu if needed.
                checkLayers();

                //and add it to layers list
                if (!layers.Contains(chkLayers.Items[e.Index] as LayerProperties))
                {
                    layers.Add(chkLayers.Items[e.Index] as LayerProperties);
                }
            }

        }

        private void checkLayers()
        {
            grpLine.Visible = false;
            grpIcon.Visible = false;
            grpPoly.Visible = false;
            //check what kind of layers are selected and activate corresponding style groups.
            for (int i = 0; i < cmbLayers.Items.Count; i++)
            {               
                ILayer layerul = (cmbLayers.Items[i] as LayerProperties).Layeru ;
                if (layerul is IFeatureLayer)
                {
                    if ((layerul as IFeatureLayer).FeatureClass.ShapeType == ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline)
                        grpLine.Visible = true;
                    if ((layerul as IFeatureLayer).FeatureClass.ShapeType == ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint)
                        grpIcon.Visible = true;
                    if ((layerul as IFeatureLayer).FeatureClass.ShapeType == ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon)
                        grpPoly.Visible = true;
                }                   

            }
        }

        private void chkSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            //when checked, select all layers.
            if (this.chkSelectAll.Checked)
            {
                //if 'select all' checked, change all to selected.
                for (int i = 0; i < chkLayers.Items.Count; i++)
                {
                    chkLayers.SetItemChecked(i, true);
                }
                // this.chkSelectAll.Checked = false;
            }

        }

        private void cmbLayers_SelectedIndexChanged(object sender, EventArgs e)
        {
            //get selected layer
            LayerProperties selectedLayer =  ((cmbLayers.Items[cmbLayers.SelectedIndex]) as LayerProperties);
            
            //if it's a known type, enable altitude group to be able to select altitude
            if (selectedLayer.Layeru is IRasterLayer || selectedLayer.Layeru is IFeatureLayer)
                grpAltitude.Enabled = true;

            //if it's Raster altitude cannot be relative.
            if (selectedLayer.Layeru is IRasterLayer)
            {
                rbRelative.Enabled = false;
            }
            else
            {
                rbRelative.Enabled = true;
                //get fields and add them to combobox.
                IFeatureLayer sel = selectedLayer.Layeru as IFeatureLayer;
                cmbAttribute.Items.Clear();
                cmbAttribute.Text = "";
                cmbDesc.Items.Clear();
                cmbDesc.Text = "";
                cmbName.Items.Clear();
                cmbName.Text = "";
                for (int i = 0; i < sel.FeatureClass.Fields.FieldCount; i++)
                {  
                    //get name for every field.
                    string fieldName = sel.FeatureClass.Fields.get_Field(i).Name;
                    //and add it to the combobox. Keep the selected one correct if recorded.
                    cmbAttribute.Items.Add(fieldName);
                    cmbDesc.Items.Add(fieldName);
                    cmbName.Items.Add(fieldName);
                    //field name is unique for a table, so this is ok.
                    if (selectedLayer.Field == fieldName)
                        cmbAttribute.SelectedIndex = i;
                    if (selectedLayer.DescField == fieldName)
                        cmbDesc.SelectedIndex = i;
                    if (selectedLayer.NameField == fieldName)
                        cmbName.SelectedIndex = i;
                }
            }
            
            //check correct button from the properties.
            switch (selectedLayer.AltitudeMode)
            {
                case AltitudeMode.absolute:
                    rbAbsolute.Checked=true;
                    numMultiplier.Value = selectedLayer.Multiplier;
                    numAltitude.Value = selectedLayer.Altitude;
                    break;
                case AltitudeMode.clampToGround:
                    grpSelAltitude.Visible = false;
                    rbClamp.Checked = true;
                    break;
                case AltitudeMode.relativeToGround:
                    rbRelative.Checked = true;
                    numAltitude.Value = selectedLayer.Altitude;
                    break;
                default:
                    break;
            }

        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (rbAbsolute.Checked)
            {
                //if its checked enable list of attributes to choose from:
                grpSelAltitude.Visible = true;
                rbSelAttrib.Checked = true;
                LayerProperties selected = cmbLayers.SelectedItem as LayerProperties;
                selected.AltitudeMode = AltitudeMode.absolute;
                numAltitude.Value = selected.Altitude;
                numMultiplier.Value = selected.Multiplier;
                chkExtrude.Visible = true;
                chkTesselate.Visible = false;
               
            }
            else
            {
                //if its unchecked disable list of attributes to choose from:
                grpSelAltitude.Visible = false;
            }
        }

        private void rbSelManualAltitude_CheckedChanged(object sender, EventArgs e)
        {
            //enable or disable correspondent controls
            numAltitude.Enabled = rbSelManualAltitude.Checked;
            numMultiplier.Enabled = rbSelAttrib.Checked;
            cmbAttribute.Enabled = rbSelAttrib.Checked;
            lblX.Enabled = rbSelAttrib.Checked;
            if (rbSelManualAltitude.Checked)
            {
                LayerProperties selected = cmbLayers.SelectedItem as LayerProperties;
                selected.Field = "";
                selected.Altitude = (int)numAltitude.Value;
            }
        }

        private void rbSelAttrib_CheckedChanged(object sender, EventArgs e)
        {
            //enable or disable correspondent controls
            numMultiplier.Enabled = rbSelAttrib.Checked;
            cmbAttribute.Enabled = rbSelAttrib.Checked;
            lblX.Enabled = rbSelAttrib.Checked;
            numAltitude.Enabled = rbSelManualAltitude.Checked;
            if (rbSelAttrib.Checked)
            {
                LayerProperties selected = cmbLayers.SelectedItem as LayerProperties;
                if (cmbAttribute.SelectedItem != null)
                    selected.Field = cmbAttribute.SelectedItem.ToString();
                selected.Altitude = 0;
            }
         }

        private void rbRelative_CheckedChanged(object sender, EventArgs e)
        {
            if (rbRelative.Checked)
            {
                LayerProperties selected = cmbLayers.SelectedItem as LayerProperties;
                selected.AltitudeMode = AltitudeMode.relativeToGround;
                grpSelAltitude.Visible = true;
                chkExtrude.Visible = true;
                chkTesselate.Visible = false;
            }
        }

        private void rbClamp_CheckedChanged(object sender, EventArgs e)
        {
            if (rbClamp.Checked)
            {
                LayerProperties selected = cmbLayers.SelectedItem as LayerProperties;
                selected.AltitudeMode = AltitudeMode.clampToGround;
                chkTesselate.Visible = true;
                chkExtrude.Visible = false;
            }
            else
                grpSelAltitude.Visible = false;
        }

        private void numMultiplier_ValueChanged(object sender, EventArgs e)
        {
            LayerProperties selected = cmbLayers.SelectedItem as LayerProperties;
            selected.Multiplier = (int)numMultiplier.Value;
        }

        private void cmbAttribute_SelectedIndexChanged(object sender, EventArgs e)
        {
            LayerProperties selected = cmbLayers.SelectedItem as LayerProperties;
            selected.Field = cmbAttribute.SelectedItem as string;
        }

        private void btnLineColor_Click(object sender, EventArgs e)
        {
            ColorDialog cd = new ColorDialog();
            if (cd.ShowDialog() == DialogResult.OK)
            {
                pbLine.BackColor = cd.Color;
            }
        }

        private void btnFillColor_Click(object sender, EventArgs e)
        {
            ColorDialog cd = new ColorDialog();
            if (cd.ShowDialog() == DialogResult.OK)
            {
                pbFill.BackColor = cd.Color;
            }
        }

        private void btnIconColor_Click(object sender, EventArgs e)
        {
            ColorDialog cd = new ColorDialog();
            if (cd.ShowDialog() == DialogResult.OK)
            {
                pbIcon.BackColor = cd.Color;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (sender is PictureBox)
                pbSelectedIcon.ImageLocation = (sender as PictureBox).ImageLocation;
        }

        private void btnNoIcon_Click(object sender, EventArgs e)
        {
            pbSelectedIcon.ImageLocation = "";
        }

        #region "Get Document Path"
        // ArcGIS Snippet Title: 
        // Get Document Path
        //
        // Add the following references to the project:
        // ESRI.ArcGIS.Framework
        // 
        // Intended ArcGIS Products for this snippet:
        // ArcGIS Desktop
        //
        // Required ArcGIS Extensions:
        // (NONE)
        //
        // Notes:
        // This snippet is intended to be inserted at the base level of a Class.
        // It is not intended to be nested within an existing Method.
        //
        // Use the following XML documentation comments to use this snippet:
        /// <summary>Get the active document path for the ArcGIS application.</summary>
        ///
        /// <param name="application">An IApplication interface.</param>
        /// 
        /// <returns>A System.String is returned that is the path for the document.</returns>
        ///
        /// <remarks></remarks>
        public System.String GetActiveDocumentPath(ESRI.ArcGIS.Framework.IApplication application)
        {
            ESRI.ArcGIS.Framework.ITemplates templates = application.Templates;
            return templates.get_Item(templates.Count - 1);
        }

        #endregion    
                
        
        private void btnCreate_Click(object sender, EventArgs e)
        {
         
            //creates kml file if selected.
            String path = txtFileName.Text;

            if ((path != null) &&(path != "" ))
            if (Directory.Exists(Path.GetDirectoryName(path)))
                //diretorul exista. Si acum fisierul are extensie kml?
                if (Path.GetFileName(path).EndsWith(".kml"))
                       //si e fisier kml. BUn...
                {   
                    KMLGenerator generator = new KMLGenerator(layers);
                    //create style in generator
                    #region Create kml style
                    
                    geStyle style = new geStyle("Shape2KMLGeneratedStyle");
                    //style.ID = "Shape2KMLGeneratedStyle";

                    //should use booleans for these, but for now checking the controls is ok...

                    if (grpIcon.Visible)
                    {
                        //icon style
                        style.IconStyle = new geIconStyle();
                        style.IconStyle.Color.SysColor = pbIcon.BackColor;
                        //style.IconStyle.ColorMode =
                        style.IconStyle.Icon = new geIcon(pbSelectedIcon.ImageLocation);
                    }

                    if (grpLine.Visible)
                    {
                        //line style, if we have lines
                        style.LineStyle = new geLineStyle();
                        style.LineStyle.Color.SysColor = pbLine.BackColor;
                        style.LineStyle.Width = (float)numLineWidth.Value;
                    }

                    //polygon style, only if we have polygons
                    if (grpPoly.Visible)
                    {
                        style.PolyStyle = new gePolyStyle();
                        style.PolyStyle.Color.SysColor = pbFill.BackColor;
                        
                    }

                    generator.DefaultStyle = style;
                    #endregion

                    pnlMain.Visible = false;
                    pnlProgress.Visible = true;
                    //IMxDocument activeDoc = m_application.Document as IMxDocument;
                    //MessageBox.Show (activeDoc.ToString);

                    //get active document name
                    String docPath = GetActiveDocumentPath(m_application);
                    generator.KmlDocument.Name = Path.GetFileName(docPath);
                    generator.KmlDocument.Description = docPath;
                    generator.Generate(path,pbarDone);

                    MessageBox.Show("Creat document "+ path);
                    pnlMain.Visible = true;
                    pnlProgress.Visible = false;
                   
                }
           // MessageBox.Show("not ok");
        }

        private void chkTesselate_CheckedChanged(object sender, EventArgs e)
        {
            LayerProperties selected = cmbLayers.SelectedItem as LayerProperties;
            selected.Tesselate = chkTesselate.Checked;
        }

        private void numAltitude_ValueChanged(object sender, EventArgs e)
        {
            LayerProperties selected = cmbLayers.SelectedItem as LayerProperties;
            selected.Altitude = (int)numAltitude.Value;
        }

        private void cmbDesc_SelectedIndexChanged(object sender, EventArgs e)
        {
            LayerProperties selected = cmbLayers.SelectedItem as LayerProperties;
            if (cmbDesc.SelectedItem != null)
                selected.DescField = cmbDesc.SelectedItem as string;
        }

        private void chkExtrude_CheckedChanged(object sender, EventArgs e)
        {
              LayerProperties selected = cmbLayers.SelectedItem as LayerProperties;
              selected.Extrude = chkExtrude.Checked;
        }

        private void cmbName_SelectedIndexChanged(object sender, EventArgs e)
        {
            LayerProperties selected = cmbLayers.SelectedItem as LayerProperties;
            if (cmbName.SelectedItem != null)
                selected.NameField = cmbName.SelectedItem as string;

        }

        private void ShapeToKMLForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            ShapeToKML.FrmMain = null;
        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://rotrekking.ro/Shape2KMLHelp/index.html");
        }

    }

    }

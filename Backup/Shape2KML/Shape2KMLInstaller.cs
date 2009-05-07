using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using ESRI.ArcGIS.ADF.BaseClasses;
using ESRI.ArcGIS.ADF.CATIDs;
using System.Runtime.InteropServices;
using System.Collections;

namespace Shape2KML
{
    [RunInstaller(true)]
    public partial class Shape2KMLInstaller : Installer
    {
        public Shape2KMLInstaller()
        {
            InitializeComponent();
        }

        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);
            RegistrationServices regSrv = new RegistrationServices();
            regSrv.RegisterAssembly(base.GetType().Assembly, AssemblyRegistrationFlags.SetCodeBase);
        }

        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);
            RegistrationServices regSrv = new RegistrationServices();
            regSrv.UnregisterAssembly(base.GetType().Assembly);
        }


    }
}
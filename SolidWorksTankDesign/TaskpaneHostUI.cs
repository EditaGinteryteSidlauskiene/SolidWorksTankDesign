using AddinWithTaskpane;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;
using WarningAndErrorService;
using Attribute = SolidWorks.Interop.sldworks.Attribute;

namespace SolidWorksTankDesign
{
    //Setting prog id the string in variable SWTASKPANE_PROGID, and then UserControl (below) gets the prog id.
    // User control will be injected into the SolidWorks by passing this id.
    [ProgId(TaskpaneIntegration.SWTASKPANE_PROGID)]
    public partial class TaskpaneHostUI : UserControl
    {
        private WarningService warningService = new WarningService();

        // Getting SW application reference.
        SldWorks swApp;
        public void AssignSWApp(SldWorks inputswApp)
        {
            swApp = inputswApp;
        }

        public TaskpaneHostUI()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            ModelDoc2 tankSiteModelDoc = swApp.IActiveDoc2;

            TankSiteAssembly tankSiteAssembly = new TankSiteAssembly(warningService, swApp, tankSiteModelDoc);
        }
    }
}

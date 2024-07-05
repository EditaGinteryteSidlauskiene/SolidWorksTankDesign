using AddinWithTaskpane;
using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace SolidWorksTankDesign
{
    internal class DishedEnd
    {
        [JsonIgnore]
        SldWorks _solidWorksApplication;

        [JsonIgnore]
        private ModelDoc2 _assemblyOfDishedEndsDoc;
        // atributai ir privatus//
        // [ ] 
        public DishedEndSettings _dishedEndSettings;

        public DishedEnd(SldWorks solidWorksApplication, ModelDoc2 assemblyOfDishedEndsDoc, Component2 dishedEnd)
        {
            _solidWorksApplication = solidWorksApplication;
            _assemblyOfDishedEndsDoc = assemblyOfDishedEndsDoc;
            _dishedEndSettings = new DishedEndSettings();

        }

        public Feature CenterAxis
        {
            get
            {
                using (var document = new SolidWorksDocumentWrapper(_solidWorksApplication, _assemblyOfDishedEndsDoc))
                {
                    return (Feature)_assemblyOfDishedEndsDoc.Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDCenterAxis,
                        out int error);
                }
            }
        }

        public Feature PositionPlane
        {
            get
            {
                using (var document = new SolidWorksDocumentWrapper(_solidWorksApplication, _assemblyOfDishedEndsDoc))
                {
                    return (Feature)_assemblyOfDishedEndsDoc.Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDPositionPlane,
                        out int error);
                }
            }
        }

        public Component2 Component
        {
            get
            {
                using (var document = new SolidWorksDocumentWrapper(_solidWorksApplication, _assemblyOfDishedEndsDoc))
                {
                    return (Component2)_assemblyOfDishedEndsDoc.Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDComponent,
                        out int error);
                }
            }
        }

        public Feature CenterAxisMate
        {
            get
            {
                using (var document = new SolidWorksDocumentWrapper(_solidWorksApplication, _assemblyOfDishedEndsDoc))
                {
                    return (Feature)_assemblyOfDishedEndsDoc.Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDCenterAxisMate,
                        out int error);
                }
            }
        }

        public Feature RightPlaneMate
        {
            get
            {
                using (var document = new SolidWorksDocumentWrapper(_solidWorksApplication, _assemblyOfDishedEndsDoc))
                {
                    return (Feature)_assemblyOfDishedEndsDoc.Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDRightPlaneMate,
                        out int error);
                }
            }
        }

        public Feature FrontPlaneMate
        {
            get
            {
                using (var document = new SolidWorksDocumentWrapper(_solidWorksApplication, _assemblyOfDishedEndsDoc))
                {
                    return (Feature)_assemblyOfDishedEndsDoc.Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDFrontPlaneMate,
                        out int error);
                }
            }
        }
    }
}

using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.Helpers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WarningAndErrorService;

namespace SolidWorksTankDesign
{
    internal class InnerDishedEnd : DishedEnd
    {
        public DishedEndSettings _dishedEndSettings;

        private const string INNER_DISHED_END_NAME = "Inner end";
        private const string POSITION_PLANE_NAME = "Position plane";
        private const string DISHED_END_ECOMPONENT_PATH = "C:\\Users\\Edita\\Desktop\\Parts\\Inner dished end.SLDPRT";
        private const string RIGHT_PLANE_NAME = "Right plane";
        private const string FRONT_PLANE_NAME = "Front plane";
        private const string CENTER_AXIS_NAME = "Center axis";

        /// <summary>
        /// This constructor is called when a new inner dished end is added.
        /// </summary>
        /// <param name="assemblyOfDishedEndsCenterAxis"></param>
        /// <param name="assemblyOfDishedEndsFrontPlane"></param>
        /// <param name="referenceDishedEnd"></param>
        /// <param name="dishedEndAlignment"></param>
        /// <param name="distance"></param>
        /// <param name="compartmentNumber"></param>
        public InnerDishedEnd(
            Feature assemblyOfDishedEndsCenterAxis,
            Feature assemblyOfDishedEndsFrontPlane,
            DishedEnd referenceDishedEnd,
            DishedEndAlignment dishedEndAlignment,
            double distance,
            int compartmentNumber)
            : base()
        {
            AddInnerEnd(assemblyOfDishedEndsCenterAxis, assemblyOfDishedEndsFrontPlane, referenceDishedEnd, dishedEndAlignment, distance, compartmentNumber);
        }

        /// <summary>
        /// Creates and positions an inner dished end within a SolidWorks assembly
        /// </summary>
        /// <param name="assemblyOfDishedEndsCenterAxis"></param>
        /// <param name="assemblyOfDishedEndsFrontPlane"></param>
        /// <param name="referenceDishedEnd"></param>
        /// <param name="dishedEndAlignment"></param>
        /// <param name="distance"></param>
        /// <param name="compartmentNumber"></param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        private void AddInnerEnd(
            Feature assemblyOfDishedEndsCenterAxis,
            Feature assemblyOfDishedEndsFrontPlane,
            DishedEnd referenceDishedEnd,
            DishedEndAlignment dishedEndAlignment,
            double distance,
            int compartmentNumber)
        {
            // 1. Input Validation
            if (assemblyOfDishedEndsCenterAxis == null)
                throw new ArgumentNullException(nameof(assemblyOfDishedEndsCenterAxis));

            if (assemblyOfDishedEndsFrontPlane == null)
                throw new ArgumentNullException(nameof(assemblyOfDishedEndsFrontPlane));

            if (referenceDishedEnd == null)
                throw new ArgumentNullException(nameof(referenceDishedEnd));

            ModelDoc2 assemblyOfDishedEndsDoc = SolidWorksDocumentProvider.ActiveDoc();
            if (assemblyOfDishedEndsDoc == null)
                throw new InvalidOperationException("Active SolidWorks document not found.");

            string positionPlaneName = $"{INNER_DISHED_END_NAME} {compartmentNumber} {POSITION_PLANE_NAME}";
            string componentName = $"{INNER_DISHED_END_NAME} {compartmentNumber}";

            // 2. Create Position Plane
            Feature positionPlane = FeatureManager.CreateReferencePlaneWithDistance(
                modelDocument: assemblyOfDishedEndsDoc, 
                existingPlane: referenceDishedEnd.PositionPlane(), 
                distance: distance, 
                name: positionPlaneName);

            // 3. Add and Make Independent Dished End Component
            Component2 dishedEnd = ComponentManager.AddComponentPart(
                solidWorksApplication: SolidWorksDocumentProvider._solidWorksApplication,
                assemblyDocument: assemblyOfDishedEndsDoc, 
                componentPath: DISHED_END_ECOMPONENT_PATH);

            ComponentManager.MakeComponentPartIndependent(assemblyOfDishedEndsDoc, dishedEnd, DISHED_END_ECOMPONENT_PATH);

            // 4. Rename the Component
            FeatureManager.GetFeatureByName(assemblyOfDishedEndsDoc, dishedEnd.Name2).Name = componentName;

            // 5. Get Features for Mating
            Feature innerDishedEndCenterAxis = FeatureManager.GetFeatureByName(dishedEnd, "Center axis");

            // 6. Create Mates
            Feature rightPlaneMate = MateManager.CreateMate(
                componentFeature1: positionPlane,
                componentFeature2: FeatureManager.GetMajorPlane(dishedEnd, MajorPlane.Right),
                alignmentType: MateAlignment.Aligned,
                name: $"{dishedEnd.Name2} - {RIGHT_PLANE_NAME}");

            Feature frontPlaneMate = MateManager.CreateMate(
                componentFeature1: assemblyOfDishedEndsFrontPlane,
                componentFeature2: FeatureManager.GetMajorPlane(dishedEnd, MajorPlane.Front),
                alignmentType: MateAlignment.Aligned,
                name: $"{dishedEnd.Name2} - {FRONT_PLANE_NAME}");

            Feature centerAxisMate = MateManager.CreateMate(
                componentFeature1: assemblyOfDishedEndsCenterAxis,
                componentFeature2: innerDishedEndCenterAxis,
                alignmentType: MateAlignment.Anti_Aligned,
                name: $"{dishedEnd.Name2} - {CENTER_AXIS_NAME}");

            // 7. Get Inner Dished End Entities and Initialize Settings
            _dishedEndSettings = new DishedEndSettings();
            try
            {
                GetInnerDishedEndEntities();
            }
            catch (Exception ex)
            {
                // Display a user-friendly error message
                MessageBox.Show($"Error getting inner dished end entities: {ex.Message}");
            }

            // 8. Helper Method to Get Inner Dished End Entities
            void GetInnerDishedEndEntities()
            {
                try
                {
                    // Populate the _dishedEndSettings with the retrieved PIDs
                    _dishedEndSettings.PIDPositionPlane = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(positionPlane);
                    _dishedEndSettings.PIDComponent = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(dishedEnd);
                    _dishedEndSettings.PIDCenterAxis = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(innerDishedEndCenterAxis);
                    _dishedEndSettings.PIDCenterAxisMate = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(centerAxisMate);
                    _dishedEndSettings.PIDRightPlaneMate = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(rightPlaneMate);
                    _dishedEndSettings.PIDFrontPlaneMate = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(frontPlaneMate);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "The attribute could not be created.");
                    return;
                }
            }
        }
    }
}

using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.Helpers;
using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.TankSiteConfigurations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SolidWorksTankDesign.MVP.Services
{
    internal class NozzleSolidWorksService : INozzleSolidWorksService
    {
        private string _nozzlePositionSketchPath = "C:\\Users\\Edita\\TankDesignStudio\\TankSite\\Manholes\\Nozzle position sketch.SLDASM";
        private string nozzleDocPath = "C:\\Users\\Edita\\TankDesignStudio\\TankSite\\Manholes\\Manhole DN600 Neck with flange.SLDASM";

        public void AddNozzle(Guid compartmentConfigId, NozzleReferenceType referenceType, double distance)
        {
            (Compartment compartment, CompartmentConfiguration config)? mapping = null;

            mapping = CompartmentsMappingHelper.GetMappingById(compartmentConfigId);

            Feature referencePlane = null;
            Compartment targetedCompartment = mapping.Value.compartment;

            switch (referenceType)
            {
                case NozzleReferenceType.LeftDishedEnd:
                    targetedCompartment.ActivateDocument();
                    referencePlane = targetedCompartment.GetLeftEndPlane();

                    if (referencePlane == null) return;

                    targetedCompartment.AddNozzle(
                        _nozzlePositionSketchPath,
                        nozzleDocPath,
                        0,
                        referencePlane,
                        distance,
                        false,
                        2500);
                    break;

                case NozzleReferenceType.RightDishedEnd:
                    targetedCompartment.ActivateDocument();
                    referencePlane = targetedCompartment.GetRightEndPlane();

                    targetedCompartment.AddNozzle(
                        _nozzlePositionSketchPath,
                        nozzleDocPath,
                        0,
                        referencePlane,
                        distance,
                        true,
                        2500);
                    break;
                case NozzleReferenceType.OtherNozzle:
                    targetedCompartment.ActivateDocument();
                    referencePlane = targetedCompartment.Nozzles.Last().GetPositionPlane();

                    targetedCompartment.AddNozzle(
                        _nozzlePositionSketchPath,
                        nozzleDocPath,
                        0,
                        referencePlane,
                        distance,
                        false,
                        2500);
                    break;
            }

            try
            {
                // Save changes after all modifications are complete.
                DocumentManager.UpdateAndSaveDocuments();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving documents: {ex.Message}");
            }
        }

        public void RepositionNozzle(
            bool isOffsetPositive,
            double distance,
            bool isRotationDirectionPositive,
            double angle)
        {
            List<Compartment> compartments = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments;

            if (distance != 0)
            {
                if (isOffsetPositive == true)
                {
                    compartments[0].ActivateDocument();

                    compartments[0].Nozzles.Last().SetOffset(distance);
                }

                else
                {
                    compartments[0].ActivateDocument();

                    compartments[0].Nozzles.Last().SetOffset(-distance);
                }
            }

            // Normalize angle to 0-360 degrees (using modulo operator)
            angle = (angle % 360 + 360) % 360;

            if (isRotationDirectionPositive == true)
            {
                compartments[0].ActivateDocument();

                compartments[0].Nozzles.Last().SetRotationAngle(angle);
            }

            else
            {
                compartments[0].ActivateDocument();

                compartments[0].Nozzles.Last().SetRotationAngle(360 - angle);
            }

            try
            {
                // Save changes after all modifications are complete.
                DocumentManager.UpdateAndSaveDocuments();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving documents: {ex.Message}");
            }
        }
    }
}

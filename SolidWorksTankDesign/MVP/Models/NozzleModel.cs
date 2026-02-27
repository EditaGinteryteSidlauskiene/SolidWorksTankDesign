using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.MVP.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace SolidWorksTankDesign.MVP.Presenters
{
    public class NozzleModel : INozzleModel
    {
        private string _nozzlePositionSketchPath = "C:\\Users\\Edita\\TankDesignStudio\\TankSite\\Manholes\\Nozzle position sketch.SLDASM";
        private string nozzleDocPath = "C:\\Users\\Edita\\TankDesignStudio\\TankSite\\Manholes\\Manhole DN600 Neck with flange.SLDASM";

        public NozzleModel() { }

        public void AddNozzle(string projectFolderPath, NozzleReferenceType referenceType, double distance)
        {
            List<Compartment> compartments = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments;

            Feature referencePlane = null;

            switch (referenceType)
            {
                case NozzleReferenceType.LeftDishedEnd:
                    compartments[0].ActivateDocument();
                    referencePlane = compartments[0].GetLeftEndPlane();

                    if (referencePlane == null) return;

                    compartments[0].AddNozzle(
                        projectFolderPath,
                        _nozzlePositionSketchPath,
                        nozzleDocPath,
                        0,
                        referencePlane,
                        distance,
                        false,
                        2500);
                    break;

                case NozzleReferenceType.RightDishedEnd:
                    compartments[0].ActivateDocument();
                    referencePlane = compartments[0].GetRightEndPlane();

                    compartments[0].AddNozzle(
                        projectFolderPath,
                        _nozzlePositionSketchPath,
                        nozzleDocPath,
                        0,
                        referencePlane,
                        distance,
                        true,
                        2500);
                    break;
                case NozzleReferenceType.OtherNozzle:
                    compartments[0].ActivateDocument();
                    referencePlane = compartments[0].Nozzles.Last().GetPositionPlane();

                    compartments[0].AddNozzle(
                        projectFolderPath,
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
            
            if(distance != 0)
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

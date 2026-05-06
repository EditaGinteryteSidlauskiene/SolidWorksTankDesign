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

        public void ApplyNozzleChanges()
        {
            try
            {
                List<Compartment> compartments = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments;
                List<CompartmentConfiguration> compartmentConfigs = SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations;

                bool hasAnyChanges = false;

                // Iterate through each compartment and its configurations
                for (int i = 0; i < compartments.Count && i < compartmentConfigs.Count; i++)
                {
                    Compartment compartment = compartments[i];
                    CompartmentConfiguration config = compartmentConfigs[i];

                    if (config.NozzleConfigurations == null)
                        continue;

                    // Check if there are any new nozzles or modified nozzles in this compartment
                    bool hasChangesInCompartment = HasNozzleChanges(compartment, config);

                    if (!hasChangesInCompartment)
                        continue;

                    hasAnyChanges = true;
                    compartment.ActivateDocument();

                    // Process each nozzle configuration
                    foreach (var nozzleConfig in config.NozzleConfigurations)
                    {
                        // Find the corresponding nozzle in the compartment by matching IDs
                        var existingNozzle = compartment.Nozzles?.FirstOrDefault(n => 
                            n._nozzleSettings != null && n._nozzleSettings.ID == nozzleConfig.Id);

                        if (existingNozzle != null)
                        {
                            ApplyExistingNozzlePositionChanges(compartment, existingNozzle, nozzleConfig);
                        }
                        else
                        {
                            // Nozzle doesn't exist - add new nozzle
                            AddNozzleFromConfiguration(compartment, config, nozzleConfig, i);
                        }
                    }

                    compartment.CloseDocument();
                }

                // Only save if there were actual changes
                if (hasAnyChanges)
                {
                    DocumentManager.UpdateAndSaveDocuments();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying nozzle changes: {ex.Message}", "Apply Nozzle Changes Error");
            }
        }

        private bool HasNozzleChanges(Compartment compartment, CompartmentConfiguration config)
        {
            if (config.NozzleConfigurations == null || config.NozzleConfigurations.Count == 0)
                return false;

            // Check if there are any new nozzles (configurations without matching nozzles in SolidWorks)
            foreach (var nozzleConfig in config.NozzleConfigurations)
            {
                var existingNozzle = compartment.Nozzles?.FirstOrDefault(n =>
                    n._nozzleSettings != null && n._nozzleSettings.ID == nozzleConfig.Id);

                if (existingNozzle == null)
                {
                    // New nozzle found
                    return true;
                }

                // Check if existing nozzle has changes
                if (HasNozzlePositionChanged(existingNozzle, nozzleConfig))
                {
                    return true;
                }
            }

            return false;
        }

        private bool HasNozzlePositionChanged(Nozzle nozzle, NozzleConfiguration nozzleConfig)
        {
            if (Math.Abs(nozzleConfig.RotationAngleDegrees) > 0.1)
                return true;

            if (Math.Abs(nozzleConfig.OffsetMeters) > 0.001)
                return true;

            return false;
        }

        private void AddNozzleFromConfiguration(Compartment compartment, CompartmentConfiguration compartmentConfig, NozzleConfiguration nozzleConfig, int compartmentIndex)
        {
            try
            {
                Feature referencePlane = null;

                switch (nozzleConfig.ReferenceType)
                {
                    case NozzleReferenceType.LeftDishedEnd:
                        referencePlane = compartment.GetLeftEndPlane();
                        if (referencePlane == null) return;

                        compartment.AddNozzle(
                            _nozzlePositionSketchPath,
                            nozzleDocPath,
                            compartmentIndex,
                            referencePlane,
                            nozzleConfig.DistanceFromReference,
                            false,
                            2500);
                        break;

                    case NozzleReferenceType.RightDishedEnd:
                        referencePlane = compartment.GetRightEndPlane();
                        if (referencePlane == null) return;

                        compartment.AddNozzle(
                            _nozzlePositionSketchPath,
                            nozzleDocPath,
                            compartmentIndex,
                            referencePlane,
                            nozzleConfig.DistanceFromReference,
                            true,
                            2500);
                        break;

                    case NozzleReferenceType.OtherNozzle:
                        // Find the reference nozzle by ID
                        if (nozzleConfig.ReferenceNozzleId.HasValue)
                        {
                            var referenceNozzle = compartment.Nozzles?.FirstOrDefault(n => 
                                n._nozzleSettings != null && n._nozzleSettings.ID == nozzleConfig.ReferenceNozzleId.Value);

                            if (referenceNozzle != null)
                            {
                                referencePlane = referenceNozzle.GetPositionPlane();
                            }
                        }

                        // Fallback to last nozzle if reference not found
                        if (referencePlane == null && compartment.Nozzles != null && compartment.Nozzles.Count > 0)
                        {
                            referencePlane = compartment.Nozzles.Last().GetPositionPlane();
                        }

                        if (referencePlane == null) return;

                        compartment.AddNozzle(
                            _nozzlePositionSketchPath,
                            nozzleDocPath,
                            compartmentIndex,
                            referencePlane,
                            nozzleConfig.DistanceFromReference,
                            false,
                            2500);
                        break;
                }

                // After adding, get the newly created nozzle, set its ID and apply position changes
                if (compartment.Nozzles != null && compartment.Nozzles.Count > 0)
                {
                    var newNozzle = compartment.Nozzles.Last();
                    if (newNozzle._nozzleSettings != null)
                    {
                        newNozzle._nozzleSettings.ID = nozzleConfig.Id;
                    }

                    ApplyInitialPositionChanges(compartment, newNozzle, nozzleConfig);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding nozzle from configuration: {ex.Message}", "Add Nozzle Error");
            }
        }

        private void ApplyExistingNozzlePositionChanges(Compartment compartment, Nozzle nozzle, NozzleConfiguration nozzleConfig)
        {
            NozzleSettings s = nozzle._nozzleSettings;

            if (nozzleConfig.OffsetPosition != FlipDot.Central && nozzleConfig.OffsetMeters != 0
                && (nozzleConfig.OffsetPosition != s.OffsetPosition
                    || Math.Abs(nozzleConfig.OffsetMeters - s.OffsetMeters) > 0.0001))
            {
                double offset = nozzleConfig.OffsetPosition == FlipDot.Left
                    ? nozzleConfig.OffsetMeters
                    : -nozzleConfig.OffsetMeters;

                compartment.ActivateDocument();
                nozzle.SetOffset(offset);
                s.OffsetPosition = nozzleConfig.OffsetPosition;
                s.OffsetMeters   = nozzleConfig.OffsetMeters;
            }

            if (Math.Abs(nozzleConfig.RotationAngleDegrees - s.RotationAngleDegrees) > 0.01)
            {
                compartment.ActivateDocument();
                nozzle.SetRotationAngle(nozzleConfig.RotationAngleDegrees);
                s.RotationAngleDegrees = nozzleConfig.RotationAngleDegrees;
            }

            if (nozzleConfig.DistanceFromTopReferenceMeters > 0
                && (nozzleConfig.TopReferenceType != s.TopReferenceType
                    || Math.Abs(nozzleConfig.DistanceFromTopReferenceMeters - s.DistanceFromTopReferenceMeters) > 0.0001))
            {
                compartment.ActivateDocument();
                nozzle.ActivateDocument();
                nozzle.SetTopReferenceDistance(nozzleConfig.TopReferenceType, nozzleConfig.DistanceFromTopReferenceMeters);
                nozzle.CloseDocument();
                s.TopReferenceType                  = nozzleConfig.TopReferenceType;
                s.DistanceFromTopReferenceMeters    = nozzleConfig.DistanceFromTopReferenceMeters;
            }

            if (nozzleConfig.DistanceFromBottomReferenceMeters > 0
                && (nozzleConfig.BottomReferencePoint != s.BottomReferencePoint
                    || Math.Abs(nozzleConfig.DistanceFromBottomReferenceMeters - s.DistanceFromBottomReferenceMeters) > 0.0001
                    || nozzleConfig.IsLongNozzle != s.IsLongNozzle))
            {
                compartment.ActivateDocument();
                nozzle.ActivateDocument();
                nozzle.SetAdjustableComponentLength(nozzleConfig.BottomReferencePoint, nozzleConfig.DistanceFromBottomReferenceMeters, nozzleConfig.IsLongNozzle);
                nozzle.CloseDocument();
                s.BottomReferencePoint              = nozzleConfig.BottomReferencePoint;
                s.DistanceFromBottomReferenceMeters = nozzleConfig.DistanceFromBottomReferenceMeters;
                s.IsLongNozzle                      = nozzleConfig.IsLongNozzle;
            }
        }

        private void ApplyInitialPositionChanges(Compartment compartment, Nozzle nozzle, NozzleConfiguration nozzleConfig)
        {
            if (nozzleConfig.OffsetPosition != FlipDot.Central && nozzleConfig.OffsetMeters != 0)
            {
                double offset = nozzleConfig.OffsetPosition == FlipDot.Left
                    ? nozzleConfig.OffsetMeters
                    : -nozzleConfig.OffsetMeters;

                compartment.ActivateDocument();
                nozzle.SetOffset(offset);
            }

            if (nozzleConfig.RotationAngleDegrees != 0)
            {
                compartment.ActivateDocument();
                nozzle.SetRotationAngle(nozzleConfig.RotationAngleDegrees);
            }

            if (nozzleConfig.DistanceFromTopReferenceMeters > 0)
            {
                compartment.ActivateDocument();
                nozzle.ActivateDocument();
                nozzle.SetTopReferenceDistance(nozzleConfig.TopReferenceType, nozzleConfig.DistanceFromTopReferenceMeters);
                nozzle.CloseDocument();
            }

            if (nozzleConfig.DistanceFromBottomReferenceMeters > 0)
            {
                compartment.ActivateDocument();
                nozzle.ActivateDocument();
                nozzle.SetAdjustableComponentLength(nozzleConfig.BottomReferencePoint, nozzleConfig.DistanceFromBottomReferenceMeters, nozzleConfig.IsLongNozzle);
                nozzle.CloseDocument();
            }
        }

        // TODO: Uncomment and fix when NozzleAssemblyComponents property is available
        /*
        private void ApplyNozzleLengthChanges(Compartment compartment, Nozzle nozzle, NozzleConfiguration nozzleConfig)
        {
            // Check if nozzle has adjustable length component
            if (nozzle.NozzleAssemblyComponents == null || nozzle.NozzleAssemblyComponents.Count == 0)
                return;

            // Get the last component (typically the adjustable one)
            var lastComponent = nozzle.NozzleAssemblyComponents[nozzle.NozzleAssemblyComponents.Count - 1];

            if (!lastComponent.Settings.HasAdjustableLength)
                return;

            // Apply bottom reference distance change
            nozzle.AdjustComponentLengthFromBottom(
                compartment,
                lastComponent,
                nozzleConfig.BottomReferencePoint,
                nozzleConfig.DistanceFromBottomReferenceMeters,
                false); // Remove the IsLongNozzle parameter as it doesn't exist in NozzleConfiguration
        }
        */
    }
}

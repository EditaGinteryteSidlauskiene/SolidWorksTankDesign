using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Windows.Forms;

namespace SolidWorksTankDesign.Helpers
{
    internal static class MateManager
    {
        /// <summary>
        /// Creates mate
        /// </summary>
        /// <param name="componentFeature1"></param>
        /// <param name="componentFeature2"></param>
        /// <param name="alignmentType"></param>
        /// <param name="name"></param>
        public static Feature CreateMate(
            Feature componentFeature1, 
            Feature componentFeature2, 
            MateAlignment alignmentType, 
            string name) =>
           CreateMate((Entity)componentFeature1, (Entity)componentFeature2, alignmentType, name);

        /// <summary>
        /// Creates a coincident mate between two component features within the active SolidWorks assembly document.
        /// </summary>
        /// <param name="componentFeature1"></param>
        /// <param name="componentFeature2"></param>
        /// <param name="alignmentType"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="Exception"></exception>
        public static Feature CreateMate(
            Entity componentFeature1, 
            Entity componentFeature2, 
            MateAlignment alignmentType, 
            string name)
        {
            // Error handling: Null checks
            if (componentFeature1 == null || componentFeature2 == null)
            {
                throw new ArgumentNullException("One or more component features are null.");
            }

            ModelDoc2 currentModelDoc = SolidWorksDocumentProvider.GetActiveDoc();

            //Allows to select objects.
            SelectionMgr selectionManager = (SelectionMgr)currentModelDoc.SelectionManager;
            SelectData selectData = selectionManager.CreateSelectData();
            selectData.Mark = 1;

            //Creates a mate feature data object for the specified mate type. This is required to access CreateMate method
            CoincidentMateFeatureData coincidentMateFeatureData = (CoincidentMateFeatureData)((AssemblyDoc)currentModelDoc).CreateMateData((int)swMateType_e.swMateCOINCIDENT);

            // VEIKIA ---------------------------------------------------------------------------------------
            //Select entities
            componentFeature1.Select4(false, selectData);
            componentFeature2.Select4(true, selectData);
            // VEIKIA ---------------------------------------------------------------------------------------

            //Alignment - Aligned
            coincidentMateFeatureData.MateAlignment = (int)alignmentType;
            
            //Create Mate
            Feature mate = ((AssemblyDoc)currentModelDoc).CreateMate(coincidentMateFeatureData);

            // Error handling: Feature creation
            if (mate == null)
            {
                throw new Exception("Failed to create mate feature.");
            }

            //Assign name for the mate
            mate.Name = name;

            //Clear selection
            currentModelDoc.ClearSelection2(true);

            return mate;
        }

        /// <summary>
        /// Creates an angle mate feature within an active SolidWorks assembly document. 
        /// It establishes a relationship between two entities(externalEntity and componentEntity), 
        /// constraining their relative angular orientation with respect to a specified reference entity(referenceEntity).
        /// </summary>
        /// <param name="externalEntity"></param>
        /// <param name="componentEntity"></param>
        /// <param name="referenceEntity"></param>
        /// <param name="angle"></param>
        /// <param name="flipDimension"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        /// <exception cref="Exception"></exception>
        public static Feature CreateMate(
           Entity externalEntity,
           Entity componentEntity,
           Entity referenceEntity,
           double angle,
           bool flipDimension,
           string name)
        {
            // Error handling: Null checks
            if (externalEntity == null || componentEntity == null || referenceEntity == null)
            {
                throw new ArgumentNullException("One or more entities are null. Please ensure valid entity selections.");
            }

            // Input validation: Angle
            if (angle < 0 || angle > 360) // Adjust range as needed
            {
                throw new ArgumentOutOfRangeException("Angle must be between 0 and 360 degrees.");
            }

            AssemblyDoc activeAssemblyDoc = (AssemblyDoc)SolidWorksDocumentProvider.GetActiveDoc();

            MateFeatureData mateFeatureData = activeAssemblyDoc.CreateMateData((int)swMateType_e.swMateANGLE);
            AngleMateFeatureData angleMateFeatureData = (AngleMateFeatureData)mateFeatureData;

            angleMateFeatureData.FlipDimension = flipDimension;
            angleMateFeatureData.Angle = angle;
            angleMateFeatureData.EntitiesToMate = new Entity[] { externalEntity, componentEntity };
            angleMateFeatureData.ReferenceEntity = referenceEntity;

            Feature mate = activeAssemblyDoc.CreateMate(angleMateFeatureData);

            // Error handling: Feature creation
            if (mate == null)
            {
                throw new Exception("Failed to create mate feature.");
            }

            mate.Name = name;

            return mate;
        }

        /// <summary>
        /// Creates distance mate
        /// </summary>
        /// <param name="componentFeature1"></param>
        /// <param name="componentFeature2"></param>
        /// <param name="alignmentType"></param>
        /// <param name="distance"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Feature CreateMate(
            Feature componentFeature1,
            Feature componentFeature2,
            MateAlignment alignmentType,
            double distance,
            string name) =>
            CreateMate((Entity)componentFeature1, (Entity)componentFeature2, alignmentType, distance, name);

        /// <summary>
        /// Creates distance mate
        /// </summary>
        /// <param name="componentFeature1"></param>
        /// <param name="componentFeature2"></param>
        /// <param name="alignmentType"></param>
        /// <param name="distance"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        private static Feature CreateMate(
            Entity componentFeature1,
            Entity componentFeature2,
            MateAlignment alignmentType,
            double distance,
            string name)
        {
            // Error handling: Null checks
            if (componentFeature1 == null || componentFeature2 == null)
            {
                throw new ArgumentNullException("One or more component features are null.");
            }

            ModelDoc2 currentModelDoc = SolidWorksDocumentProvider.GetActiveDoc();

            //Allows to select objects.
            SelectionMgr selectionManager = (SelectionMgr)currentModelDoc.SelectionManager;
            SelectData selectData = selectionManager.CreateSelectData();
            selectData.Mark = 1;

            //Select entities
            componentFeature1.Select4(false, selectData);
            componentFeature2.Select4(true, selectData);

            //Create Mate
            Feature mate = (Feature)((AssemblyDoc)currentModelDoc).AddDistanceMate(
                AlignFromEnum: (int)alignmentType,
                Flip: false,
                Distance: distance,
                DistanceAbsUpperLimit: distance,
                DistanceAbsLowerLimit: distance,
                FirstArcCondition: 0,
                SecondArcCondition: 0,
                ErrorStatus: out _);

            //Assign name for the mate
            mate.Name = name;

            currentModelDoc.ClearSelection2(true);

            return mate;
        }

        /// <summary>
        /// Change alignment of a mate
        /// </summary>
        /// <param name="mate"></param>
        public static bool ChangeAlignment(Feature mate)
        {
            // Check mate type to ensure it's a coincident mate
            if (!(mate.GetDefinition() is CoincidentMateFeatureData mateData))
            {
                MessageBox.Show("The selected feature is not a coincident mate.");
                return false;
            }

            // Access mate's feature and get the MateFeatureData object. 
            mateData = mate.GetDefinition();

            // Cast the MateFeatureData object to a CoincidentMateFeatureData object. 
            CoincidentMateFeatureData coincMateData = (CoincidentMateFeatureData)mateData;

            // Change the alignment. 0 - ALIGNED
            coincMateData.MateAlignment = (coincMateData.MateAlignment == 0 ? (int)swMateAlign_e.swMateAlignANTI_ALIGNED : (int)swMateAlign_e.swMateAlignALIGNED);

            // Updates the definition of a feature with the new values in an associated feature data object
            return mate.ModifyDefinition(mateData, SolidWorksDocumentProvider.GetActiveDoc(), null);
        }

        /// <summary>
        /// Changes distance property of the mate
        /// </summary>
        /// <param name="mate"></param>
        /// <param name="newDistance"></param>
        public static void ChangeDistance(Feature mate, double newDistance)
        {
            // Check mate type to ensure it's a distance mate
            if (!(mate.GetDefinition() is DistanceMateFeatureData mateData))
            {
                MessageBox.Show("The selected feature is not a distance mate.");
                return;
            }

            // Validate the new distance
            if (newDistance < 0)
            {
                MessageBox.Show("Invalid distance. Distance must be a positive value.");
                return;
            }

            mateData.Distance = newDistance;
            // Toggle the dimension
            mateData.FlipDimension = true;

            // Updates the definition of a feature with the new values in an associated feature data object 
            mate.ModifyDefinition(mateData, SolidWorksDocumentProvider.GetActiveDoc(), null);
        }

        /// <summary>
        /// Moves entities to opposite sides of the dimension of this distance mate
        /// </summary>
        /// <param name="mate"></param>
        public static void FlipDimension(Feature mate)
        {
            // Check mate type to ensure it's a distance mate
            if (!(mate.GetDefinition() is IDistanceMateFeatureData mateData))
            {
                MessageBox.Show("The selected feature is not a distance mate.");
                return;
            }

            // Toggle the flip dimension state
            mateData.FlipDimension = !mateData.FlipDimension;

            try
            {
                // Updates the definition of a feature with the new values in an associated feature data object 
                mate.ModifyDefinition(mateData, SolidWorksDocumentProvider.GetActiveDoc(), null);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error flipping mate dimension: {ex.Message}");
            }
        }

        /// <summary>
        /// Edits an existing coincident mate within the active SolidWorks assembly document.
        /// </summary>
        /// <param name="mate">The existing coincident mate feature to be modified.</param>
        /// <param name="entity1">The first feature/entity involved in the mate.</param>
        /// <param name="entity2">The second feature/entity involved in the mate.</param>
        public static void EditCoincidentMate(Feature mate, Feature entity1, Feature entity2)
        {
            // Get a reference to the currently active SOLIDWORKS assembly document
            ModelDoc2 activeModelDoc = SolidWorksDocumentProvider.GetActiveDoc();

            // 1. SELECT FEATURES FOR EDITING:
            // Select the mate and both associated entities to initiate the mate edit operation
            mate.Select2(false, 1);
            entity1.Select2(true, 1);
            entity2.Select2(true, 1);

            // 2. RETRIEVE EXISTING MATE DATA:
            // Obtain the current definition of the coincident mate to access its properties
            CoincidentMateFeatureData coincMateData = (CoincidentMateFeatureData)mate.GetDefinition();

            // 3. APPLY MATE EDITS:
            // Perform the actual mate edit using SolidWorks' EditMate4 method. Parameters are 
            // based on the retrieved mate data, maintaining the coincident type, original alignment,
            // and locking rotation. 
            ((AssemblyDoc)activeModelDoc).EditMate4(
                MateTypeFromEnum: 0,
                AlignFromEnum: coincMateData.MateAlignment,
                Flip: false,
                Distance: 0,
                DistanceAbsUpperLimit: 0,
                DistanceAbsLowerLimit: 0,
                GearRatioNumerator: 0,
                GearRatioDenominator: 0,
                Angle: 0,
                AngleAbsUpperLimit: 0,
                AngleAbsLowerLimit: 0,
                ForPositioningOnly: false,
                LockRotation: true,
                WidthMateOption: 0,
                RepairMatesWithSameMissingEntity: false,
                ErrorStatus: out int _);

            // 4. CLEANUP:
            // Clear any remaining selections to avoid unintended interactions in the document
            activeModelDoc.ClearSelection2(true);
        }
    }
}

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
        /// Change alignment of a mate
        /// </summary>
        /// <param name="mate"></param>
        public static bool ChangeAlignment(Feature mate)
        {
            // Access mate's feature and get the MateFeatureData object. 
            IMateFeatureData mateData = mate.GetDefinition();

            // Cast the MateFeatureData object to a CoincidentMateFeatureData object. 
            CoincidentMateFeatureData coincMateData = (CoincidentMateFeatureData)mateData;

            // Change the alignment. 0 - ALIGNED
            coincMateData.MateAlignment = (coincMateData.MateAlignment == 0 ? (int)swMateAlign_e.swMateAlignANTI_ALIGNED : (int)swMateAlign_e.swMateAlignALIGNED);

            // Updates the definition of a feature with the new values in an associated feature data object
            return mate.ModifyDefinition(mateData, SolidWorksDocumentProvider.GetActiveDoc(), null);
        }
    }
}

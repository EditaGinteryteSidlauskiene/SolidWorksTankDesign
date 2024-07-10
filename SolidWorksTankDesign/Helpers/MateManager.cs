using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;

namespace SolidWorksTankDesign.Helpers
{
    internal static class MateManager
    {
        /// <summary>
        /// Creates mate
        /// </summary>
        /// <param name="ComponentFeature1"></param>
        /// <param name="ComponentFeature2"></param>
        /// <param name="AlignmentType"></param>
        /// <param name="Name"></param>
        public static Feature CreateMate(
            Feature componentFeature1, 
            Feature componentFeature2, 
            MateAlignment alignmentType, 
            string name) =>
           CreateMate((Entity)componentFeature1, (Entity)componentFeature2, alignmentType, name);

        /// <summary>
        /// Creates mate
        /// </summary>
        /// <param name="ModelDocument"></param>
        /// <param name="AssemblyFeature"></param>
        /// <param name="Component"></param>
        /// <param name="ComponentFeature"></param>
        public static Feature CreateMate(
            Entity componentFeature1, 
            Entity componentFeature2, 
            MateAlignment alignmentType, 
            string name)
        {
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

            //Assign name for the mate
            mate.Name = name;

            //Clear selection
            currentModelDoc.ClearSelection2(true);

            return mate;
        }

        public static Feature CreateMate(
           Entity ExternalEntity,
           Entity ComponentEntity,
           Entity ReferenceEntity,
           double Angle,
           bool FlipDimension,
           string Name)
        {
            AssemblyDoc activeAssemblyDoc = (AssemblyDoc)SolidWorksDocumentProvider.GetActiveDoc();

            MateFeatureData mateFeatureData = activeAssemblyDoc.CreateMateData((int)swMateType_e.swMateANGLE);
            AngleMateFeatureData angleMateFeatureData = (AngleMateFeatureData)mateFeatureData;

            angleMateFeatureData.FlipDimension = FlipDimension;
            angleMateFeatureData.Angle = Angle;
            angleMateFeatureData.EntitiesToMate = new Entity[] { ExternalEntity, ComponentEntity };
            angleMateFeatureData.ReferenceEntity = ReferenceEntity;

            Feature mate = activeAssemblyDoc.CreateMate(angleMateFeatureData);
            mate.Name = Name;

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

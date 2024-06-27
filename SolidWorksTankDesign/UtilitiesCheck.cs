using SolidWorks.Interop.sldworks;
using System.Collections.Generic;
using WarningAndErrorService;

namespace SolidWorksTankDesign
{
    internal class UtilitiesCheck
    {
        /// <summary>
        /// Checks if number of reference axis is correct.
        /// </summary>
        /// <param name="WarningService"></param>
        /// <param name="Document"></param>
        /// <param name="RequiredCount"></param>
        /// <param name="AxisList"></param>
        /// <returns></returns>
        public static bool IsNumberOfReferenceAxisCorrect(WarningService WarningService, ModelDoc2 Document, int RequiredCount, out List<FeatureAxis> AxisList)
        {
            //Get all reference axises
            AxisList = Utilities.GetAllReferenceAxisFeatures(Document);

            if (AxisList.Count != RequiredCount)
            {
                WarningService.AddWarning("Incorrect number of axis.");
            }

            return RequiredCount == AxisList.Count;
        }
    }
}

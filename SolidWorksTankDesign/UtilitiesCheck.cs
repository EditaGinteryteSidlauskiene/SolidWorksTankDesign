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
        /// <param name="warningService"></param>
        /// <param name="document"></param>
        /// <param name="requiredCount"></param>
        /// <param name="axisList"></param>
        /// <returns></returns>
        public static bool IsNumberOfReferenceAxisCorrect(WarningService warningService, ModelDoc2 document, int requiredCount, out List<FeatureAxis> axisList)
        {
            //Get all reference axises
            axisList = Utilities.GetAllReferenceAxisFeatures(document);

            if (axisList.Count != requiredCount)
            {
                warningService.AddWarning("Incorrect number of axis.");
            }

            return requiredCount == axisList.Count;
        }
    }
}

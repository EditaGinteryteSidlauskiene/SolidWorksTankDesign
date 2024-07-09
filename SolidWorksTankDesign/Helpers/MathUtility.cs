using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SolidWorksTankDesign
{
    internal static class MathUtility
    {
        /// <summary>
        /// Applies a mathematical transformation to a vector expressed as 3D coordinates
        /// </summary>
        /// <param name="transform"></param>
        /// <param name="vector"></param>
        /// <returns></returns>
        public static double[] TransformVector(ISldWorks solidWorksApplication, MathTransform transform, double[] vector)
        {
            //Create a SolidWorks MathVector object from the input vector
            //.
            MathVector mathVector = solidWorksApplication.GetMathUtility().CreateVector(vector);

            //Apply the provided mathematical transformation to the vector. The 'MultiplyTransform'
            //method updates the MathVector with the results of the transformation.
            mathVector = mathVector.MultiplyTransform(transform);

            //Return the transformed vector as a double array.
            return mathVector.ArrayData;
        }
    }
}

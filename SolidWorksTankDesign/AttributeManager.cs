using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using Attribute = SolidWorks.Interop.sldworks.Attribute;

namespace SolidWorksTankDesign
{
    internal static class AttributeManager
    {
        /// <summary>
        /// Creates a SolidWorks attribute with the specified parameters and attaches it to the given object.
        /// </summary>
        /// <param name="solidWorksApplication"></param>
        /// <param name="ownerDoc"></param>
        /// <param name="ownerObj"></param>
        /// <param name="parametersList"></param>
        /// <param name="attributeName"></param>
        /// <returns></returns>
        public static bool CreateAttribute(
            SldWorks solidWorksApplication, 
            ModelDoc2 ownerDoc, 
            object ownerObj, 
            List<(string ParameterName, string ParameterValue)> parametersList, 
            string attributeName)
        {
            AttributeDef attributeDefinition = null;
            Attribute attribute = null;

            // Validate input parameters
            if (solidWorksApplication == null)
                throw new ArgumentNullException(nameof(solidWorksApplication));
            if (ownerDoc == null)
                throw new ArgumentNullException(nameof(ownerDoc));
            if (ownerObj == null)
                throw new ArgumentNullException(nameof(ownerObj));
            if (parametersList == null)
                throw new ArgumentNullException(nameof(parametersList));
            if (string.IsNullOrEmpty(attributeName))
            {
                attributeName = "";
                MessageBox.Show("Attribute created without a name. Please add a name manually.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            try
            {
                // Create attribute definition
                attributeDefinition = solidWorksApplication.DefineAttribute(attributeName);

                foreach (var (parameterName, _) in parametersList)
                {
                    attributeDefinition.AddParameter(
                        NameIn: parameterName,
                        Type: (int)swParamType_e.swParamTypeString,
                        DefaultValue: 0.0,
                        Options: 0);
                }

                // Register this parameters into attribute definition
                attributeDefinition.Register();

                // Create the attribute instance and attach it to the owner object
                attribute = attributeDefinition.CreateInstance5(
                    OwnerDoc: ownerDoc,
                    OwnerObj: ownerObj,
                    NameIn: attributeName,
                    Options: 0,
                    ConfigurationOption: (int)swInConfigurationOpts_e.swThisConfiguration);

                // Set parameter values
                foreach (var (parameterName, parameterValue) in parametersList)
                {
                    // Get parameter
                    Parameter parameter = attribute.GetParameter(parameterName);

                    // Check if the parameter exists
                    if (parameter != null)
                    {
                        parameter.SetStringValue2(
                            StringValue: parameterValue,
                            ConfigurationOption: (int)swInConfigurationOpts_e.swAllConfiguration,
                            ConfigurationName: "");
                    }

                    else
                    {
                        throw new ArgumentException($"Parameter '{parameterName}' not found in the attribute '{attributeName}'.");
                    }
                }
            }
            
            catch (Exception ex)
            {
                if (attribute != null)
                {
                    attribute.Delete(false);
                }

                return false;
            }

            return true;
        }

        /// <summary>
        /// Retrieves the persistent reference ID stored as a string value in the specified attribute parameter.
        /// </summary>
        /// <param name="attributeOwnerDoc"></param>
        /// <param name="attributeName"></param>
        /// <param name="parameterName"></param>
        /// <returns></returns>
        public static string GetPersistentReferenceIdFromAttribute(
            ModelDoc2 attributeOwnerDoc,
            string attributeName,
            string parameterName)
        {
            // Validate input parameters
            if (attributeOwnerDoc == null)
                throw new ArgumentNullException(nameof(attributeOwnerDoc), "Owner document cannot be null.");
            if (string.IsNullOrEmpty(attributeName))
                throw new ArgumentException("Attribute name cannot be null or empty.", nameof(attributeName));
            if (string.IsNullOrEmpty(parameterName))
                throw new ArgumentException("Parameter name cannot be null or empty.", nameof(parameterName));

            // Get the attribute as a feature
            Feature attributeAsFeature = Utilities.GetFeatureByName(attributeOwnerDoc, attributeName);

            if (attributeAsFeature == null)
                throw new InvalidOperationException($"Attribute '{attributeName}' not found.");

            // Get the attribute object
            Attribute attribute = attributeAsFeature.GetSpecificFeature2();

            // Get the specified parameter from the attribute
            Parameter parameter = attribute.GetParameter(parameterName);

            if (parameter == null)
                throw new InvalidOperationException($"Parameter '{parameterName}' not found in attribute '{attributeName}'.");

            // Get the persistent reference ID
            string persistentReferenceId = parameter.GetStringValue();

            // Validate the retrieved ID
            if (string.IsNullOrEmpty(persistentReferenceId))
                throw new InvalidOperationException($"Parameter '{parameterName}' in attribute '{attributeName}' does not contain a valid value.");

            return persistentReferenceId;
        }

        /// <summary>
        /// Deletes the specified attribute from the given SolidWorks model document.
        /// </summary>
        /// <param name="attributeOwnerDoc"></param>
        /// <param name="attributeName"></param>
        /// <returns></returns>
        public static bool DeleteAttribute(ModelDoc2 attributeOwnerDoc, string attributeName)
        {
            // Validate input parameters
            if (attributeOwnerDoc == null)
                throw new ArgumentNullException(nameof(attributeOwnerDoc), "Owner document cannot be null.");
            if (string.IsNullOrEmpty(attributeName))
                throw new ArgumentException("Attribute name cannot be null or empty.", nameof(attributeName));

            // Get the attribute as a feature
            Feature attributeAsFeature = Utilities.GetFeatureByName(attributeOwnerDoc, attributeName);

            if (attributeAsFeature == null)
            {
                throw new InvalidOperationException($"Attribute {attributeName} was not found in document {nameof(attributeOwnerDoc)}");
            }

            // Get the attribute object
            Attribute attribute = attributeAsFeature.GetSpecificFeature2();

            // Delete the attribute and update the FeatureManager design tree
            return attribute.Delete(BuildTree: true);
        }

        /// <summary>
        /// Edits the value of a specific parameter within a SolidWorks attribute.
        /// </summary>
        /// <param name="attributeOwnerDoc">The SolidWorks document containing the attribute.</param>
        /// <param name="attributeName">The name of the attribute to modify.</param>
        /// <param name="parameterName">The name of the parameter to edit.</param>
        /// <param name="newValue">The new value to assign to the parameter.</param>
        /// <returns>True if the edit was successful, false otherwise.</returns>
        public static bool EditAttributeParameterValue(
            ModelDoc2 attributeOwnerDoc,
            string attributeName,
            string parameterName,
            string newValue)
        {
            // Validate input parameters
            if (attributeOwnerDoc == null)
                throw new ArgumentNullException(nameof(attributeOwnerDoc), "Owner document cannot be null.");
            if (string.IsNullOrEmpty(attributeName))
                throw new ArgumentException("Attribute name cannot be null or empty.", nameof(attributeName));
            if (string.IsNullOrEmpty(parameterName))
                throw new ArgumentException("Parameter name cannot be null or empty.", nameof(parameterName));

            // Get the attribute as feature
            Feature attributeFeature = Utilities.GetFeatureByName(attributeOwnerDoc, attributeName);

            // Handle the case where the attribute is not found
            if (attributeFeature == null)
            {
                throw new ArgumentException($"Attribute {attributeName} wa not found in the document {attributeOwnerDoc.GetTitle()}");
            }

            // Get the attribute and parameter objects
            Attribute attribute = attributeFeature.GetSpecificFeature2();
            Parameter parameter = attribute.GetParameter(parameterName);

            // Handle the case where the parameter is not found
            if (parameter == null)
            {
                throw new Exception($"Parameter {parameterName} was not found in the attribute {attributeName}.");
            }

            // Check if the parameter is of the correct type
            if (parameter.GetType() != (int)swParamType_e.swParamTypeString)
            {
                throw new Exception($"The type of parameter {parameterName} is not string and its value cannot be changed.");
            }

            // Set the new value
            return parameter.SetStringValue2(
                        StringValue: newValue,
                        ConfigurationOption: (int)swInConfigurationOpts_e.swAllConfiguration,
                        ConfigurationName: "");
        }

        /// <summary>
        /// Edits the value of a specific parameter within a SolidWorks attribute.
        /// </summary>
        /// <param name="attributeOwnerDoc">The SolidWorks document containing the attribute.</param>
        /// <param name="attributeName">The name of the attribute to modify.</param>
        /// <param name="parameterName">The name of the parameter to edit.</param>
        /// <param name="newValue">The new value to assign to the parameter.</param>
        /// <returns>True if the edit was successful, false otherwise.</returns>
        public static bool EditAttributeParameterValue(
            ModelDoc2 attributeOwnerDoc,
            string attributeName,
            string parameterName,
            double newValue)
        {
            // Validate input parameters
            if (attributeOwnerDoc == null)
                throw new ArgumentNullException(nameof(attributeOwnerDoc), "Owner document cannot be null.");
            if (string.IsNullOrEmpty(attributeName))
                throw new ArgumentException("Attribute name cannot be null or empty.", nameof(attributeName));
            if (string.IsNullOrEmpty(parameterName))
                throw new ArgumentException("Parameter name cannot be null or empty.", nameof(parameterName));

            // Get the attribute as feature
            Feature attributeFeature = Utilities.GetFeatureByName(attributeOwnerDoc, attributeName);

            // Handle the case where the attribute is not found
            if (attributeFeature == null)
            {
                throw new ArgumentException($"Attribute {attributeName} wa not found in the document {attributeOwnerDoc.GetTitle()}");
            }

            // Get the attribute and parameter objects
            Attribute attribute = attributeFeature.GetSpecificFeature2();
            Parameter parameter = attribute.GetParameter(parameterName);

            // Handle the case where the parameter is not found
            if (parameter == null)
            {
                throw new Exception($"Parameter {parameterName} was not found in the attribute {attributeName}.");
            }

            // Check if the parameter is of the correct type
            if (parameter.GetType() != (int)swParamType_e.swParamTypeDouble)
            {
                throw new Exception($"The type of parameter {parameterName} is not double and its value cannot be changed.");
            }

            // Set the new value
            return parameter.SetDoubleValue(
                        Value: newValue);
        }
    }
}

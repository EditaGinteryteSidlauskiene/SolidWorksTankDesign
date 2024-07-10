using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Windows.Forms;
using Attribute = SolidWorks.Interop.sldworks.Attribute;

namespace SolidWorksTankDesign
{
    internal static class AttributeManager
    {
        /// <summary>
        /// Creates an attribute in a SolidWorks document with the document as the owner. This method will be executed only once unless there is a need to change something. In that case, the attribute will have to be deleted before creating a new one.
        /// </summary>
        /// <param name="solidWorksApplication">The SolidWorks application instance.</param>
        /// <param name="ownerDoc">The document where the attribute will be created and owned.</param>
        /// <param name="attributeDefinitionName">Name of the attribute definition to use or create.</param>
        /// <param name="attributeName">The name of the new attribute.</param>
        /// <param name="parameterName">The name of the parameter within the attribute.</param>
        /// <param name="parameterValue">The value to set for the parameter.</param>
        /// <returns>True if the attribute was created successfully, false otherwise.</returns>
        /// <exception cref="ArgumentNullException">Thrown if any required argument is null.</exception>
        /// <exception cref="ArgumentException">Thrown if any argument is invalid (e.g., empty string).</exception>
        public static bool CreateAttribute(
            SldWorks solidWorksApplication, 
            ModelDoc2 ownerDoc,
            string attributeDefinitionName,
            string attributeName,
            string parameterName,
            string parameterValue)
        {
            // 1. Robust Input Validation
            if (solidWorksApplication == null) throw new ArgumentNullException(nameof(solidWorksApplication));
            if (ownerDoc == null) throw new ArgumentNullException(nameof(ownerDoc));
            if (string.IsNullOrWhiteSpace(attributeDefinitionName)) throw new ArgumentException("Attribute definition name cannot be null or empty.", nameof(attributeDefinitionName));
            if (string.IsNullOrWhiteSpace(attributeName)) throw new ArgumentException("Attribute name cannot be null or empty.", nameof(attributeName));
            if (string.IsNullOrWhiteSpace(parameterName)) throw new ArgumentException("Parameter name cannot be null or empty.", nameof(parameterName));

            try
            {
                // 2.Create Attribute Definition
                AttributeDef attributeDefinition = CreateAttributeDefinitionString(
                    solidWorksApplication,
                    attributeDefinitionName,
                    parameterName);

                // 3.Verify definition creation and create the attribute
                if (attributeDefinition == null ||
                    !CreateAttribute(
                    attributeDefinition,
                    ownerDoc,
                    ownerDoc,
                    attributeName)) 
                {
                    return false;
                }

                // Attempt to set the parameter value within the newly created attribute
                if (!EditAttributeParameterValue(
                    ownerDoc,
                    attributeName,
                    parameterName,
                    parameterValue))
                {
                    // Notify the user if parameter value update failed
                    MessageBox.Show($"The {parameterName} parameter value in attribute {attributeName} was not changed.");
                }
                
            }
            catch (Exception ex)
            {
                return false;
            }

            return true;
        }

        private static AttributeDef CreateAttributeDefinitionString(
            SldWorks solidWorksApplication,
            string attributeDefinitionName,
            string parameterName)
        {
            // Validate input parameters
            if (solidWorksApplication == null)
                throw new ArgumentNullException(nameof(solidWorksApplication));
            if (string.IsNullOrEmpty(attributeDefinitionName))
                throw new ArgumentNullException(nameof(attributeDefinitionName));

            // Create attribute definition
            AttributeDef attributeDefinition = solidWorksApplication.DefineAttribute(attributeDefinitionName);

            attributeDefinition.AddParameter(
                NameIn: parameterName,
                Type: (int)swParamType_e.swParamTypeString,
                DefaultValue: 0.0,
                Options: 0);

            // Register this parameters into attribute definition
            attributeDefinition.Register();

            return attributeDefinition;
        }

        /// <summary>
        /// Creates a SolidWorks attribute with the specified parameters and attaches it to the given object.
        /// </summary>
        /// <param name="solidWorksApplication"></param>
        /// <param name="ownerDoc"></param>
        /// <param name="ownerObj"></param>
        /// <param name="parametersList"></param>
        /// <param name="attributeName"></param>
        /// <returns></returns>
        private static bool CreateAttribute(
            AttributeDef attributeDefinition,
            ModelDoc2 ownerDoc,
            object ownerObj,
            string attributeName)
        {
            Attribute attribute = null;

            // Validate input parameters
            if (attributeDefinition == null)
                throw new ArgumentNullException(nameof(attributeDefinition));
            if (ownerDoc == null)
                throw new ArgumentNullException(nameof(ownerDoc));
            if (ownerObj == null)
                throw new ArgumentNullException(nameof(ownerObj));
            if (string.IsNullOrEmpty(attributeName))
            {
                attributeName = "";
                MessageBox.Show("Attribute created without a name. Please add a name manually.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            try
            {
                // Create the attribute instance and attach it to the owner object
                attribute = attributeDefinition.CreateInstance5(
                    OwnerDoc: ownerDoc,
                    OwnerObj: ownerObj,
                    NameIn: attributeName,
                    Options: 0,
                    ConfigurationOption: (int)swInConfigurationOpts_e.swAllConfiguration);
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
            Feature attributeAsFeature = FeatureManager.GetFeatureByName(SolidWorksDocumentProvider.GetActiveDoc(), attributeName);

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
            Feature attributeFeature = FeatureManager.GetFeatureByName(attributeOwnerDoc, attributeName);

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
            parameter.SetStringValue2(
                        StringValue: newValue,
                        ConfigurationOption: (int)swInConfigurationOpts_e.swAllConfiguration,
                        ConfigurationName: "");

            return true;
        }

        /// <summary>
        /// Method extracts the string value of a specific parameter from a given attribute within a model document.
        /// If the string value is not retrieved, returns empty string.
        /// </summary>
        /// <param name="attributeOwnerDoc"></param>
        /// <param name="attributeName"></param>
        /// <param name="parameterName"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        public static string GetAttributeParameterValue(ModelDoc2 attributeOwnerDoc, string attributeName, string parameterName)
        {
            // 1. Null Checks and Early Returns
            if (attributeOwnerDoc == null || string.IsNullOrWhiteSpace(attributeName) || string.IsNullOrWhiteSpace(parameterName))
            {
                throw new ArgumentNullException("One or more arguments are null or empty.");
            }

            // 2. Feature Retrieval (Improved Error Handling)
            Feature attributeFeature = FeatureManager.GetFeatureByName(attributeOwnerDoc, attributeName);
            if (attributeFeature == null)
            {
                throw new InvalidOperationException($"Attribute '{attributeName}' not found.");
            }

            // 3. Attribute Cast and Parameter Retrieval
            Attribute attribute = attributeFeature.GetSpecificFeature2();
            Parameter attributeParameter = attribute.GetParameter(parameterName);

            // 4. Parameter Value Retrieval (with Null Handling)
            if (attributeParameter == null)
            {
                throw new InvalidOperationException($"Parameter '{parameterName}' not found in attribute '{attributeName}'.");
            }

            return attributeParameter.GetStringValue() ?? string.Empty; // Handle null values
        }
    }
}

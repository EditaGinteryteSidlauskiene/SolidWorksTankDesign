using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System.Reflection;

namespace SolidWorksTankDesign
{
    /// <summary>
    /// Custom contract resolver for Newtonsoft.Json to include private properties in serialization/deserialization.
    /// </summary>
    public class PrivatePropertyContractResolver : DefaultContractResolver
    {
        /// <summary>
        /// Overrides the default behavior to determine which properties to include in the JSON.
        /// </summary>
        /// <param name="member">The member information for the property.</param>
        /// <param name="memberSerialization">The serialization settings for the member.</param>
        /// <returns>A JsonProperty object representing the property's serialization settings.</returns>
        protected override JsonProperty CreateProperty(MemberInfo member, MemberSerialization memberSerialization)
        {
            // Get the default JsonProperty for this member.
            var prop = base.CreateProperty(member, memberSerialization);

            // Check if the property is read-only (doesn't have a public setter).
            if (!prop.Writable) // Only non-writable properties should be included
            {
                // If it's a PropertyInfo object (not a field or something else),
                // check if it has a private setter we can use.
                var property = member as PropertyInfo;

                if (property != null)
                    // Set the Writable flag to true if there's a private setter.
                    prop.Writable = property.GetSetMethod(true) != null;
            }
            // Return the modified JsonProperty object.
            return prop;
        }
    }

}

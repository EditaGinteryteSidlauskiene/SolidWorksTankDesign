using System;
using System.ComponentModel;

namespace SolidWorksTankDesign.MVP.Enums
{
    public static class EnumManager
    {
        public static string GetEnumDescription(Enum value)
        {
            var fieldInfo = value.GetType().GetField(value.ToString());
            var attributes = (DescriptionAttribute[])fieldInfo.GetCustomAttributes(typeof(DescriptionAttribute),
                                false);

            return attributes.Length > 0 ? attributes[0].Description : value.ToString();

        }
    }
}

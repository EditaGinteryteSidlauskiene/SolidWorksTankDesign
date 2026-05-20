using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace FlangeSetup
{
    public static class FlangeType05Setup
    {
        public static void AddConfigurations(ModelDoc2 doc)
        {
            // Type 05 — configurations where dimensions are unique (target PN = input PN).
            // Based on get_pn_type_05 from EN 1092-1 synoptic table (Table 7).
            var configurations = new[]
            {
                // PN 6 — DN 10..600
                "Pn6_Dn10",  "Pn6_Dn15",  "Pn6_Dn20",  "Pn6_Dn25",  "Pn6_Dn32",
                "Pn6_Dn40",  "Pn6_Dn50",  "Pn6_Dn65",  "Pn6_Dn80",  "Pn6_Dn100",
                "Pn6_Dn125", "Pn6_Dn150", "Pn6_Dn200", "Pn6_Dn250", "Pn6_Dn300",
                "Pn6_Dn350", "Pn6_Dn400", "Pn6_Dn450", "Pn6_Dn500", "Pn6_Dn600",

                // PN 10 — DN 200..600 (DN ≤ 40 → Pn40, DN ≤ 150 → Pn16)
                "Pn10_Dn200", "Pn10_Dn250", "Pn10_Dn300", "Pn10_Dn350", "Pn10_Dn400",
                "Pn10_Dn450", "Pn10_Dn500", "Pn10_Dn600",

                // PN 16 — DN 50..600 (DN ≤ 40 → Pn40)
                "Pn16_Dn50",  "Pn16_Dn65",  "Pn16_Dn80",  "Pn16_Dn100", "Pn16_Dn125",
                "Pn16_Dn150", "Pn16_Dn200", "Pn16_Dn250", "Pn16_Dn300", "Pn16_Dn350",
                "Pn16_Dn400", "Pn16_Dn450", "Pn16_Dn500", "Pn16_Dn600",

                // PN 25 — DN 200..600 (DN ≤ 150 → Pn40)
                "Pn25_Dn200", "Pn25_Dn250", "Pn25_Dn300", "Pn25_Dn350", "Pn25_Dn400",
                "Pn25_Dn450", "Pn25_Dn500", "Pn25_Dn600",

                // PN 40 — DN 10..600 (extended vs Type 01 which stops at DN 400)
                "Pn40_Dn10",  "Pn40_Dn15",  "Pn40_Dn20",  "Pn40_Dn25",  "Pn40_Dn32",
                "Pn40_Dn40",  "Pn40_Dn50",  "Pn40_Dn65",  "Pn40_Dn80",  "Pn40_Dn100",
                "Pn40_Dn125", "Pn40_Dn150", "Pn40_Dn200", "Pn40_Dn250", "Pn40_Dn300",
                "Pn40_Dn350", "Pn40_Dn400", "Pn40_Dn450", "Pn40_Dn500", "Pn40_Dn600",

                // PN 63 — DN 50..400 (DN ≤ 40 → Pn100)
                "Pn63_Dn50",  "Pn63_Dn65",  "Pn63_Dn80",  "Pn63_Dn100", "Pn63_Dn125",
                "Pn63_Dn150", "Pn63_Dn200", "Pn63_Dn250", "Pn63_Dn300", "Pn63_Dn350",
                "Pn63_Dn400",

                // PN 100 — DN 10..350
                "Pn100_Dn10",  "Pn100_Dn15",  "Pn100_Dn20",  "Pn100_Dn25",  "Pn100_Dn32",
                "Pn100_Dn40",  "Pn100_Dn50",  "Pn100_Dn65",  "Pn100_Dn80",  "Pn100_Dn100",
                "Pn100_Dn125", "Pn100_Dn150", "Pn100_Dn200", "Pn100_Dn250", "Pn100_Dn300",
                "Pn100_Dn350",

                // PN 160 — DN 10..300, excluding Dn20 and Dn32
                "Pn160_Dn10",  "Pn160_Dn15",  "Pn160_Dn25",  "Pn160_Dn40",  "Pn160_Dn50",
                "Pn160_Dn65",  "Pn160_Dn80",  "Pn160_Dn100", "Pn160_Dn125", "Pn160_Dn150",
                "Pn160_Dn200", "Pn160_Dn250", "Pn160_Dn300",

                // PN 250 — DN 15..250, excluding Dn20 and Dn32 (DN 10 → Pn320)
                "Pn250_Dn15",  "Pn250_Dn25",  "Pn250_Dn40",  "Pn250_Dn50",  "Pn250_Dn65",
                "Pn250_Dn80",  "Pn250_Dn100", "Pn250_Dn125", "Pn250_Dn150", "Pn250_Dn200",
                "Pn250_Dn250",

                // PN 320 — DN 10..250, excluding Dn20 and Dn32
                "Pn320_Dn10",  "Pn320_Dn15",  "Pn320_Dn25",  "Pn320_Dn40",  "Pn320_Dn50",
                "Pn320_Dn65",  "Pn320_Dn80",  "Pn320_Dn100", "Pn320_Dn125", "Pn320_Dn150",
                "Pn320_Dn200", "Pn320_Dn250",

                // PN 400 — DN 10..200, excluding Dn20 and Dn32
                "Pn400_Dn10",  "Pn400_Dn15",  "Pn400_Dn25",  "Pn400_Dn40",  "Pn400_Dn50",
                "Pn400_Dn65",  "Pn400_Dn80",  "Pn400_Dn100", "Pn400_Dn125", "Pn400_Dn150",
                "Pn400_Dn200",
            };

            int added = 0;
            foreach (string name in configurations)
            {
                // Skip if configuration already exists
                if (doc.GetConfigurationByName(name) != null) continue;

                doc.AddConfiguration3(
                    name, "", "",
                    (int)swConfigurationOptions2_e.swConfigOption_DontActivate);
                added++;
            }
        }
    }
}

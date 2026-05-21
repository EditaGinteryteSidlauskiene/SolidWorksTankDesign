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
                "Pn6_Dn10_Type_A",  "Pn6_Dn15_Type_A",  "Pn6_Dn20_Type_A",  "Pn6_Dn25_Type_A",  "Pn6_Dn32_Type_A",
                "Pn6_Dn40_Type_A",  "Pn6_Dn50_Type_A",  "Pn6_Dn65_Type_A",  "Pn6_Dn80_Type_A",  "Pn6_Dn100_Type_A",
                "Pn6_Dn125_Type_A", "Pn6_Dn150_Type_A", "Pn6_Dn200_Type_A", "Pn6_Dn250_Type_A", "Pn6_Dn300_Type_A",
                "Pn6_Dn350_Type_A", "Pn6_Dn400_Type_A", "Pn6_Dn450_Type_A", "Pn6_Dn500_Type_A", "Pn6_Dn600_Type_A",

                // PN 10 — DN 200..600 (DN ≤ 40 → Pn40, DN ≤ 150 → Pn16)
                "Pn10_Dn200_Type_A", "Pn10_Dn250_Type_A", "Pn10_Dn300_Type_A", "Pn10_Dn350_Type_A", "Pn10_Dn400_Type_A",
                "Pn10_Dn450_Type_A", "Pn10_Dn500_Type_A", "Pn10_Dn600_Type_A",

                // PN 16 — DN 50..600 (DN ≤ 40 → Pn40)
                "Pn16_Dn50_Type_A",  "Pn16_Dn65_Type_A",  "Pn16_Dn80_Type_A",  "Pn16_Dn100_Type_A", "Pn16_Dn125_Type_A",
                "Pn16_Dn150_Type_A", "Pn16_Dn200_Type_A", "Pn16_Dn250_Type_A", "Pn16_Dn300_Type_A", "Pn16_Dn350_Type_A",
                "Pn16_Dn400_Type_A", "Pn16_Dn450_Type_A", "Pn16_Dn500_Type_A", "Pn16_Dn600_Type_A",

                // PN 25 — DN 200..600 (DN ≤ 150 → Pn40)
                "Pn25_Dn200_Type_A", "Pn25_Dn250_Type_A", "Pn25_Dn300_Type_A", "Pn25_Dn350_Type_A", "Pn25_Dn400_Type_A",
                "Pn25_Dn450_Type_A", "Pn25_Dn500_Type_A", "Pn25_Dn600_Type_A",

                // PN 40 — DN 10..600 (extended vs Type 01 which stops at DN 400)
                "Pn40_Dn10_Type_A",  "Pn40_Dn15_Type_A",  "Pn40_Dn20_Type_A",  "Pn40_Dn25_Type_A",  "Pn40_Dn32_Type_A",
                "Pn40_Dn40_Type_A",  "Pn40_Dn50_Type_A",  "Pn40_Dn65_Type_A",  "Pn40_Dn80_Type_A",  "Pn40_Dn100_Type_A",
                "Pn40_Dn125_Type_A", "Pn40_Dn150_Type_A", "Pn40_Dn200_Type_A", "Pn40_Dn250_Type_A", "Pn40_Dn300_Type_A",
                "Pn40_Dn350_Type_A", "Pn40_Dn400_Type_A", "Pn40_Dn450_Type_A", "Pn40_Dn500_Type_A", "Pn40_Dn600_Type_A",

                // PN 63 — DN 50..400 (DN ≤ 40 → Pn100)
                "Pn63_Dn50_Type_A",  "Pn63_Dn65_Type_A",  "Pn63_Dn80_Type_A",  "Pn63_Dn100_Type_A", "Pn63_Dn125_Type_A",
                "Pn63_Dn150_Type_A", "Pn63_Dn200_Type_A", "Pn63_Dn250_Type_A", "Pn63_Dn300_Type_A", "Pn63_Dn350_Type_A",
                "Pn63_Dn400_Type_A",

                // PN 100 — DN 10..350
                "Pn100_Dn10_Type_A",  "Pn100_Dn15_Type_A",  "Pn100_Dn20_Type_A",  "Pn100_Dn25_Type_A",  "Pn100_Dn32_Type_A",
                "Pn100_Dn40_Type_A",  "Pn100_Dn50_Type_A",  "Pn100_Dn65_Type_A",  "Pn100_Dn80_Type_A",  "Pn100_Dn100_Type_A",
                "Pn100_Dn125_Type_A", "Pn100_Dn150_Type_A", "Pn100_Dn200_Type_A", "Pn100_Dn250_Type_A", "Pn100_Dn300_Type_A",
                "Pn100_Dn350_Type_A",

                // PN 160 — DN 10..300, excluding Dn20 and Dn32
                "Pn160_Dn10_Type_A",  "Pn160_Dn15_Type_A",  "Pn160_Dn25_Type_A",  "Pn160_Dn40_Type_A",  "Pn160_Dn50_Type_A",
                "Pn160_Dn65_Type_A",  "Pn160_Dn80_Type_A",  "Pn160_Dn100_Type_A", "Pn160_Dn125_Type_A", "Pn160_Dn150_Type_A",
                "Pn160_Dn200_Type_A", "Pn160_Dn250_Type_A", "Pn160_Dn300_Type_A",

                // PN 250 — DN 15..250, excluding Dn20 and Dn32 (DN 10 → Pn320)
                "Pn250_Dn15_Type_A",  "Pn250_Dn25_Type_A",  "Pn250_Dn40_Type_A",  "Pn250_Dn50_Type_A",  "Pn250_Dn65_Type_A",
                "Pn250_Dn80_Type_A",  "Pn250_Dn100_Type_A", "Pn250_Dn125_Type_A", "Pn250_Dn150_Type_A", "Pn250_Dn200_Type_A",
                "Pn250_Dn250_Type_A",

                // PN 320 — DN 10..250, excluding Dn20 and Dn32
                "Pn320_Dn10_Type_A",  "Pn320_Dn15_Type_A",  "Pn320_Dn25_Type_A",  "Pn320_Dn40_Type_A",  "Pn320_Dn50_Type_A",
                "Pn320_Dn65_Type_A",  "Pn320_Dn80_Type_A",  "Pn320_Dn100_Type_A", "Pn320_Dn125_Type_A", "Pn320_Dn150_Type_A",
                "Pn320_Dn200_Type_A", "Pn320_Dn250_Type_A",

                // PN 400 — DN 10..200, excluding Dn20 and Dn32
                "Pn400_Dn10_Type_A",  "Pn400_Dn15_Type_A",  "Pn400_Dn25_Type_A",  "Pn400_Dn40_Type_A",  "Pn400_Dn50_Type_A",
                "Pn400_Dn65_Type_A",  "Pn400_Dn80_Type_A",  "Pn400_Dn100_Type_A", "Pn400_Dn125_Type_A", "Pn400_Dn150_Type_A",
                "Pn400_Dn200_Type_A",
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

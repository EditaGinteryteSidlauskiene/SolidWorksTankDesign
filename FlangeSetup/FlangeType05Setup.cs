using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Runtime.InteropServices;

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

        public static void ChangeDimensions(ModelDoc2 doc)
        {
            string originalConfig = ((Configuration)doc.GetActiveConfiguration()).Name;

            Feature flangeSketch            = GetFeatureByName(doc, "Flange sketch");
            Feature boltHoleSketch          = GetFeatureByName(doc, "Bolt hole sketch");
            Feature boltHolesPatternFeature = GetFeatureByName(doc, "Bolt hole pattern");

            // ── Pn6 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn10_Type_A",  D: 75.0,  C4: 12.0, K: 50.0,  L: 11.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn15_Type_A",  D: 80.0,  C4: 12.0, K: 55.0,  L: 11.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn20_Type_A",  D: 90.0,  C4: 14.0, K: 65.0,  L: 11.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn25_Type_A",  D: 100.0, C4: 14.0, K: 75.0,  L: 11.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn32_Type_A",  D: 120.0, C4: 14.0, K: 90.0,  L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn40_Type_A",  D: 130.0, C4: 14.0, K: 100.0, L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn50_Type_A",  D: 140.0, C4: 14.0, K: 110.0, L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn65_Type_A",  D: 160.0, C4: 14.0, K: 130.0, L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn80_Type_A",  D: 190.0, C4: 16.0, K: 150.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn100_Type_A", D: 210.0, C4: 16.0, K: 170.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn125_Type_A", D: 240.0, C4: 18.0, K: 200.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn150_Type_A", D: 265.0, C4: 18.0, K: 225.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn200_Type_A", D: 320.0, C4: 20.0, K: 280.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn250_Type_A", D: 375.0, C4: 22.0, K: 335.0, L: 18.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn300_Type_A", D: 440.0, C4: 22.0, K: 395.0, L: 22.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn350_Type_A", D: 490.0, C4: 22.0, K: 445.0, L: 22.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn400_Type_A", D: 540.0, C4: 22.0, K: 495.0, L: 22.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn450_Type_A", D: 595.0, C4: 24.0, K: 550.0, L: 22.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn500_Type_A", D: 645.0, C4: 24.0, K: 600.0, L: 22.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn600_Type_A", D: 755.0, C4: 30.0, K: 705.0, L: 26.0, boltCount: 20);

            // ── Pn10 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn200_Type_A", D: 340.0, C4: 24.0, K: 295.0, L: 22.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn250_Type_A", D: 395.0, C4: 26.0, K: 350.0, L: 22.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn300_Type_A", D: 445.0, C4: 26.0, K: 400.0, L: 22.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn350_Type_A", D: 505.0, C4: 26.0, K: 460.0, L: 22.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn400_Type_A", D: 565.0, C4: 26.0, K: 515.0, L: 26.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn450_Type_A", D: 615.0, C4: 28.0, K: 565.0, L: 26.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn500_Type_A", D: 670.0, C4: 28.0, K: 620.0, L: 26.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn600_Type_A", D: 780.0, C4: 34.0, K: 725.0, L: 30.0, boltCount: 20);

            // ── Pn16 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn50_Type_A",  D: 165.0, C4: 18.0, K: 125.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn65_Type_A",  D: 185.0, C4: 18.0, K: 145.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn80_Type_A",  D: 200.0, C4: 20.0, K: 160.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn100_Type_A", D: 220.0, C4: 20.0, K: 180.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn125_Type_A", D: 250.0, C4: 22.0, K: 210.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn150_Type_A", D: 285.0, C4: 22.0, K: 240.0, L: 22.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn200_Type_A", D: 340.0, C4: 24.0, K: 295.0, L: 22.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn250_Type_A", D: 405.0, C4: 26.0, K: 355.0, L: 26.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn300_Type_A", D: 460.0, C4: 28.0, K: 410.0, L: 26.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn350_Type_A", D: 520.0, C4: 30.0, K: 470.0, L: 26.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn400_Type_A", D: 580.0, C4: 32.0, K: 525.0, L: 30.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn450_Type_A", D: 640.0, C4: 40.0, K: 585.0, L: 30.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn500_Type_A", D: 715.0, C4: 44.0, K: 650.0, L: 33.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn600_Type_A", D: 840.0, C4: 54.0, K: 770.0, L: 36.0, boltCount: 20);

            // ── Pn25 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn200_Type_A", D: 360.0, C4: 30.0, K: 310.0, L: 26.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn250_Type_A", D: 425.0, C4: 32.0, K: 370.0, L: 30.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn300_Type_A", D: 485.0, C4: 34.0, K: 430.0, L: 30.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn350_Type_A", D: 555.0, C4: 38.0, K: 490.0, L: 33.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn400_Type_A", D: 620.0, C4: 40.0, K: 550.0, L: 36.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn450_Type_A", D: 670.0, C4: 50.0, K: 600.0, L: 36.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn500_Type_A", D: 730.0, C4: 51.0, K: 660.0, L: 36.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn600_Type_A", D: 845.0, C4: 66.0, K: 770.0, L: 39.0, boltCount: 20);

            // ── Pn40 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn10_Type_A",  D: 90.0,  C4: 16.0, K: 60.0,  L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn15_Type_A",  D: 95.0,  C4: 16.0, K: 65.0,  L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn20_Type_A",  D: 105.0, C4: 18.0, K: 75.0,  L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn25_Type_A",  D: 115.0, C4: 18.0, K: 85.0,  L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn32_Type_A",  D: 140.0, C4: 18.0, K: 100.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn40_Type_A",  D: 150.0, C4: 18.0, K: 110.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn50_Type_A",  D: 165.0, C4: 20.0, K: 125.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn65_Type_A",  D: 185.0, C4: 22.0, K: 145.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn80_Type_A",  D: 200.0, C4: 24.0, K: 160.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn100_Type_A", D: 235.0, C4: 24.0, K: 190.0, L: 22.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn125_Type_A", D: 270.0, C4: 26.0, K: 220.0, L: 26.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn150_Type_A", D: 300.0, C4: 28.0, K: 250.0, L: 26.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn200_Type_A", D: 375.0, C4: 36.0, K: 320.0, L: 30.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn250_Type_A", D: 450.0, C4: 38.0, K: 385.0, L: 33.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn300_Type_A", D: 515.0, C4: 42.0, K: 450.0, L: 33.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn350_Type_A", D: 580.0, C4: 46.0, K: 510.0, L: 36.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn400_Type_A", D: 660.0, C4: 50.0, K: 585.0, L: 39.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn450_Type_A", D: 685.0, C4: 57.0, K: 610.0, L: 39.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn500_Type_A", D: 755.0, C4: 57.0, K: 670.0, L: 42.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn600_Type_A", D: 890.0, C4: 72.0, K: 795.0, L: 48.0, boltCount: 20);

            // ── Pn63 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn50_Type_A",  D: 180.0, C4: 26.0, K: 135.0, L: 22.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn65_Type_A",  D: 205.0, C4: 26.0, K: 160.0, L: 22.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn80_Type_A",  D: 215.0, C4: 28.0, K: 170.0, L: 22.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn100_Type_A", D: 250.0, C4: 30.0, K: 200.0, L: 26.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn125_Type_A", D: 295.0, C4: 34.0, K: 240.0, L: 30.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn150_Type_A", D: 345.0, C4: 36.0, K: 280.0, L: 33.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn200_Type_A", D: 415.0, C4: 42.0, K: 345.0, L: 36.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn250_Type_A", D: 470.0, C4: 46.0, K: 400.0, L: 36.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn300_Type_A", D: 530.0, C4: 52.0, K: 460.0, L: 36.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn350_Type_A", D: 600.0, C4: 56.0, K: 525.0, L: 39.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn400_Type_A", D: 670.0, C4: 60.0, K: 585.0, L: 42.0, boltCount: 16);

            // Delete configurations not covered by data (Pn100, Pn160, Pn250, Pn320, Pn400)
            string[] allNames = doc.GetConfigurationNames();
            foreach (string name in allNames)
            {
                if (name.StartsWith("Pn100_") || name.StartsWith("Pn160_") ||
                    name.StartsWith("Pn250_") || name.StartsWith("Pn320_") ||
                    name.StartsWith("Pn400_"))
                {
                    doc.DeleteConfiguration2(name);
                }
            }

            Marshal.ReleaseComObject(flangeSketch);
            Marshal.ReleaseComObject(boltHoleSketch);
            Marshal.ReleaseComObject(boltHolesPatternFeature);

            doc.ShowConfiguration2(originalConfig);
            doc.EditRebuild3();
        }

        public static void ChangeDimensionsForFaces(ModelDoc2 doc)
        {
            string[] allNames = doc.GetConfigurationNames();

            string[] typeANames = Array.FindAll(allNames, n => n.EndsWith("_Type_A"));

            for (int i = 80; i < 90 &&i < typeANames.Length; i++)
            {
                string configurationName = typeANames[i];

                int pn = ParsePn(configurationName);
                int dn = ParseDn(configurationName);
                if (pn == 0 || dn == 0) continue;

                ConfigurationManager configManager = doc.ConfigurationManager;
                try
                {
                    // Type B — all PN
                    AddTypeBFaceConfiguration(doc, configManager, configurationName, pn, dn);

                    // Types C–F — not for Pn6
                    if (pn != 6)
                    {
                        AddTypeCFaceConfiguration(doc, configManager, configurationName, dn);
                        AddTypeDFaceConfiguration(doc, configManager, configurationName, pn, dn);
                        AddTypeEFaceConfiguration(doc, configManager, configurationName, dn);
                        AddTypeFFaceConfiguration(doc, configManager, configurationName, pn, dn);

                        // Types G–H — only for Pn10 to Pn40
                        if (pn >= 10 && pn <= 40)
                        {
                            AddTypeGFaceConfiguration(doc, configManager, configurationName, pn, dn);
                            AddTypeHFaceConfiguration(doc, configManager, configurationName, dn);
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(configManager);
                }

                doc.EditRebuild3();

                SetPlane2Distance(doc, configurationName, dn);

                doc.EditRebuild3();
            }

            doc.EditRebuild3();
        }

        private static int ParsePn(string configName)
        {
            int pnStart = configName.IndexOf("Pn") + 2;
            int pnEnd = configName.IndexOf("_Dn");
            if (pnStart < 2 || pnEnd < 0) return 0;
            string pnStr = configName.Substring(pnStart, pnEnd - pnStart).Replace("_5", "");
            return int.TryParse(pnStr, out int pn) ? pn : 0;
        }

        private static int ParseDn(string configName)
        {
            int dnStart = configName.IndexOf("_Dn") + 3;
            int dnEnd = configName.IndexOf("_Type_");
            if (dnStart < 3 || dnEnd < 0) return 0;
            string dnStr = configName.Substring(dnStart, dnEnd - dnStart);
            return int.TryParse(dnStr, out int dn) ? dn : 0;
        }

        private static double GetD1(int dn, int pn)
        {
            switch (dn)
            {
                case 10: return (pn == 6) ? 35.0 : 40.0;
                case 15: return (pn == 6) ? 40.0 : 45.0;
                case 20: return (pn == 6) ? 50.0 : 58.0;
                case 25: return (pn == 6) ? 60.0 : 68.0;
                case 32: return (pn == 6) ? 70.0 : 78.0;
                case 40: return (pn == 6) ? 80.0 : 88.0;
                case 50: return (pn == 6) ? 90.0 : 102.0;
                case 65: return (pn == 6) ? 110.0 : 122.0;
                case 80: return (pn == 6) ? 128.0 : 138.0;
                case 100:
                    if (pn == 6) return 148.0;
                    if (pn == 10 || pn == 16) return 158.0;
                    return 162.0;
                case 125:
                    return (pn == 6) ? 178.0 : 188.0;
                case 150:
                    if (pn == 6) return 202.0;
                    if (pn == 10 || pn == 16) return 212.0;
                    return 218.0;
                case 200:
                    if (pn == 6) return 258.0;
                    if (pn == 10 || pn == 16) return 268.0;
                    if (pn == 25) return 278.0;
                    return 285.0;
                case 250:
                    if (pn == 6) return 312.0;
                    if (pn == 10 || pn == 16) return 320.0;
                    if (pn == 25) return 335.0;
                    return 345.0;
                case 300:
                    if (pn == 6) return 365.0;
                    if (pn == 10) return 370.0;
                    if (pn == 16) return 378.0;
                    if (pn == 25) return 395.0;
                    return 410.0;
                case 350:
                    if (pn == 6) return 415.0;
                    if (pn == 10) return 430.0;
                    if (pn == 16) return 438.0;
                    if (pn == 25) return 450.0;
                    return 465.0;
                case 400:
                    if (pn == 6) return 465.0;
                    if (pn == 10) return 482.0;
                    if (pn == 16) return 490.0;
                    if (pn == 25) return 505.0;
                    return 535.0;
                case 450:
                    if (pn == 6) return 520.0;
                    if (pn == 10) return 532.0;
                    if (pn == 16) return 550.0;
                    if (pn == 25) return 555.0;
                    return 560.0;
                case 500:
                    if (pn == 6) return 570.0;
                    if (pn == 10) return 585.0;
                    if (pn == 16) return 610.0;
                    return 615.0;
                case 600:
                    if (pn == 6) return 670.0;
                    if (pn == 10) return 685.0;
                    if (pn == 16) return 725.0;
                    if (pn == 25) return 720.0;
                    return 735.0;
                default: return 0.0;
            }
        }

        private static double GetF1(int dn)
        {
            if (dn <= 32) return 2.0;
            if (dn <= 250) return 3.0;
            if (dn <= 500) return 4.0;
            return 5.0;
        }

        private static double GetF2(int dn)
        {
            if (dn <= 80) return 4.5;
            if (dn <= 300) return 5.0;
            if (dn <= 900) return 5.5;
            return 6.5;
        }

        private static double GetF3(int dn)
        {
            if (dn <= 80) return 4.0;
            if (dn <= 300) return 4.5;
            if (dn <= 900) return 5.0;
            return 6.0;
        }

        private static double GetF4(int dn)
        {
            if (dn <= 80) return 2.0;
            if (dn <= 300) return 2.5;
            if (dn <= 900) return 3.0;
            return 4.0;
        }

        private static double GetW(int dn)
        {
            switch (dn)
            {
                case 10: return 24.0; case 15: return 29.0; case 20: return 36.0;
                case 25: return 43.0; case 32: return 51.0; case 40: return 61.0;
                case 50: return 73.0; case 65: return 95.0; case 80: return 106.0;
                case 100: return 129.0; case 125: return 155.0; case 150: return 183.0;
                case 200: return 239.0; case 250: return 292.0; case 300: return 343.0;
                case 350: return 395.0; case 400: return 447.0; case 450: return 497.0;
                case 500: return 549.0; case 600: return 649.0;
                default: return 0.0;
            }
        }

        private static double GetX(int dn)
        {
            switch (dn)
            {
                case 10: return 34.0; case 15: return 39.0; case 20: return 50.0;
                case 25: return 57.0; case 32: return 65.0; case 40: return 75.0;
                case 50: return 87.0; case 65: return 109.0; case 80: return 120.0;
                case 100: return 149.0; case 125: return 175.0; case 150: return 203.0;
                case 200: return 259.0; case 250: return 312.0; case 300: return 363.0;
                case 350: return 421.0; case 400: return 473.0; case 450: return 523.0;
                case 500: return 575.0; case 600: return 675.0;
                default: return 0.0;
            }
        }

        private static double GetY(int dn)
        {
            switch (dn)
            {
                case 10: return 35.0; case 15: return 40.0; case 20: return 51.0;
                case 25: return 58.0; case 32: return 66.0; case 40: return 76.0;
                case 50: return 88.0; case 65: return 110.0; case 80: return 121.0;
                case 100: return 150.0; case 125: return 176.0; case 150: return 204.0;
                case 200: return 260.0; case 250: return 313.0; case 300: return 364.0;
                case 350: return 422.0; case 400: return 474.0; case 450: return 524.0;
                case 500: return 576.0; case 600: return 676.0;
                default: return 0.0;
            }
        }

        private static double GetZ(int dn)
        {
            switch (dn)
            {
                case 10: return 23.0; case 15: return 28.0; case 20: return 35.0;
                case 25: return 42.0; case 32: return 50.0; case 40: return 60.0;
                case 50: return 72.0; case 65: return 94.0; case 80: return 105.0;
                case 100: return 128.0; case 125: return 154.0; case 150: return 182.0;
                case 200: return 238.0; case 250: return 291.0; case 300: return 342.0;
                case 350: return 394.0; case 400: return 446.0; case 450: return 496.0;
                case 500: return 548.0; case 600: return 648.0;
                default: return 0.0;
            }
        }

        private static double GetAlpha(int dn)
        {
            if (dn <= 80) return 41.0;
            if (dn <= 300) return 32.0;
            return 27.0;
        }

        private static double GetR(int dn)
        {
            if (dn <= 80) return 2.5;
            if (dn <= 300) return 3.0;
            return 3.5;
        }

        private static void SuppressAllFaceCutsForConfig(ModelDoc2 doc, string configName)
        {
            string[] faceCutNames = new[]
            {
                "Type B face cut", "Type C face cut", "Type D face cut",
                "Type E face cut", "Type F face cut", "Type G face cut", "Type H face cut",
            };

            foreach (string name in faceCutNames)
            {
                Feature faceCut = GetFeatureByName(doc, name);
                if (faceCut == null) continue;

                faceCut.SetSuppression2(
                    (int)swFeatureSuppressionAction_e.swSuppressFeature,
                    (int)swInConfigurationOpts_e.swSpecifyConfiguration,
                    configName);

                Marshal.ReleaseComObject(faceCut);
            }
        }

        private static void AddTypeBFaceConfiguration(ModelDoc2 doc, ConfigurationManager configManager, string configName, int pn, int dn)
        {
            string derivedConfigName = $"{configName.Replace("_Type_A", "")}_Type_B";

            configManager.AddConfiguration2(derivedConfigName, null, null, 0, configName, null, false);

            SuppressAllFaceCutsForConfig(doc, derivedConfigName);

            Feature faceCutSketch = GetFeatureByName(doc, "Type B face cut sketch");
            SetDim(faceCutSketch, "D1", GetD1(dn, pn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "F1", GetF1(dn) / 1000.0, derivedConfigName);
            Marshal.ReleaseComObject(faceCutSketch);

            Feature faceCut = GetFeatureByName(doc, "Type B face cut");
            faceCut.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swThisConfiguration, "");
            Marshal.ReleaseComObject(faceCut);

            doc.EditRebuild3();
        }

        private static void AddTypeCFaceConfiguration(ModelDoc2 doc, ConfigurationManager configManager, string configName, int dn)
        {
            string derivedConfigName = $"{configName.Replace("_Type_A", "")}_Type_C";

            configManager.AddConfiguration2(derivedConfigName, null, null, 0, configName, null, false);

            SuppressAllFaceCutsForConfig(doc, derivedConfigName);

            Feature faceCutSketch = GetFeatureByName(doc, "Type C face cut sketch");
            SetDim(faceCutSketch, "X", GetX(dn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "W", GetW(dn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "F2", GetF2(dn) / 1000.0, derivedConfigName);
            Marshal.ReleaseComObject(faceCutSketch);

            Feature faceCut = GetFeatureByName(doc, "Type C face cut");
            faceCut.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swThisConfiguration, "");
            Marshal.ReleaseComObject(faceCut);

            doc.EditRebuild3();
        }

        private static void AddTypeDFaceConfiguration(ModelDoc2 doc, ConfigurationManager configManager, string configName, int pn, int dn)
        {
            string derivedConfigName = $"{configName.Replace("_Type_A", "")}_Type_D";

            configManager.AddConfiguration2(derivedConfigName, null, null, 0, configName, null, false);

            SuppressAllFaceCutsForConfig(doc, derivedConfigName);

            Feature faceCutSketch = GetFeatureByName(doc, "Type D face cut sketch");
            SetDim(faceCutSketch, "F1", GetF1(dn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "F3", GetF3(dn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "Z", GetZ(dn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "Y", GetY(dn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "D1", GetD1(dn, pn) / 1000.0, derivedConfigName);
            Marshal.ReleaseComObject(faceCutSketch);

            Feature faceCut = GetFeatureByName(doc, "Type D face cut");
            faceCut.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swThisConfiguration, "");
            Marshal.ReleaseComObject(faceCut);

            doc.EditRebuild3();
        }

        private static void AddTypeEFaceConfiguration(ModelDoc2 doc, ConfigurationManager configManager, string configName, int dn)
        {
            string derivedConfigName = $"{configName.Replace("_Type_A", "")}_Type_E";

            configManager.AddConfiguration2(derivedConfigName, null, null, 0, configName, null, false);

            SuppressAllFaceCutsForConfig(doc, derivedConfigName);

            Feature faceCutSketch = GetFeatureByName(doc, "Type E face cut sketch");
            SetDim(faceCutSketch, "F2", GetF2(dn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "X", GetX(dn) / 1000.0, derivedConfigName);
            Marshal.ReleaseComObject(faceCutSketch);

            Feature faceCut = GetFeatureByName(doc, "Type E face cut");
            faceCut.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swThisConfiguration, "");
            Marshal.ReleaseComObject(faceCut);

            doc.EditRebuild3();
        }

        private static void AddTypeFFaceConfiguration(ModelDoc2 doc, ConfigurationManager configManager, string configName, int pn, int dn)
        {
            string derivedConfigName = $"{configName.Replace("_Type_A", "")}_Type_F";

            configManager.AddConfiguration2(derivedConfigName, null, null, 0, configName, null, false);

            SuppressAllFaceCutsForConfig(doc, derivedConfigName);

            Feature faceCutSketch = GetFeatureByName(doc, "Type F face cut sketch");
            SetDim(faceCutSketch, "F3", GetF3(dn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "F1", GetF1(dn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "Y", GetY(dn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "D1", GetD1(dn, pn) / 1000.0, derivedConfigName);
            Marshal.ReleaseComObject(faceCutSketch);

            Feature faceCut = GetFeatureByName(doc, "Type F face cut");
            faceCut.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swThisConfiguration, "");
            Marshal.ReleaseComObject(faceCut);

            doc.EditRebuild3();
        }

        private static void AddTypeGFaceConfiguration(ModelDoc2 doc, ConfigurationManager configManager, string configName, int pn, int dn)
        {
            string derivedConfigName = $"{configName.Replace("_Type_A", "")}_Type_G";

            configManager.AddConfiguration2(derivedConfigName, null, null, 0, configName, null, false);

            SuppressAllFaceCutsForConfig(doc, derivedConfigName);

            Feature faceCutSketch = GetFeatureByName(doc, "Type G face cut sketch");
            SetDim(faceCutSketch, "F2", GetF2(dn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "F1", GetF1(dn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "D1", GetD1(dn, pn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "W", GetW(dn) / 1000.0, derivedConfigName);
            Marshal.ReleaseComObject(faceCutSketch);

            Feature faceCut = GetFeatureByName(doc, "Type G face cut");
            faceCut.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swThisConfiguration, "");
            Marshal.ReleaseComObject(faceCut);

            doc.EditRebuild3();
        }

        private static void AddTypeHFaceConfiguration(ModelDoc2 doc, ConfigurationManager configManager, string configName, int dn)
        {
            string derivedConfigName = $"{configName.Replace("_Type_A", "")}_Type_H";

            configManager.AddConfiguration2(derivedConfigName, null, null, 0, configName, null, false);

            SuppressAllFaceCutsForConfig(doc, derivedConfigName);

            Feature faceCutSketch = GetFeatureByName(doc, "Type H face cut sketch");
            SetDim(faceCutSketch, "F3", GetF3(dn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "F4", GetF4(dn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "Y", GetY(dn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "Z", GetZ(dn) / 1000.0, derivedConfigName);
            SetDim(faceCutSketch, "Alpha", GetAlpha(dn) * Math.PI / 180.0, derivedConfigName);
            SetDim(faceCutSketch, "R", GetR(dn) / 1000.0, derivedConfigName);
            Marshal.ReleaseComObject(faceCutSketch);

            Feature faceCut = GetFeatureByName(doc, "Type H face cut");
            faceCut.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swThisConfiguration, "");
            Marshal.ReleaseComObject(faceCut);

            doc.EditRebuild3();
        }

        private static void SetPlane2Distance(ModelDoc2 doc, string typeAConfigName, int dn)
        {
            double c4 = GetC4(doc, typeAConfigName);
            double f3 = GetF3(dn) / 1000.0;
            double f4 = GetF4(dn) / 1000.0;
            string baseConfigName = typeAConfigName.Replace("_Type_A", "");

            Feature plane2 = GetFeatureByName(doc, "Plane2");

            SetDim(plane2, "Distance", c4, typeAConfigName);

            foreach (string suffix in new[] { "_Type_B", "_Type_C", "_Type_E", "_Type_G" })
            {
                string derivedName = baseConfigName + suffix;
                if (doc.GetConfigurationByName(derivedName) != null)
                    SetDim(plane2, "Distance", c4, derivedName);
            }

            foreach (string suffix in new[] { "_Type_D", "_Type_F" })
            {
                string derivedName = baseConfigName + suffix;
                if (doc.GetConfigurationByName(derivedName) != null)
                    SetDim(plane2, "Distance", c4 - f3, derivedName);
            }

            string typeH = baseConfigName + "_Type_H";
            if (doc.GetConfigurationByName(typeH) != null)
                SetDim(plane2, "Distance", c4 - f4, typeH);

            Marshal.ReleaseComObject(plane2);
        }

        private static double GetC4(ModelDoc2 doc, string configName)
        {
            Feature flangeSketch = GetFeatureByName(doc, "Flange sketch");
            Dimension dim = flangeSketch.Parameter("C4");
            double value = ((double[])dim.GetSystemValue3((int)swInConfigurationOpts_e.swSpecifyConfiguration, configName))[0];
            Marshal.ReleaseComObject(dim);
            Marshal.ReleaseComObject(flangeSketch);
            return value;
        }

        private static Feature GetFeatureByName(ModelDoc2 doc, string name)
        {
            Feature loopFeature = doc.IFirstFeature();
            while (loopFeature != null)
            {
                if (loopFeature.Name == name)
                    return loopFeature;

                Feature nextFeature = (Feature)loopFeature.GetNextFeature();
                Marshal.ReleaseComObject(loopFeature);
                loopFeature = nextFeature;
            }
            return null;
        }

        private static void SetDim(Feature sketch, string paramName, double value, string config)
        {
            Dimension dim = sketch.Parameter(paramName);
            if (dim == null) return;
            dim.SetSystemValue3(value, (int)swSetValueInConfiguration_e.swSetValue_InSpecificConfigurations, config);
            Marshal.ReleaseComObject(dim);
        }

        private static void SetConfig(
            ModelDoc2 doc,
            Feature flangeSketch,
            Feature boltHoleSketch,
            Feature boltHolesPatternFeature,
            string config,
            double D, double C4, double K, double L,
            int boltCount)
        {
            Dimension dimD = flangeSketch.Parameter("D");
            double currentD = ((double[])dimD.GetSystemValue3((int)swInConfigurationOpts_e.swSpecifyConfiguration, config))[0];
            Marshal.ReleaseComObject(dimD);

            double radians = (360.0 / (boltCount * 2)) * Math.PI / 180.0;

            if (currentD > D / 1000.0)
            {
                SetDim(boltHolesPatternFeature, "BoltingNumber", boltCount, config);
                SetDim(boltHoleSketch, "L", L / 1000.0, config);
                SetDim(boltHoleSketch, "Angle", radians, config);
                SetDim(boltHoleSketch, "K", K / 1000.0, config);
                SetDim(flangeSketch, "D", D / 1000.0, config);
                SetDim(flangeSketch, "C4", C4 / 1000.0, config);
            }
            else
            {
                SetDim(flangeSketch, "D", D / 1000.0, config);
                SetDim(flangeSketch, "C4", C4 / 1000.0, config);
                SetDim(boltHoleSketch, "K", K / 1000.0, config);
                SetDim(boltHoleSketch, "Angle", radians, config);
                SetDim(boltHoleSketch, "L", L / 1000.0, config);
                SetDim(boltHolesPatternFeature, "BoltingNumber", boltCount, config);
            }

            doc.EditRebuild3();
        }
    }
}

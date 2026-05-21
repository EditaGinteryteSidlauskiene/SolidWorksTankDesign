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
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn10_Type_A",  D: 75.0,  C1: 12.0, K: 50.0,  L: 11.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn15_Type_A",  D: 80.0,  C1: 12.0, K: 55.0,  L: 11.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn20_Type_A",  D: 90.0,  C1: 14.0, K: 65.0,  L: 11.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn25_Type_A",  D: 100.0, C1: 14.0, K: 75.0,  L: 11.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn32_Type_A",  D: 120.0, C1: 14.0, K: 90.0,  L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn40_Type_A",  D: 130.0, C1: 14.0, K: 100.0, L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn50_Type_A",  D: 140.0, C1: 14.0, K: 110.0, L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn65_Type_A",  D: 160.0, C1: 14.0, K: 130.0, L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn80_Type_A",  D: 190.0, C1: 16.0, K: 150.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn100_Type_A", D: 210.0, C1: 16.0, K: 170.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn125_Type_A", D: 240.0, C1: 18.0, K: 200.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn150_Type_A", D: 265.0, C1: 18.0, K: 225.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn200_Type_A", D: 320.0, C1: 20.0, K: 280.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn250_Type_A", D: 375.0, C1: 22.0, K: 335.0, L: 18.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn300_Type_A", D: 440.0, C1: 22.0, K: 395.0, L: 22.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn350_Type_A", D: 490.0, C1: 22.0, K: 445.0, L: 22.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn400_Type_A", D: 540.0, C1: 22.0, K: 495.0, L: 22.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn450_Type_A", D: 595.0, C1: 24.0, K: 550.0, L: 22.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn500_Type_A", D: 645.0, C1: 24.0, K: 600.0, L: 22.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn600_Type_A", D: 755.0, C1: 30.0, K: 705.0, L: 26.0, boltCount: 20);

            // ── Pn10 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn200_Type_A", D: 340.0, C1: 24.0, K: 295.0, L: 22.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn250_Type_A", D: 395.0, C1: 26.0, K: 350.0, L: 22.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn300_Type_A", D: 445.0, C1: 26.0, K: 400.0, L: 22.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn350_Type_A", D: 505.0, C1: 26.0, K: 460.0, L: 22.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn400_Type_A", D: 565.0, C1: 26.0, K: 515.0, L: 26.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn450_Type_A", D: 615.0, C1: 28.0, K: 565.0, L: 26.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn500_Type_A", D: 670.0, C1: 28.0, K: 620.0, L: 26.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn600_Type_A", D: 780.0, C1: 34.0, K: 725.0, L: 30.0, boltCount: 20);

            // ── Pn16 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn50_Type_A",  D: 165.0, C1: 18.0, K: 125.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn65_Type_A",  D: 185.0, C1: 18.0, K: 145.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn80_Type_A",  D: 200.0, C1: 20.0, K: 160.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn100_Type_A", D: 220.0, C1: 20.0, K: 180.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn125_Type_A", D: 250.0, C1: 22.0, K: 210.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn150_Type_A", D: 285.0, C1: 22.0, K: 240.0, L: 22.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn200_Type_A", D: 340.0, C1: 24.0, K: 295.0, L: 22.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn250_Type_A", D: 405.0, C1: 26.0, K: 355.0, L: 26.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn300_Type_A", D: 460.0, C1: 28.0, K: 410.0, L: 26.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn350_Type_A", D: 520.0, C1: 30.0, K: 470.0, L: 26.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn400_Type_A", D: 580.0, C1: 32.0, K: 525.0, L: 30.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn450_Type_A", D: 640.0, C1: 40.0, K: 585.0, L: 30.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn500_Type_A", D: 715.0, C1: 44.0, K: 650.0, L: 33.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn600_Type_A", D: 840.0, C1: 54.0, K: 770.0, L: 36.0, boltCount: 20);

            // ── Pn25 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn200_Type_A", D: 360.0, C1: 30.0, K: 310.0, L: 26.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn250_Type_A", D: 425.0, C1: 32.0, K: 370.0, L: 30.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn300_Type_A", D: 485.0, C1: 34.0, K: 430.0, L: 30.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn350_Type_A", D: 555.0, C1: 38.0, K: 490.0, L: 33.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn400_Type_A", D: 620.0, C1: 40.0, K: 550.0, L: 36.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn450_Type_A", D: 670.0, C1: 50.0, K: 600.0, L: 36.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn500_Type_A", D: 730.0, C1: 51.0, K: 660.0, L: 36.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn600_Type_A", D: 845.0, C1: 66.0, K: 770.0, L: 39.0, boltCount: 20);

            // ── Pn40 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn10_Type_A",  D: 90.0,  C1: 16.0, K: 60.0,  L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn15_Type_A",  D: 95.0,  C1: 16.0, K: 65.0,  L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn20_Type_A",  D: 105.0, C1: 18.0, K: 75.0,  L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn25_Type_A",  D: 115.0, C1: 18.0, K: 85.0,  L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn32_Type_A",  D: 140.0, C1: 18.0, K: 100.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn40_Type_A",  D: 150.0, C1: 18.0, K: 110.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn50_Type_A",  D: 165.0, C1: 20.0, K: 125.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn65_Type_A",  D: 185.0, C1: 22.0, K: 145.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn80_Type_A",  D: 200.0, C1: 24.0, K: 160.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn100_Type_A", D: 235.0, C1: 24.0, K: 190.0, L: 22.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn125_Type_A", D: 270.0, C1: 26.0, K: 220.0, L: 26.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn150_Type_A", D: 300.0, C1: 28.0, K: 250.0, L: 26.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn200_Type_A", D: 375.0, C1: 36.0, K: 320.0, L: 30.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn250_Type_A", D: 450.0, C1: 38.0, K: 385.0, L: 33.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn300_Type_A", D: 515.0, C1: 42.0, K: 450.0, L: 33.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn350_Type_A", D: 580.0, C1: 46.0, K: 510.0, L: 36.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn400_Type_A", D: 660.0, C1: 50.0, K: 585.0, L: 39.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn450_Type_A", D: 685.0, C1: 57.0, K: 610.0, L: 39.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn500_Type_A", D: 755.0, C1: 57.0, K: 670.0, L: 42.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn600_Type_A", D: 890.0, C1: 72.0, K: 795.0, L: 48.0, boltCount: 20);

            // ── Pn63 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn50_Type_A",  D: 180.0, C1: 26.0, K: 135.0, L: 22.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn65_Type_A",  D: 205.0, C1: 26.0, K: 160.0, L: 22.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn80_Type_A",  D: 215.0, C1: 28.0, K: 170.0, L: 22.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn100_Type_A", D: 250.0, C1: 30.0, K: 200.0, L: 26.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn125_Type_A", D: 295.0, C1: 34.0, K: 240.0, L: 30.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn150_Type_A", D: 345.0, C1: 36.0, K: 280.0, L: 33.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn200_Type_A", D: 415.0, C1: 42.0, K: 345.0, L: 36.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn250_Type_A", D: 470.0, C1: 46.0, K: 400.0, L: 36.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn300_Type_A", D: 530.0, C1: 52.0, K: 460.0, L: 36.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn350_Type_A", D: 600.0, C1: 56.0, K: 525.0, L: 39.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn400_Type_A", D: 670.0, C1: 60.0, K: 585.0, L: 42.0, boltCount: 16);

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
            double D, double C1, double K, double L,
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
                SetDim(flangeSketch, "C1", C1 / 1000.0, config);
            }
            else
            {
                SetDim(flangeSketch, "D", D / 1000.0, config);
                SetDim(flangeSketch, "C1", C1 / 1000.0, config);
                SetDim(boltHoleSketch, "K", K / 1000.0, config);
                SetDim(boltHoleSketch, "Angle", radians, config);
                SetDim(boltHoleSketch, "L", L / 1000.0, config);
                SetDim(boltHolesPatternFeature, "BoltingNumber", boltCount, config);
            }

            doc.EditRebuild3();
        }
    }
}

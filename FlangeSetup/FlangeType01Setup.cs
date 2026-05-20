using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FlangeSetup
{
    public static class FlangeType01Setup
    {
        public static void SuppressAllFaceCuts(ModelDoc2 doc)
        {
            string[] faceCutNames = new[]
            {
                "Type B face cut",
                "Type C face cut",
                "Type D face cut",
                "Type E face cut",
                "Type F face cut",
                "Type G face cut",
                "Type H face cut",
            };

            foreach (string featureName in faceCutNames)
            {
                Feature faceCut = GetFeatureByName(doc, featureName);
                if (faceCut == null) continue;

                faceCut.SetSuppression2(
                    (int)swFeatureSuppressionAction_e.swSuppressFeature,
                    (int)swInConfigurationOpts_e.swAllConfiguration,
                    null);
            }

            doc.EditRebuild3();
        }

        public static void AddConfigurations(ModelDoc2 doc)
        {
            var configurations = new[]
            {
                // PN 6 — DN 10..600
                "Pn6_Dn10",  "Pn6_Dn15", "Pn6_Dn20",  "Pn6_Dn25",  "Pn6_Dn32",  
                "Pn6_Dn40",  "Pn6_Dn50", "Pn6_Dn65",  "Pn6_Dn80",  "Pn6_Dn100", 
                "Pn6_Dn125", "Pn6_Dn150", "Pn6_Dn200", "Pn6_Dn250", "Pn6_Dn300", 
                "Pn6_Dn350", "Pn6_Dn400", "Pn6_Dn450", "Pn6_Dn500", "Pn6_Dn600",

                // PN 10 — DN 200..600
                "Pn10_Dn200", "Pn10_Dn250", "Pn10_Dn300", "Pn10_Dn350", "Pn10_Dn400",
                "Pn10_Dn450", "Pn10_Dn500", "Pn10_Dn600",

                // PN 16 — DN 50..600
                "Pn16_Dn50",  "Pn16_Dn65",  "Pn16_Dn80",  "Pn16_Dn100", "Pn16_Dn125",
                "Pn16_Dn150", "Pn16_Dn200", "Pn16_Dn250", "Pn16_Dn300", "Pn16_Dn350",
                "Pn16_Dn400", "Pn16_Dn450", "Pn16_Dn500", "Pn16_Dn600",

                // PN 25 — DN 200..600
                "Pn25_Dn200", "Pn25_Dn250", "Pn25_Dn300", "Pn25_Dn350", "Pn25_Dn400",
                "Pn25_Dn450", "Pn25_Dn500", "Pn25_Dn600",

                // PN 40 — DN 10..400
                "Pn40_Dn10",  "Pn40_Dn15",  "Pn40_Dn20",  "Pn40_Dn25",  "Pn40_Dn32",
                "Pn40_Dn40",  "Pn40_Dn50",  "Pn40_Dn65",  "Pn40_Dn80",  "Pn40_Dn100",
                "Pn40_Dn125", "Pn40_Dn150", "Pn40_Dn200", "Pn40_Dn250", "Pn40_Dn300",
                "Pn40_Dn350", "Pn40_Dn400",

                // PN 63 — DN 50..400
                "Pn63_Dn50",  "Pn63_Dn65",  "Pn63_Dn80",  "Pn63_Dn100", "Pn63_Dn125",
                "Pn63_Dn150", "Pn63_Dn200", "Pn63_Dn250", "Pn63_Dn300", "Pn63_Dn350",
                "Pn63_Dn400",


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

        public static void DeleteDerivedConfigurations(ModelDoc2 doc)
        {
            string[] allNames = doc.GetConfigurationNames();

            foreach (string configName in allNames)
            {
                if (configName.EndsWith("_Type_A") || !configName.Contains("_Type_"))
                    continue;

                doc.DeleteConfiguration2(configName);
            }

            doc.EditRebuild3();
        }

        public static void UnsuppressConfigurations(ModelDoc2 doc)
        {
            string[] allNames = doc.GetConfigurationNames();

            string[] typeSuffixes = new[] { "_Type_B", "_Type_C", "_Type_D", "_Type_E", "_Type_F", "_Type_G", "_Type_H" };

            foreach (string configName in allNames)
            {
                string matchedSuffix = null;
                foreach (string suffix in typeSuffixes)
                {
                    if (configName.EndsWith(suffix))
                    {
                        matchedSuffix = suffix;
                        break;
                    }
                }

                if (matchedSuffix == null) continue;

                foreach (string suffix in typeSuffixes)
                {
                    // e.g. "_Type_B" → "Type B face cut"
                    string faceCutName = suffix.Substring(1).Replace("_", " ") + " face cut";

                    Feature faceCut = GetFeatureByName(doc, faceCutName);
                    if (faceCut == null) continue;

                    bool isMatch = suffix == matchedSuffix;

                    bool isCurrentlySuppressed = faceCut.IsSuppressed2(
                        (int)swInConfigurationOpts_e.swSpecifyConfiguration,
                        new string[] { configName }) is bool[] states
                        && states.Length > 0 && states[0];

                    if (isMatch && isCurrentlySuppressed)
                    {
                        faceCut.SetSuppression2(
                            (int)swFeatureSuppressionAction_e.swUnSuppressFeature,
                            (int)swInConfigurationOpts_e.swSpecifyConfiguration,
                            configName);
                    }
                    else if (!isMatch && !isCurrentlySuppressed)
                    {
                        faceCut.SetSuppression2(
                            (int)swFeatureSuppressionAction_e.swSuppressFeature,
                            (int)swInConfigurationOpts_e.swSpecifyConfiguration,
                            configName);
                    }

                    Marshal.ReleaseComObject(faceCut);
                }
            }

            doc.EditRebuild3();
        }

        public static void ChangeDimensions(ModelDoc2 doc)
        {
            // Remember the active configuration so we can restore it afterwards
            string originalConfig = ((Configuration)doc.GetActiveConfiguration()).Name;

            Feature flangeSketch            = GetFeatureByName(doc, "Flange sketch");
            Feature boreCutSketch           = GetFeatureByName(doc, "Bore cut sketch");
            Feature boltHoleSketch          = GetFeatureByName(doc, "Bolt hole sketch");
            Feature boltHolesPatternFeature = GetFeatureByName(doc, "Bolt hole pattern");
            // ── Pn2_5 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn2_5_Dn1200", D: 1375.0, C1: 60.0, B1: 616.5, K: 1320.0, L: 30.0, boltCount: 32);

            // ── Pn6 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn10", D: 75.0, C1: 12.0, B1: 18.0, K: 50.0, L: 11.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn15", D: 80.0, C1: 12.0, B1: 22.0, K: 55.0, L: 11.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn20", D: 90.0, C1: 14.0, B1: 27.5, K: 65.0, L: 11.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn25", D: 100.0, C1: 14.0, B1: 34.5, K: 75.0, L: 11.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn32", D: 120.0, C1: 16.0, B1: 43.5, K: 90.0, L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn40", D: 130.0, C1: 16.0, B1: 49.5, K: 100.0, L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn50", D: 140.0, C1: 16.0, B1: 61.5, K: 110.0, L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn65", D: 160.0, C1: 16.0, B1: 77.5, K: 130.0, L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn80", D: 190.0, C1: 18.0, B1: 90.5, K: 150.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn100", D: 210.0, C1: 18.0, B1: 116.0, K: 170.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn125", D: 240.0, C1: 20.0, B1: 141.5, K: 200.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn150", D: 265.0, C1: 20.0, B1: 170.5, K: 225.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn200", D: 320.0, C1: 22.0, B1: 221.5, K: 280.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn250", D: 375.0, C1: 24.0, B1: 276.5, K: 335.0, L: 18.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn300", D: 440.0, C1: 24.0, B1: 327.5, K: 395.0, L: 22.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn350", D: 490.0, C1: 26.0, B1: 359.5, K: 445.0, L: 22.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn400", D: 540.0, C1: 28.0, B1: 411.0, K: 495.0, L: 22.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn450", D: 595.0, C1: 30.0, B1: 462.0, K: 550.0, L: 22.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn500", D: 645.0, C1: 30.0, B1: 513.5, K: 600.0, L: 22.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn600", D: 755.0, C1: 32.0, B1: 616.5, K: 705.0, L: 26.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn700", D: 860.0, C1: 40.0, B1: 616.5, K: 810.0, L: 26.0, boltCount: 24);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn800", D: 975.0, C1: 44.0, B1: 616.5, K: 920.0, L: 30.0, boltCount: 24);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn900", D: 1075.0, C1: 48.0, B1: 616.5, K: 1020.0, L: 30.0, boltCount: 24);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn1000", D: 1175.0, C1: 52.0, B1: 616.5, K: 1120.0, L: 30.0, boltCount: 28);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn1200", D: 1405.0, C1: 60.0, B1: 616.5, K: 1340.0, L: 33.0, boltCount: 32);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn1400", D: 1630.0, C1: 72.0, B1: 616.5, K: 1560.0, L: 36.0, boltCount: 36);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn1600", D: 1830.0, C1: 80.0, B1: 616.5, K: 1760.0, L: 36.0, boltCount: 40);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn1800", D: 2045.0, C1: 88.0, B1: 616.5, K: 1970.0, L: 39.0, boltCount: 44);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn6_Dn2000", D: 2265.0, C1: 96.0, B1: 616.5, K: 2180.0, L: 42.0, boltCount: 48);

            // ── Pn10 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn200", D: 340.0, C1: 24.0, B1: 221.5, K: 295.0, L: 22.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn250", D: 395.0, C1: 26.0, B1: 276.5, K: 350.0, L: 22.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn300", D: 445.0, C1: 26.0, B1: 327.5, K: 400.0, L: 22.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn350", D: 505.0, C1: 26.0, B1: 359.5, K: 460.0, L: 22.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn400", D: 565.0, C1: 26.0, B1: 411.0, K: 515.0, L: 26.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn450", D: 615.0, C1: 28.0, B1: 462.0, K: 565.0, L: 26.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn500", D: 670.0, C1: 28.0, B1: 513.5, K: 620.0, L: 26.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn600", D: 780.0, C1: 30.0, B1: 616.5, K: 725.0, L: 30.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn700", D: 895.0, C1: 35.0, B1: 616.5, K: 840.0, L: 30.0, boltCount: 24);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn800", D: 1015.0, C1: 38.0, B1: 616.5, K: 950.0, L: 33.0, boltCount: 24);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn900", D: 1115.0, C1: 38.0, B1: 616.5, K: 1050.0, L: 33.0, boltCount: 28);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn1000", D: 1230.0, C1: 70.0, B1: 616.5, K: 1160.0, L: 36.0, boltCount: 28);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn10_Dn1200", D: 1455.0, C1: 55.0, B1: 616.5, K: 1380.0, L: 39.0, boltCount: 32);

            // ── Pn16 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn50", D: 165.0, C1: 20.0, B1: 61.5, K: 125.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn65", D: 185.0, C1: 20.0, B1: 77.5, K: 145.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn80", D: 200.0, C1: 20.0, B1: 90.5, K: 160.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn100", D: 220.0, C1: 22.0, B1: 116.0, K: 180.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn125", D: 250.0, C1: 22.0, B1: 141.5, K: 210.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn150", D: 285.0, C1: 24.0, B1: 170.5, K: 240.0, L: 22.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn200", D: 340.0, C1: 26.0, B1: 221.5, K: 295.0, L: 22.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn250", D: 405.0, C1: 29.0, B1: 276.5, K: 355.0, L: 26.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn300", D: 460.0, C1: 32.0, B1: 327.5, K: 410.0, L: 26.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn350", D: 520.0, C1: 35.0, B1: 359.5, K: 470.0, L: 26.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn400", D: 580.0, C1: 38.0, B1: 411.0, K: 525.0, L: 30.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn450", D: 640.0, C1: 42.0, B1: 462.0, K: 585.0, L: 30.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn500", D: 715.0, C1: 46.0, B1: 513.5, K: 650.0, L: 33.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn600", D: 840.0, C1: 55.0, B1: 616.5, K: 770.0, L: 36.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn700", D: 910.0, C1: 63.0, B1: 616.5, K: 840.0, L: 36.0, boltCount: 24);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn800", D: 1025.0, C1: 74.0, B1: 616.5, K: 950.0, L: 39.0, boltCount: 24);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn900", D: 1125.0, C1: 82.0, B1: 616.5, K: 1050.0, L: 39.0, boltCount: 28);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn16_Dn1000", D: 1255.0, C1: 90.0, B1: 616.5, K: 1170.0, L: 42.0, boltCount: 28);

            // ── Pn25 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn200", D: 360.0, C1: 32.0, B1: 221.5, K: 310.0, L: 26.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn250", D: 425.0, C1: 35.0, B1: 276.5, K: 370.0, L: 30.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn300", D: 485.0, C1: 38.0, B1: 327.5, K: 430.0, L: 30.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn350", D: 555.0, C1: 42.0, B1: 359.5, K: 490.0, L: 33.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn400", D: 620.0, C1: 48.0, B1: 411.0, K: 550.0, L: 36.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn450", D: 670.0, C1: 54.0, B1: 462.0, K: 600.0, L: 36.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn500", D: 730.0, C1: 58.0, B1: 513.5, K: 660.0, L: 36.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn600", D: 845.0, C1: 68.0, B1: 616.5, K: 770.0, L: 39.0, boltCount: 20);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn700", D: 960.0, C1: 85.0, B1: 616.5, K: 875.0, L: 42.0, boltCount: 24);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn25_Dn800", D: 1085.0, C1: 95.0, B1: 616.5, K: 990.0, L: 48.0, boltCount: 24);

            // ── Pn40 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn10", D: 90.0, C1: 14.0, B1: 18.0, K: 60.0, L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn15", D: 95.0, C1: 14.0, B1: 22.0, K: 65.0, L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn20", D: 105.0, C1: 16.0, B1: 27.5, K: 75.0, L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn25", D: 115.0, C1: 16.0, B1: 34.5, K: 85.0, L: 14.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn32", D: 140.0, C1: 18.0, B1: 43.5, K: 100.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn40", D: 150.0, C1: 18.0, B1: 49.5, K: 110.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn50", D: 165.0, C1: 20.0, B1: 61.5, K: 125.0, L: 18.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn65", D: 185.0, C1: 22.0, B1: 77.5, K: 145.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn80", D: 200.0, C1: 24.0, B1: 90.5, K: 160.0, L: 18.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn100", D: 235.0, C1: 26.0, B1: 116.0, K: 190.0, L: 22.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn125", D: 270.0, C1: 28.0, B1: 141.5, K: 220.0, L: 26.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn150", D: 300.0, C1: 30.0, B1: 170.5, K: 250.0, L: 26.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn200", D: 375.0, C1: 36.0, B1: 221.5, K: 320.0, L: 30.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn250", D: 450.0, C1: 42.0, B1: 276.5, K: 385.0, L: 33.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn300", D: 515.0, C1: 52.0, B1: 327.5, K: 450.0, L: 33.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn350", D: 580.0, C1: 58.0, B1: 359.5, K: 510.0, L: 36.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn40_Dn400", D: 660.0, C1: 65.0, B1: 411.0, K: 585.0, L: 39.0, boltCount: 16);

            // ── Pn63 ──────────────────────────────────────────────────────
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn50", D: 180.0, C1: 26.0, B1: 61.5, K: 135.0, L: 22.0, boltCount: 4);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn65", D: 205.0, C1: 26.0, B1: 77.5, K: 160.0, L: 22.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn80", D: 215.0, C1: 30.0, B1: 90.5, K: 170.0, L: 22.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn100", D: 250.0, C1: 32.0, B1: 116.0, K: 200.0, L: 26.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn125", D: 295.0, C1: 34.0, B1: 141.5, K: 240.0, L: 30.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn150", D: 345.0, C1: 36.0, B1: 170.5, K: 280.0, L: 33.0, boltCount: 8);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn200", D: 415.0, C1: 48.0, B1: 221.5, K: 345.0, L: 36.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn250", D: 470.0, C1: 55.0, B1: 276.5, K: 400.0, L: 36.0, boltCount: 12);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn300", D: 530.0, C1: 65.0, B1: 327.5, K: 460.0, L: 36.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn350", D: 600.0, C1: 72.0, B1: 359.5, K: 525.0, L: 39.0, boltCount: 16);
            SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature, "Pn63_Dn400", D: 670.0, C1: 80.0, B1: 411.0, K: 585.0, L: 42.0, boltCount: 16);

            // Restore the originally active configuration
            doc.ShowConfiguration2(originalConfig);
            doc.EditRebuild3();
        }

        public static void ChangeDimensionsForFaces(ModelDoc2 doc)
        {
            string[] allNames = doc.GetConfigurationNames();

            string[] typeANames = Array.FindAll(allNames, n => n.EndsWith("_Type_A"));

            for (int i = 20; i < 30 && i < typeANames.Length; i++)
            {
                string configurationName = typeANames[i];

                int pn = ParsePn(configurationName);
                int dn = ParseDn(configurationName);
                if (pn == 0 || dn == 0) continue;

                // Re-acquire a fresh configManager RCW on each iteration to avoid
                // DisconnectedContext errors caused by SolidWorks invalidating the
                // COM context after AddConfiguration2 / SetSuppression2 calls.
                ConfigurationManager configManager = doc.ConfigurationManager;
                try
                {
                    // Type B — all PN
                    AddTypeBFaceConfiguration(doc, configManager, configurationName, pn, dn);

                    // Types C–F — not for Pn2.5 and Pn6
                    if (pn != 2 && pn != 6)
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

                // Set Plane2 Distance for the Type_A config and all derived configs
                SetPlane2Distance(doc, configurationName, dn);

                doc.EditRebuild3();
            }

            doc.EditRebuild3();
        }

        private static int ParsePn(string configName)
        {
            // configName format: Pn6_Dn200_Type_A or Pn2_5_Dn200_Type_A
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
                case 10: return (pn == 2 || pn == 6) ? 35.0 : 40.0;
                case 15: return (pn == 2 || pn == 6) ? 40.0 : 45.0;
                case 20: return (pn == 2 || pn == 6) ? 50.0 : 58.0;
                case 25: return (pn == 2 || pn == 6) ? 60.0 : 68.0;
                case 32: return (pn == 2 || pn == 6) ? 70.0 : 78.0;
                case 40: return (pn == 2 || pn == 6) ? 80.0 : 88.0;
                case 50: return (pn == 2 || pn == 6) ? 90.0 : 102.0;
                case 65: return (pn == 2 || pn == 6) ? 110.0 : 122.0;
                case 80: return (pn == 2 || pn == 6) ? 128.0 : 138.0;
                case 100:
                    if (pn == 2 || pn == 6) return 148.0;
                    if (pn == 10 || pn == 16) return 158.0;
                    return 162.0;
                case 125:
                    return (pn == 2 || pn == 6) ? 178.0 : 188.0;
                case 150:
                    if (pn == 2 || pn == 6) return 202.0;
                    if (pn == 10 || pn == 16) return 212.0;
                    return 218.0;
                case 200:
                    if (pn == 2 || pn == 6) return 258.0;
                    if (pn == 10 || pn == 16) return 268.0;
                    if (pn == 25) return 278.0;
                    return 285.0;
                case 250:
                    if (pn == 2 || pn == 6) return 312.0;
                    if (pn == 10 || pn == 16) return 320.0;
                    if (pn == 25) return 335.0;
                    return 345.0;
                case 300:
                    if (pn == 2 || pn == 6) return 365.0;
                    if (pn == 10) return 370.0;
                    if (pn == 16) return 378.0;
                    if (pn == 25) return 395.0;
                    return 410.0;
                case 350:
                    if (pn == 2 || pn == 6) return 415.0;
                    if (pn == 10) return 430.0;
                    if (pn == 16) return 438.0;
                    if (pn == 25) return 450.0;
                    return 465.0;
                case 400:
                    if (pn == 2 || pn == 6) return 465.0;
                    if (pn == 10) return 482.0;
                    if (pn == 16) return 490.0;
                    if (pn == 25) return 505.0;
                    return 535.0;
                case 450:
                    if (pn == 2 || pn == 6) return 520.0;
                    if (pn == 10) return 532.0;
                    if (pn == 16) return 550.0;
                    if (pn == 25) return 555.0;
                    return 560.0;
                case 500:
                    if (pn == 2 || pn == 6) return 570.0;
                    if (pn == 10) return 585.0;
                    if (pn == 16) return 610.0;
                    return 615.0;
                case 600:
                    if (pn == 2 || pn == 6) return 670.0;
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

        private static void SetConfig(
            ModelDoc2 doc,
            Feature flangeSketch,
            Feature boreCutSketch,
            Feature boltHoleSketch,
            Feature boltHolesPatternFeature,
            string config,
            double D, double C1, double B1, double K, double L,
            int boltCount)
        {
            Dimension dimD = flangeSketch.Parameter("D");
            double currentD = ((double[])dimD.GetSystemValue3((int)swInConfigurationOpts_e.swSpecifyConfiguration, config))[0];
            Marshal.ReleaseComObject(dimD);

            double radians = (360.0 / (boltCount * 2)) * Math.PI / 180.0;

            if (currentD > D / 1000.0)
            {
                SetDim(boreCutSketch, "B1", B1 / 1000.0, config);
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
                SetDim(boreCutSketch, "B1", B1 / 1000.0, config);
            }

            doc.EditRebuild3();
        }

        /// <summary>
        /// Sets the Plane2 Distance dimension for a Type_A configuration and all its derived
        /// face-type configurations. Type_A, B, C, E, G use C1; Type_D and F use C1 - F3;
        /// Type_H uses C1 - F4.
        /// </summary>
        private static void SetPlane2Distance(ModelDoc2 doc, string typeAConfigName, int dn)
        {
            double c1 = GetC1(doc, typeAConfigName);
            double f3 = GetF3(dn) / 1000.0;
            double f4 = GetF4(dn) / 1000.0;
            string baseConfigName = typeAConfigName.Replace("_Type_A", "");

            Feature plane2 = GetFeatureByName(doc, "Plane2");

            // Type A — C1
            SetDim(plane2, "Distance", c1, typeAConfigName);

            // Type B, C, E, G — C1
            foreach (string suffix in new[] { "_Type_B", "_Type_C", "_Type_E", "_Type_G" })
            {
                string derivedName = baseConfigName + suffix;
                if (doc.GetConfigurationByName(derivedName) != null)
                    SetDim(plane2, "Distance", c1, derivedName);
            }

            // Type D, F — C1 - F3
            foreach (string suffix in new[] { "_Type_D", "_Type_F" })
            {
                string derivedName = baseConfigName + suffix;
                if (doc.GetConfigurationByName(derivedName) != null)
                    SetDim(plane2, "Distance", c1 - f3, derivedName);
            }

            // Type H — C1 - F4
            string typeH = baseConfigName + "_Type_H";
            if (doc.GetConfigurationByName(typeH) != null)
                SetDim(plane2, "Distance", c1 - f4, typeH);

            Marshal.ReleaseComObject(plane2);
        }

        /// <summary>
        /// Reads the C1 value from the flange sketch for the given configuration.
        /// </summary>
        private static double GetC1(ModelDoc2 doc, string configName)
        {
            Feature flangeSketch = GetFeatureByName(doc, "Flange sketch");
            Dimension dim = flangeSketch.Parameter("C1");
            double value = ((double[])dim.GetSystemValue3((int)swInConfigurationOpts_e.swSpecifyConfiguration, configName))[0];
            Marshal.ReleaseComObject(dim);
            Marshal.ReleaseComObject(flangeSketch);
            return value;
        }

        /// <summary>
        /// Sets a sketch dimension value for a specific configuration and immediately releases
        /// the COM RCW to prevent DisconnectedContext errors from GC finalizer running on MTA thread.
        /// </summary>
        private static void SetDim(Feature sketch, string paramName, double value, string config)
        {
            Dimension dim = sketch.Parameter(paramName);
            dim.SetSystemValue3(value, (int)swSetValueInConfiguration_e.swSetValue_InSpecificConfigurations, config);
            Marshal.ReleaseComObject(dim);
        }

        private static Feature GetFeatureByName(ModelDoc2 assemblyModelDoc, string name)
        {
            //Starting from the first feature
            Feature loopFeature = assemblyModelDoc.IFirstFeature();

            //Loop features until the requested feature is found
            while (loopFeature != null)
            {
                if (loopFeature.Name == name)
                {
                    return loopFeature;
                }

                //Get next feature and release the current one to avoid RCW accumulation on the MTA finalizer thread
                Feature nextFeature = (Feature)loopFeature.GetNextFeature();
                Marshal.ReleaseComObject(loopFeature);
                loopFeature = nextFeature;
            }

            return null;
        }
    }
}



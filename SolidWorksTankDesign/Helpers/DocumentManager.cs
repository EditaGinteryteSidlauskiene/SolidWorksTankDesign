using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;

namespace SolidWorksTankDesign
{
    internal static  class DocumentManager
    {
        /// <summary>
        /// Updates tanksite assembly attribute, saves and closes subassembly, and saves tank site assembly doc.
        /// </summary>
        /// <param name="subassemblyModelDoc"></param>
        public static void UpdateAndSaveDocuments()
        {
            // Update SW attribute parameter
            TankSiteDataManager.SerializeAndStoreTankSiteAssemblyData();

            //Save and close all subassemblies starting from the lowest in the hierarchy
            while (!ReferenceEquals(SolidWorksDocumentProvider._solidWorksApplication.ActiveDoc, SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc))
            {
                ModelDoc2 subassemblyDoc = SolidWorksDocumentProvider.GetActiveDoc();

                subassemblyDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);

                SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(subassemblyDoc.GetTitle());
            }

            // Save the document of tank site assembly
            SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
        }

        /// <summary>
        /// Packages a SolidWorks assembly document (along with its associated drawings) into  a single, 
        /// timestamped folder. It ensures file uniqueness by incorporating the current 
        /// timestamp into both the folder name and the packed file names. The method returns the full path 
        /// to the packed assembly file (.SLDASM) for further processing or reference.
        /// </summary>
        /// <param name="assemblyModelDoc"></param>
        /// <returns></returns>
        public static string PackAndGo(string projectFolder, ModelDoc2 assemblyModelDoc, string compartmentName, string serialNumber)
        {
            // Get the Pack and Go interface for the assembly document
            PackAndGo packAndGo = assemblyModelDoc.Extension.GetPackAndGo();

            // Configure Pack and Go options
            packAndGo.IncludeDrawings = true;           // Include associated drawings in the Pack and Go
            packAndGo.FlattenToSingleFolder = true;     // Save all files to a single folder (no subfolders)

            // Define the base folder where Pack and Go files will be saved
            string packAndGoFolderPath = projectFolder;

            // Generate a unique folder name using the current timestamp (ticks)
            double ticks = DateTime.Now.Ticks;
            string timestampedPackAndGoFolder = $"{packAndGoFolderPath}\\{ticks}";
            
            if (compartmentName != null && compartmentName != string.Empty)
            {
                packAndGo.SetSaveToName(
                    true,
                    $"{packAndGoFolderPath}\\{ticks.ToString()}_{compartmentName}");

                // Execute the Pack and Go operation
                assemblyModelDoc.Extension.SavePackAndGo(packAndGo);

                // Construct and return the full path to the packed assembly file
                return $"{packAndGoFolderPath}\\{ticks.ToString()}_{compartmentName}\\{assemblyModelDoc.GetTitle()}.SLDASM";
            }

            if(serialNumber != null && serialNumber != string.Empty)
            {
                timestampedPackAndGoFolder = $"{packAndGoFolderPath}\\{serialNumber}";
            }

            // Add a prefix to all Pack and Go file names using the timestamp
            packAndGo.AddPrefix = $"{ticks}_";

            // Set the save location for the Pack and Go files
            packAndGo.SetSaveToName(true, timestampedPackAndGoFolder);

            // Execute the Pack and Go operation
            assemblyModelDoc.Extension.SavePackAndGo(packAndGo);

            // Construct and return the full path to the packed assembly file
            return $"{timestampedPackAndGoFolder}\\{ticks}_{assemblyModelDoc.GetTitle()}.SLDASM";
        }
    }
}

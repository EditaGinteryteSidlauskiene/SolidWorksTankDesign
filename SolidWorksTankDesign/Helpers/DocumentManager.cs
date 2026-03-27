using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

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

            TankSiteDataManager.UpdateTankProperties();

            TankSiteDataManager.UpdateFilesToDeleteList();

            //SaveInitialConfiguration and close all subassemblies starting from the lowest in the hierarchy
            while (!ReferenceEquals(SolidWorksDocumentProvider._solidWorksApplication.ActiveDoc, SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc))
            {
                ModelDoc2 subassemblyDoc = SolidWorksDocumentProvider.GetActiveDoc();

                subassemblyDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);

                SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(subassemblyDoc.GetTitle());
            }

            try
            {
                ModelDoc2 tankSiteDoc = SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc;

                // Force rebuild to propagate assembly-level features (like cut extrudes) to referenced parts
                tankSiteDoc.ForceRebuild3(true);

                // Save all modified documents referenced by the assembly (parts, subassemblies)
                tankSiteDoc.Save3(
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent |
                    (int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced,
                    (int)swFileSaveError_e.swGenericSaveError,
                    (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
            }
            catch (Exception ex) { }
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
            packAndGo.FlattenToSingleFolder = true;     // SaveInitialConfiguration all files to a single folder (no subfolders)

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
                return $"{packAndGoFolderPath}\\{ticks.ToString()}_{compartmentName}\\{assemblyModelDoc.GetTitle()}_{ticks.ToString()}.SLDASM";
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

        /// <summary>
        /// Sets the folder where to store this project's folder.
        /// Default folder's path is stored in Settings.txt
        /// </summary>
        public static string GetProjectsFolder(string settingsFilePath)
        {
            string folderForAllProjects = string.Empty ;
            using (StreamReader reader = new StreamReader(settingsFilePath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (line.StartsWith("DefaultProjectFolder="))
                    {
                        folderForAllProjects = line.Substring("DefaultProjectFolder=".Length);
                        folderForAllProjects = folderForAllProjects.Replace("\"", string.Empty);

                        break;
                    }
                }
            }

            return folderForAllProjects ;
        }

        /// <summary>
        /// Copies the empty tank site assembly document and documents of components that are in tank site assembly.
        /// Documents are copied and renamed to the folder selected by the user
        /// </summary>
        /// <param name="solidWorksApp"></param>
        /// <param name="emptyTankSiteAssemblyDoc"></param>
        public static string CopyDocuments(SldWorks solidWorksApp, ModelDoc2 emptyTankSiteAssemblyDoc, string folderForAllProjects, string serialNumber)
        {
            // Packs all documents that are in emptyTankSiteAssemblyDoc and saves them in a new foler
            string path = DocumentManager.PackAndGo(folderForAllProjects, emptyTankSiteAssemblyDoc, null, serialNumber);

            // Get path of folder, where all documents of the current project will be stored
            string documentDirectory = Path.GetDirectoryName(path);

            // Open just copied and renamed tank site assembly document
            DocumentSpecification documentSpecification = (DocumentSpecification)solidWorksApp.GetOpenDocSpec(path);
            solidWorksApp.OpenDoc7(documentSpecification);

            // Close the primary empty tank site assembly doc
            solidWorksApp.CloseDoc(emptyTankSiteAssemblyDoc.GetTitle());

            return documentDirectory;
        }

        public static void RenameFolderContainingDocument(ModelDoc2 documentInFolder, string oldPart, string newPart)
        {
            if (documentInFolder == null)
            {
                return;
            }

            // Step 2: Get the full path of the current document
            string documentPath = documentInFolder.GetPathName();
            if (string.IsNullOrEmpty(documentPath))
            {
                Console.WriteLine("Document path is invalid.");
                return;
            }

            // Step 3: Get the current folder path and its parent directory
            string currentFolderPath = Path.GetDirectoryName(documentPath);
            string parentDirectory = Path.GetDirectoryName(currentFolderPath);

            // Extract the folder name and replace the desired part
            string folderName = Path.GetFileName(currentFolderPath);
            string newFolderName = folderName.Replace(oldPart, newPart);

            // Build the new folder path
            string newFolderPath = Path.Combine(parentDirectory, newFolderName);

            try
            {
                // Step 4: Rename the folder
                if (Directory.Exists(currentFolderPath) && !Directory.Exists(newFolderPath))
                {
                    Directory.Move(currentFolderPath, newFolderPath);
                    Console.WriteLine($"Folder renamed successfully to: {newFolderPath}");
                }
                else
                {
                    Console.WriteLine("Folder rename failed: New folder name already exists or current folder doesn't exist.");
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error renaming folder: {ex.Message}");
            }
        }

        public static async Task DeleteFiles()
        {
            List<string> pathsToDelete = SolidWorksDocumentProvider._filesToDelete;

            // We schedule the entire deletion loop on a background thread:
            for (int i = 0; i < pathsToDelete.Count; i++)
            {
                try
                {
                    if (!File.Exists(pathsToDelete[i]))
                    {
                        SolidWorksDocumentProvider._filesToDelete.Remove(pathsToDelete[i]);
                    } 

                    else
                    {
                        File.Delete(pathsToDelete[i]);
                        SolidWorksDocumentProvider._filesToDelete.RemoveAt(i);
                    }
                }
                catch (Exception ex) { }
            }

            TankSiteDataManager.UpdateFilesToDeleteList();
        }
    }
}

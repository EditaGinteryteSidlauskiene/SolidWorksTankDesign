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
        /// Packages the nozzle position sketch assembly (and all referenced files such as the
        /// envelope part) into <paramref name="projectFolder"/> using Pack and Go.
        /// All files receive a suffix of <c>{nozzleNumber}_{ticks}</c> so that cross-references
        /// between the assembly and its parts remain valid. No sub-folder is created.
        /// </summary>
        /// <param name="projectFolder">Destination folder for the packed files.</param>
        /// <param name="assemblyModelDoc">The nozzle position sketch assembly to pack.</param>
        /// <param name="nozzleNumber">The nozzle's sequence number used in the output filename.</param>
        /// <returns>Full path to the packed nozzle assembly file (.SLDASM).</returns>
        public static string PackAndGoManhole(string projectFolder, ModelDoc2 assemblyModelDoc, int nozzleNumber)
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

            // Add a prefix to all Pack and Go file names using the timestamp
            packAndGo.AddSuffix = $"{nozzleNumber}_{ticks}";

            // Set the save location for the Pack and Go files
            packAndGo.SetSaveToName(true, packAndGoFolderPath);

            // Execute the Pack and Go operation
            assemblyModelDoc.Extension.SavePackAndGo(packAndGo);

            // Construct and return the full path to the packed assembly file
            return $"{packAndGoFolderPath}\\M{nozzleNumber}_{ticks}.SLDASM";
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
        /// <summary>
        /// Copies the empty tank site assembly and all its referenced documents into a new
        /// project folder using Pack and Go, then opens the copied assembly and closes the
        /// original template.
        /// </summary>
        /// <param name="solidWorksApp">The active SolidWorks application instance.</param>
        /// <param name="emptyTankSiteAssemblyDoc">The template tank site assembly to copy.</param>
        /// <param name="folderForAllProjects">Root folder under which the new project folder is created.</param>
        /// <param name="serialNumber">Serial number used to name the new project folder.</param>
        /// <returns>Full path to the directory where the copied documents were saved.</returns>
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

        /// <summary>
        /// Renames the folder that contains <paramref name="documentInFolder"/> by replacing
        /// <paramref name="oldPart"/> with <paramref name="newPart"/> in the folder name.
        /// </summary>
        /// <param name="documentInFolder">A document whose containing folder should be renamed.</param>
        /// <param name="oldPart">The substring in the folder name to replace.</param>
        /// <param name="newPart">The replacement substring.</param>
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

        /// <summary>
        /// Deletes all files listed in <see cref="SolidWorksDocumentProvider._filesToDelete"/>.
        /// Files that no longer exist on disk are removed from the list silently.
        /// After deletion the stored file list is updated via
        /// <see cref="TankSiteDataManager.UpdateFilesToDeleteList"/>.
        /// </summary>
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

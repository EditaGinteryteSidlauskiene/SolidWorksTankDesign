using System;
using System.IO;
using System.Windows.Forms;

namespace SolidWorksTankDesign.Windows
{
    public partial class SettingsWindow : UserControl
    {
        private string _settingsFilePath;
        public SettingsWindow(string settingsFilePath)
        {
            _settingsFilePath = settingsFilePath;
            InitializeComponent();
        }

        ///// <summary>
        ///// Changes projects folder path in the settings file
        ///// </summary>
        ///// <returns>New projects folder path</returns>
        //private string ChangeProjectsFolderPath()
        //{
        //    try
        //    {
        //        // Get projects folder, selected by the user
        //        string newProjectsFolderPath = ProjectsFolder_BrowserDialog.SelectedPath;

        //        // Get all lines in settings file
        //        string[] settingsLines = File.ReadAllLines(_settingsFilePath);

        //        // Find and update the line with the DefaultProjectFolder
        //        for (int i = 0; i < settingsLines.Length; i++)
        //        {
        //            if (settingsLines[i].StartsWith("DefaultProjectFolder="))
        //            {
        //                settingsLines[i] = "DefaultProjectFolder=" + newProjectsFolderPath;
        //                break;
        //            }
        //        }

        //        // Write the updated lines back to the file
        //        File.WriteAllLines(_settingsFilePath, settingsLines);

        //        return newProjectsFolderPath;
        //    }
        //    catch (Exception ex)
        //    {
        //        // Handle exceptions (e.g., file not found, access denied)
        //        MessageBox.Show($"Error updating settings: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        return null;
        //    }
        //}

        ///// <summary>
        ///// Event, when the projects folder combo box is drop down,
        ///// folder browser window is displayed and projects folder is changed
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void ProjectsFolder_ComboBox_DropDown(object sender, EventArgs e)
        //{
        //    // Display folder browser window
        //    ProjectsFolder_BrowserDialog.ShowDialog();

        //    // Change projects folder path in the settings file and
        //    // assign it to the variable 
        //    string newProjectsFolder = ChangeProjectsFolderPath();
        //    if(newProjectsFolder != null)
        //    {
        //        // Display new projects folder path in the combo box
        //        ProjectsFolder_ComboBox.Text = newProjectsFolder;
        //    }

        //    // Collapse the drop down.
        //    BeginInvoke(new Action(() => ProjectsFolder_ComboBox.DroppedDown = false));
        //}

        ///// <summary>
        ///// Event, makes SettingWindow invisible
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void Close_Button_Click(object sender, EventArgs e)
        //{
        //    this.Visible = false;
        //}
    }
}

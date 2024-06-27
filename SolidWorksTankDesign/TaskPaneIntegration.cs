using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorksTankDesign;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace AddinWithTaskpane
{
    /* To install this addin, make sure that TaskpaneIntegration class is public. There are two options to install:
     1. Install SolidWorksAddinInstaller from Github angelsix https://github.com/angelsix/solidworks-api/tree/develop/Tools/Addin%20Installer
    2. Run Command Prompt as administrator. Go to directory: C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe <this is the path to the dll file (e.g. "C:\Users\Edita\Desktop\SolidWorksTankDesign\SolidWorksTankDesign\bin\Debug\SolidWorksTankDesign.dll")>
    Hit Enter. Check SolidWorks if the UI is dispayed*/
    /// <summary>
    /// This is the main Add-in. This is were you register the add-in, register entries to
    /// RegEdit, and it also responds to SolidWorks telling it to connect and disconnect, and
    /// it creates the taskpane window.
    /// </summary>
    public class TaskpaneIntegration : ISwAddin
    {
        #region Private Members

        // The cookie to the current instance of SolidWorks we are running inside of
        private int mSwCookie;

        // The taskpane view for our add-in
        private TaskpaneView mTaskpaneView;

        // Variable of type of TaskpaneHostUI.cs [Design] control, because we are going to create an instance of it.
        // The UI control that is going to be inside the SolidWorks taskpane view.
        private TaskpaneHostUI mTaskpaneHostUI;

        // Instance of SolidWorks application
        private SldWorks mSolidWorksApplication;

        #endregion

        #region Public Members

        /* When we register our add-in (the thing that's going to appear on the right hand side
       the taskpane, it needs a unique ID and it's some way to differentiate between our add-in from
       every other add-in. This ID needs to be set to the user control.
       The unique Id to the taskpane used for registration in COM*/
        public const string SWTASKPANE_PROGID = "Edita.AddinWithTaskpane.Taskpane";
        public const string SWTASKPANE_PROGID2 = "Edita.AddinWithTaskpane.Taskpane";

        #endregion

        #region SolidWorks Add-in Callbacks

        /// <summary>
        /// Called when SolidWorks has loaded our add-in and wants to do our connection logic
        /// </summary>
        /// <param name="ThisSW">The current SolidWorks instance</param>
        /// <param name="Cookie">The current SolidWorks cookie Id</param>
        /// <returns></returns>
        /// <exception cref="System.NotImplementedException"></exception>
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            // When we connect to SolidWorks, the first thing we want to do is to store reference to the current SolidWorks instance,
            // so we can actually do thing with SolidWorks.
            // This cast the object to what we know is SolidWorks instance, and stores in it. This is where we access SolidWorks from.
            mSolidWorksApplication = (SldWorks)ThisSW;

            //Store the Cookie Id
            mSwCookie = Cookie;

            // Setup callback info
            var ok = mSolidWorksApplication.SetAddinCallbackInfo2(0, this, mSwCookie);

            // Inject our UI, this is where we want to create the taskpane.
            LoadUI();

            //// Attach event handlers
            //EventHandlers.AttachEventHandlers(mSolidWorksApplication);

            //// Passing SW application reference to EventHandlers class.
            //EventHandlers eventHandlers = new EventHandlers();
            //eventHandlers.AssignSWApp(mSolidWorksApplication);


            //Return ok
            return true;
        }

        /// <summary>
        /// Called when SolidWorks is about to unload our add-in and wants to do our disconnection logic
        /// </summary>
        /// <returns></returns>
        /// <exception cref="System.NotImplementedException"></exception>
        public bool DisconnectFromSW()
        {
            // Clean up our UI
            UnloadUI();

            //// Detaching from event handlers
            //EventHandlers.DetachEventHandlers(mSolidWorksApplication);

            // Return ok
            return true;
        }

        #endregion

        #region Create UI

        /// <summary>
        /// Load our Taskpane and our host UI
        /// </summary>
        private void LoadUI()
        {
            // Getting path to the image
            var imagePath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).Replace(@"file:\", string.Empty), "logo-small.bmp");

            mTaskpaneView = mSolidWorksApplication.CreateTaskpaneView2(imagePath, "Welcome");

            // Load the UI into the taskpane. This will tell SolidWorks to go ahead and find the class that's inside this
            // dll  with that id (SWTASKPANE_PROGID), and then inject it into the UI.
            mTaskpaneHostUI = (TaskpaneHostUI)mTaskpaneView.AddControl(TaskpaneIntegration.SWTASKPANE_PROGID, string.Empty);

            // Passing SW application reference to TaskpaneHostUI class.
            mTaskpaneHostUI.AssignSWApp(mSolidWorksApplication);
        }

        /// <summary>
        /// Cleanup the taskpane when we disconnect/unload
        /// </summary>
        /// <exception cref="System.NotImplementedException"></exception>
        private void UnloadUI()
        {
            mTaskpaneHostUI = null;

            // Remove taskpane view
            mTaskpaneView.DeleteView();

            // Release COM reference and cleanup memory
            Marshal.ReleaseComObject(mTaskpaneView);

            // To make sure there is no references we just cleanup our own reference and make sure it's null.
            mTaskpaneView = null;
        }

        #endregion

        #region COM Registration

        /*We need to flag the ComRegister function with this attribute, so when you call RegAsm, 
        it knows where to find this function and what to do.*/
        /// <summary>
        /// The COM registration call to add our registry entries to the SolidWorks add-in registry
        /// </summary>
        /// <param name="t"></param>
        [ComRegisterFunction()]
        private static void ComRegister(Type t)
        {
            // Path to our add-in
            var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

            // Add to the registry. Create our registry folder for the add-in
            using (var rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(keyPath))
            {
                // Load add-in when SolidWorks opens
                rk.SetValue(null, 1);

                // Set SolidWorks add-in title and description
                rk.SetValue("Title", "SwAddin With Taskpane");
                rk.SetValue("Description", "All your pixels are belong to us");
            }
        }

        /// <summary>
        /// The COM unregister call to remove our custom entries we added in the COM register function
        /// </summary>
        /// <param name="t"></param>
        [ComUnregisterFunction()]
        private static void ComUnregister(Type t)
        {
            var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

            // Remove our registry entry
            Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(keyPath);
        }

        #endregion
    }
}
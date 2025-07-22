using AlibreAddOn;
using AlibreX;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using Application = System.Windows.Application;
using MessageBox = System.Windows.MessageBox;

namespace AlibreAddOnAssembly
{
    /// <summary>
    /// Contains all constant string values used throughout the add-on.
    /// </summary>
    public static class AddOnConstants
    {
        // General Add-on Info
        public const string AddOnName = "stltostp";
        public const string InvokeMessage = "stltostp.AddOnInvoke";
        public const string FailedToInitializeRibbonPrefix = "Failed to initialize AddOnRibbon: ";
        public const string AddOnLoadedMessage = "stltostp Add-on has been loaded! 👍";
        public const string AddOnLoadedTitle = "Add-on Loaded";

        // File Dialog and Paths
        public const string OpenFileDialogTitle = "Select STL file to convert";
        public const string StlFileFilter = "STL files (*.stl)|*.stl";
        public const string StepFileExtension = ".stp";
        public const string ToolsDirectoryName = "tools";
        public const string ConverterExeName = "stltostp.exe";
        public const string IconPath = @"Icons\logo.ico";

        // Menu Text and ToolTips
        public const string RootMenuToolTip = "alibre-stltostp-addon";
        public const string SubMenuGroupText = "STL to STP";
        public const string AboutMenuText = "About";
        public const string AboutMenuToolTip = "About stltostp";
        public const string RunMenuText = "Run";
        public const string RunMenuToolTip = "Run stltostp tool";

        // Messages and Titles
        public const string AboutMessage = "About stltostp - Version X.Y.Z\n(Details about the tool or link can be placed here)";
        public const string ConverterMissingTitle = "stltostp.exe missing";
        public const string StlToStepErrorTitle = "STL → STEP Error";
        public const string StlToStepConversionErrorTitle = "STL → STEP Conversion Error";
        public const string ProcessErrorTitle = "STL → STEP Process Error";
        public const string VerifyImportTitle = "Verify STEP Path & File for Import";
        public const string AlibreImportErrorTitle = "Alibre Design Import Error";

        // Process and Error Messages (Formats)
        public const string SessionInfoFormat = "{0} : {1}";
        public const string ConverterNotFoundFormat = "Converter not found at expected location:\n{0}";
        public const string ArgumentFormat = "\"{0}\" \"{1}\"";
        public const string ProcessLaunchFailedMessage = "Failed to launch stltostp.exe. The process could not be started.";
        public const string ConversionSuccessMessage = "stltostp.exe completed successfully.\n\n";
        public const string ImportAttemptMessageFormat = "Attempting to import STEP file from:\n{0}\n\n";
        public const string FileExistsMessage = "File exists: True\n";
        public const string FileSizeMessageFormat = "File size: {0} bytes";
        public const string AlibreRootNotFoundMessage = "Failed to get Alibre Root object. Cannot import STEP file.";
        public const string StepImportFailedFormat = "STEP import failed (Alibre API error):\n{0}\n\nFile was: {1}";
        public const string ConversionFailedFormat = "stltostp.exe failed to convert '{0}'.\n\n";
        public const string ExitCodeFormat = "Exit Code: {0}\n";
        public const string StdOutputFormat = "\nStandard Output:\n{0}\n";
        public const string StdErrorFormat = "\nStandard Error:\n{0}\n";
        public const string OutputFileMissingMessage = "\nThe output STEP file was not found at the expected location:\n";
        public const string ToolExitedWithError = "\nThe tool exited with an error. The STEP file might be incomplete or corrupt even if it exists.";
        public const string UnexpectedProcessErrorFormat = "An unexpected error occurred during the STL to STEP process:\n{0}";
    }

    public static class AlibreAddOn
    {
        private static IADRoot AlibreRoot { get; set; }
        private static IntPtr _parentWinHandle;
        private static AddOnRibbon _AssimpInsideAlibreDesignHandle;

        public static void AddOnLoad(IntPtr hwnd, IAutomationHook pAutomationHook, IntPtr unused)
        {
            // Show the message box when the add-on loads
            MessageBox.Show(AddOnConstants.AddOnLoadedMessage,
                            AddOnConstants.AddOnLoadedTitle,
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);

            // Existing initialization code
            AlibreRoot = (IADRoot)pAutomationHook.Root;
            _parentWinHandle = hwnd;
            _AssimpInsideAlibreDesignHandle = new AddOnRibbon(AlibreRoot, _parentWinHandle);
        }

        public static void AddOnUnload(IntPtr hwnd, bool forceUnload, ref bool cancel, int reserved1, int reserved2)
        {
            _AssimpInsideAlibreDesignHandle = null;
            AlibreRoot = null;
        }

        public static IADRoot GetRoot()
        {
            return AlibreRoot;
        }

        public static void AddOnInvoke(IntPtr hwnd, IntPtr pAutomationHook, string sessionName, bool isLicensed, int reserved1, int reserved2)
        {
            MessageBox.Show(AddOnConstants.InvokeMessage);
        }

        public static IAlibreAddOn GetAddOnInterface()
        {
            return _AssimpInsideAlibreDesignHandle;
        }
    }

    public class AddOnRibbon : IAlibreAddOn
    {
        private readonly MenuManager _menuManager;
        public IADRoot _AlibreRoot;
        public IntPtr _parentWinHandle;

        public AddOnRibbon(IADRoot AlibreRoot, IntPtr parentWinHandle)
        {
            _AlibreRoot = AlibreRoot;
            _parentWinHandle = parentWinHandle;
            try
            {
                _menuManager = new MenuManager(_AlibreRoot.TopmostSession);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(AddOnConstants.FailedToInitializeRibbonPrefix + ex.Message);
            }
        }

        public int RootMenuItem => _menuManager.GetRootMenuItem().Id;

        public IAlibreAddOnCommand? InvokeCommand(int menuID, string sessionIdentifier)
        {
            var session = _AlibreRoot.Sessions.Item(sessionIdentifier);
            var menuItem = _menuManager.GetMenuItemById(menuID);
            return menuItem?.Command?.Invoke(session);
        }

        public bool HasSubMenus(int menuID)
        {
            var menuItem = _menuManager.GetMenuItemById(menuID);
            return menuItem != null && menuItem.SubItems.Count > 0;
        }

        public Array? SubMenuItems(int menuID)
        {
            var menuItem = _menuManager.GetMenuItemById(menuID);
            return menuItem?.SubItems.Select(subItem => subItem.Id).ToArray();
        }

        public string? MenuItemText(int menuID) => _menuManager.GetMenuItemById(menuID)?.Text;

        public ADDONMenuStates MenuItemState(int menuID, string sessionIdentifier) => ADDONMenuStates.ADDON_MENU_ENABLED;

        public string? MenuItemToolTip(int menuID) => _menuManager.GetMenuItemById(menuID)?.ToolTip;

        public string? MenuIcon(int menuID) => _menuManager.GetMenuItemById(menuID)?.Icon;

        public bool PopupMenu(int menuID) => false;

        public bool HasPersistentDataToSave(string sessionIdentifier) => false;

        public void SaveData(IStream pCustomData, string sessionIdentifier) { }

        public void LoadData(IStream pCustomData, string sessionIdentifier) { }

        public bool UseDedicatedRibbonTab() => false;

        private void IAlibreAddOn_setIsAddOnLicensed(bool isLicensed) { }
        void IAlibreAddOn.setIsAddOnLicensed(bool isLicensed) => IAlibreAddOn_setIsAddOnLicensed(isLicensed);
    }

    public class MenuItem
    {
        public int Id { get; set; }
        public string Text { get; set; }
        public string ToolTip { get; set; }
        public string Icon { get; set; }
        public Func<IADSession, IAlibreAddOnCommand?>? Command { get; set; }
        public List<MenuItem> SubItems { get; set; }
        public MenuItem(int id, string text, string toolTip = "", string icon = "", Func<IADSession, IAlibreAddOnCommand?>? command = null)
        {
            Id = id;
            Text = text;
            ToolTip = toolTip;
            Icon = icon;
            Command = command;
            SubItems = new List<MenuItem>();
        }

        public void AddSubItem(MenuItem subItem) => SubItems.Add(subItem);
        public IAlibreAddOnCommand? DummyFunction(IADSession session)
        {
            MessageBox.Show(string.Format(AddOnConstants.SessionInfoFormat, session.Name, session.FilePath));
            return null;
        }
        public IAlibreAddOnCommand? Aboutmd(IADSession session)
        {
            MessageBox.Show(AddOnConstants.AboutMessage);
            return null;
        }
        public IAlibreAddOnCommand? RunCmd(IADSession session)
        {
            var ofd = new OpenFileDialog
            {
                Title = AddOnConstants.OpenFileDialogTitle,
                Filter = AddOnConstants.StlFileFilter,
                CheckFileExists = true,
                Multiselect = false
            };
            if (ofd.ShowDialog() != DialogResult.OK)
                return null;
            string stlPath = ofd.FileName;
            string stepPath = Path.ChangeExtension(stlPath, AddOnConstants.StepFileExtension);
            string addInDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)!;
            string exePath = Path.Combine(addInDir, AddOnConstants.ToolsDirectoryName, AddOnConstants.ConverterExeName);
            if (!File.Exists(exePath))
            {
                MessageBox.Show(string.Format(AddOnConstants.ConverterNotFoundFormat, exePath), AddOnConstants.ConverterMissingTitle, MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
            string args = string.Format(AddOnConstants.ArgumentFormat, stlPath, stepPath);
            try
            {
                var psi = new ProcessStartInfo(exePath, args)
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                using var proc = Process.Start(psi);
                if (proc == null)
                {
                    MessageBox.Show(AddOnConstants.ProcessLaunchFailedMessage, AddOnConstants.StlToStepErrorTitle, MessageBoxButton.OK, MessageBoxImage.Error);
                    return null;
                }
                string stdOutput = proc.StandardOutput.ReadToEnd();
                string stdError = proc.StandardError.ReadToEnd();
                proc.WaitForExit();

                if (proc.ExitCode == 0 && File.Exists(stepPath))
                {
                    FileInfo stepFileInfo = new FileInfo(stepPath);
                    string successMessage = AddOnConstants.ConversionSuccessMessage +
                                            string.Format(AddOnConstants.ImportAttemptMessageFormat, stepPath) +
                                            AddOnConstants.FileExistsMessage +
                                            string.Format(AddOnConstants.FileSizeMessageFormat, stepFileInfo.Length);
                    MessageBox.Show(successMessage, AddOnConstants.VerifyImportTitle, MessageBoxButton.OK, MessageBoxImage.Information);
                    try
                    {
                        IADRoot alibreRoot = AlibreAddOn.GetRoot();
                        if (alibreRoot != null)
                        {
                            alibreRoot.ImportSTEPFile(stepPath);
                        }
                        else
                        {
                            MessageBox.Show(AddOnConstants.AlibreRootNotFoundMessage, AddOnConstants.AlibreImportErrorTitle, MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(string.Format(AddOnConstants.StepImportFailedFormat, ex.Message, stepPath), AddOnConstants.AlibreImportErrorTitle, MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    string errorDetails = string.Format(AddOnConstants.ConversionFailedFormat, Path.GetFileName(stlPath));
                    errorDetails += string.Format(AddOnConstants.ExitCodeFormat, proc.ExitCode);

                    if (!string.IsNullOrWhiteSpace(stdOutput))
                        errorDetails += string.Format(AddOnConstants.StdOutputFormat, stdOutput);
                    if (!string.IsNullOrWhiteSpace(stdError))
                        errorDetails += string.Format(AddOnConstants.StdErrorFormat, stdError);
                    if (!File.Exists(stepPath))
                        errorDetails += AddOnConstants.OutputFileMissingMessage + stepPath;
                    else if (proc.ExitCode != 0)
                        errorDetails += AddOnConstants.ToolExitedWithError;
                    MessageBox.Show(errorDetails, AddOnConstants.StlToStepConversionErrorTitle, MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(AddOnConstants.UnexpectedProcessErrorFormat, ex.Message), AddOnConstants.ProcessErrorTitle, MessageBoxButton.OK, MessageBoxImage.Error);
            }
            return null;
        }
    }
    public class MenuManager
    {
        private readonly MenuItem _rootMenuItem;
        private readonly Dictionary<int, MenuItem> _menuItems;
        public MenuManager(IADSession topMostSession)
        {
            _rootMenuItem = new MenuItem(401, AddOnConstants.AddOnName, AddOnConstants.RootMenuToolTip);
            _menuItems = new Dictionary<int, MenuItem>();
            BuildMenus();
        }
        private void BuildMenus()
        {
            var stlStpGroup = new MenuItem(9088, AddOnConstants.SubMenuGroupText);
            var aboutStlStp = new MenuItem(9089, AddOnConstants.AboutMenuText, AddOnConstants.AboutMenuToolTip, AddOnConstants.IconPath);
            aboutStlStp.Command = aboutStlStp.Aboutmd;
            stlStpGroup.AddSubItem(aboutStlStp);
            var runStlStp = new MenuItem(9090, AddOnConstants.RunMenuText, AddOnConstants.RunMenuToolTip, AddOnConstants.IconPath);
            runStlStp.Command = runStlStp.RunCmd;
            stlStpGroup.AddSubItem(runStlStp);
            _rootMenuItem.AddSubItem(stlStpGroup);
            RegisterMenuItem(_rootMenuItem);
        }
        private void RegisterMenuItem(MenuItem menuItem)
        {
            _menuItems[menuItem.Id] = menuItem;
            foreach (var subItem in menuItem.SubItems)
                RegisterMenuItem(subItem);
        }
        public MenuItem? GetMenuItemById(int id)
        {
            _menuItems.TryGetValue(id, out MenuItem? menuItem);
            return menuItem;
        }
        public MenuItem GetRootMenuItem() => _rootMenuItem;
    }
}
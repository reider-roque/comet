#define NEWOUTPUT

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;
using TsudaKageyu;

namespace Comet
{
    public class AppContext : ApplicationContext
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool SetForegroundWindow(HandleRef hWnd);

        private const string MenuFileName = "menu.config";
        
        private static readonly IContainer _components = new Container();
        private static NotifyIcon _notifyIcon;
        private static FileInfo _menuFileInfo;
        private static Point _mousePosition;
        private static ContextMenuStrip _contextMenu;
        private static int _hotkeyId;
        
        public AppContext()
        {
            _menuFileInfo = new FileInfo(MenuFileName);
            if (!_menuFileInfo.Exists)
            {
                MessageBox.Show("Menu configuration file wasn't found.\nFile: " + _menuFileInfo.FullName, "Error");
                Environment.Exit(1);
            }

            _notifyIcon = new NotifyIcon(_components)
            {
                // Using the same icon that was used in Application properties.
                Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath),
                Text = "Comet",
                Visible = true
            };

            HotKeyManager.HotKeyPressed += HotKeyManager_HotKeyPressed;
            BuildIconContextMenu();
            BuildPopUpContextMenu();
        }

        private static void ActivateShortcut(XElement element)
        {
            foreach (var el in element.Elements())
            {
                if (el.Name == "keyboardShortcut")
                {
                    var keyCombo = el.Attribute("keyCombo");
                    if (keyCombo == null)
                    {
                        continue;
                    }

                    var keyModifiers = KeyModifiers.None;
                    if (keyCombo.Value.Contains("^")) { keyModifiers |= KeyModifiers.Control; }
                    if (keyCombo.Value.Contains("+")) { keyModifiers |= KeyModifiers.Shift; }
                    if (keyCombo.Value.Contains("!")) { keyModifiers |= KeyModifiers.Alt; }
                    if (keyCombo.Value.Contains("#")) { keyModifiers |= KeyModifiers.Windows; }

                    // Last char
                    var chr = keyCombo.Value[keyCombo.Value.Length - 1];
                    var key = (Keys) char.ToUpper(chr);

                    // Allow only chars in A-Za-z0-9 range
                    if (!((key > Keys.A && key < Keys.Z) || (key > Keys.D0 && key < Keys.D9)))
                    {
                        _notifyIcon.ShowBalloonTip(
                            10000,
                            "Error",
                            String.Format("Failed to create the \"{0}\" shortcut. The last key in the shortcut " +
                                          "sequence can only be one from the A-Za-z0-9 range.", keyCombo.Value),
                            ToolTipIcon.Error
                            );
                    }
                    HotKeyManager.UnRegisterHotKey(_hotkeyId);
                    _hotkeyId = HotKeyManager.RegisterHotKey(key, keyModifiers);

                    break;
                }
                // Besides keyBoardShortcut in the root menu tag all other elements are menuItem
            }
        }

        private static List<MenuItem> GetMenuItems(XElement element)
        {
            var menuItems = new List<MenuItem>();
            foreach (var el in element.Elements())
            {
                // Only process menuItem tags
                if (el.Name != "menuItem")
                {
                    continue;
                }
                // Besides keyBoardShortcut in the root menu tag all other elements are menuItem

                var name    = el.Attribute("name");
                var program = el.Attribute("program");
                var args    = el.Attribute("args");
                var icon    = el.Attribute("icon");
                var hidden  = el.Attribute("hidden");

                if (name == null) continue;

                var nameVal    = name.Value;
                var programVal = (program == null) ? string.Empty : program.Value;
                var argsVal    = (args    == null) ? string.Empty : args.Value;
                var iconVal    = (icon    == null) ? string.Empty : icon.Value;
                // Handle special keyword program in icon attribute
                if (iconVal == "program") { iconVal = programVal; }
                var hiddenVal  = false; //Default
                if (hidden != null)
                {
                    Boolean parsingResult;
                    if (Boolean.TryParse(hidden.Value, out parsingResult))
                    {
                        hiddenVal = parsingResult;
                    }
                }

                var menuItem = new MenuItem
                    {
                        Name = nameVal,
                        Program = programVal,
                        Arguments = argsVal,
                        Hidden = hiddenVal,
                        Icon = iconVal,
                        SubMenus = GetMenuItems(el)
                    };

                menuItems.Add(menuItem);
            }
            return menuItems;
        }

        private static void BuildIconContextMenu()
        {
            // Get MenuItems from menu configuration file
            var menuRoot = XDocument.Load(_menuFileInfo.FullName).Element("menu");
            var menuItems = GetMenuItems(menuRoot);
            
            // menuItem can never be null, no need to handle
            // If menuItem == 0 that's also fine, the menu will consist of Help and Exit items only

            var niContextMenu = new ContextMenuStrip(_components);

            // Capture mouse position for 'Reload Menu' functionality
            niContextMenu.MouseClick += (sender, e) => _mousePosition = Control.MousePosition;

            foreach (var menuItem in menuItems)
            {
                niContextMenu.Items.Add(CreateToolStripMenuItem(menuItem));
            }

            /* Reload [Help submenu] */
            var reloadMenu = new ToolStripMenuItem {Text = "Reload Menu"};
            reloadMenu.Click += (sender, e) =>
                {
                    BuildIconContextMenu();
                    BuildPopUpContextMenu();
                    _notifyIcon.ContextMenuStrip.Show(_mousePosition);
                };

            /* Comet Folder [Help submenu] */
            var folderMenu = new ToolStripMenuItem { Text = "Comet Folder" };
            folderMenu.Click += (sender, e) => Process.Start(Application.StartupPath);

            /* About [Help submenu] */
            var aboutMenu = new ToolStripMenuItem { Text = "About" };
            aboutMenu.Click += (sender, e) => MessageBox.Show(
                String.Format("Cosy Menu Tool (CoMeT)\n\nOleg Mitrofanov (www.wryway.com)\n" +
                              "All rights reserved © 2013 - {0}", DateTime.Now.Year), 
                "About Comet", 
                MessageBoxButtons.OK, 
                MessageBoxIcon.Information
            );

            /* Help menu */
            var helpMenu = new ToolStripMenuItem { Text = "Help" };
            helpMenu.DropDown.Items.AddRange(new ToolStripItem[] {reloadMenu, folderMenu, aboutMenu});

            /* Exit menu */
            var exitMenu = new ToolStripMenuItem { Text = "Exit" };
            exitMenu.Click += (sender, e) => {
                _notifyIcon.Visible = false;
                Environment.Exit(0);
            };
            
            niContextMenu.Items.AddRange(new ToolStripItem[] { helpMenu, exitMenu });

            _notifyIcon.ContextMenuStrip = niContextMenu;
        }

        private static void BuildPopUpContextMenu()
        {
            // Destroy the _contextMenu
            if (_contextMenu != null)
            {
                _contextMenu.Dispose();
            }

             _contextMenu = new ContextMenuStrip(_components);

            // Get MenuItems from menu configuration file
            var menuRoot = XDocument.Load(_menuFileInfo.FullName).Element("menu");
            ActivateShortcut(menuRoot);
            var menuItems = GetMenuItems(menuRoot);

            // Have to handle the case when menuItems.count is 0
            // Add an About entry in case there are 0 actual menu entries
            if (menuItems.Count == 0)
            {
                var aboutMenu = new ToolStripMenuItem { Text = "About" };
                aboutMenu.Click += (sender, e) => MessageBox.Show(
                    String.Format("Cosy Menu Tool (CoMeT)\n\nOleg Mitrofanov (www.wryway.com)\n" +
                                  "All rights reserved © 2013 - {0}", DateTime.Now.Year),
                    "About Comet",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                _contextMenu.Items.Add(aboutMenu);
            }
            else
            {
                foreach (var menuItem in menuItems)
                {
                    _contextMenu.Items.Add(CreateToolStripMenuItem(menuItem));
                }    
            }
        }

        static void HotKeyManager_HotKeyPressed(object sender, HotKeyEventArgs e)
        {
            // SetForegroundWindow is needed for not showing the ugly app icon it Windows taskbar
            SetForegroundWindow(new HandleRef(_notifyIcon.ContextMenuStrip, _notifyIcon.ContextMenuStrip.Handle));
            _contextMenu.Show(Cursor.Position);
        }

        private static ToolStripMenuItem CreateToolStripMenuItem(MenuItem menuItem)
        {
            if (menuItem == null) throw new ArgumentNullException();

            var toolStripMenuItem = new ToolStripMenuItem {Text = menuItem.Name};
            if (menuItem.Program != string.Empty)
            {
                toolStripMenuItem.Click += (sender, e) =>
                {
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = menuItem.Program,
                        Arguments = menuItem.Arguments,
                        WindowStyle = ProcessWindowStyle.Normal
                    };

                    // Setting non-defaults in this case; otherwise defaults are what we need
                    if (menuItem.Hidden)
                    {
                        startInfo.CreateNoWindow = true;
                        // Setting this property to false enables you to redirect input, output, and error streams.
                        startInfo.UseShellExecute = false;
                        startInfo.RedirectStandardOutput = true;
                    }
                    
                    /* A better place for updating environment variables would be by
                     * reacting on WM_SETTINGCHANGE message, but because this application
                     * does not have a top window (only top window receive that message)
                     * it is updated here. */
                    UpdateEnvironmentVariables();

                    var process = new Process {StartInfo = startInfo};
                    try
                    {
                        process.Start();
                    }
                    catch
                    {
                        var tipText = String.Format("Error executing \"{0}\".", menuItem.Name);
                        _notifyIcon.ShowBalloonTip(5000, "Error", tipText, ToolTipIcon.Error);
                        return;
                    }

                    if (menuItem.Hidden)
                    {
                        var outputLines = new List<string>();
                        process.OutputDataReceived += (proc, line) => { if (!String.IsNullOrEmpty(line.Data)) outputLines.Add(line.Data); };
                        process.BeginOutputReadLine();
                        process.WaitForExit();
                        String output = String.Empty;
                        if (outputLines.Any()) output = outputLines.Last();
                        if (output.Equals(String.Empty)) output = "Completed without output.";
                        _notifyIcon.ShowBalloonTip(5000, "Completed", output, ToolTipIcon.Info);
                    }
/*
                    else
                    {
                        _notifyIcon.ShowBalloonTip(5000, "Completed", "Application started.", ToolTipIcon.Info);
                    }
*/
                };
            }

            var icon = GetIcon(menuItem.Icon);
            if (icon != null)
            {
                toolStripMenuItem.Image = icon;
            }


            foreach (var item in menuItem.SubMenus)
            {
                toolStripMenuItem.DropDown.Items.Add(CreateToolStripMenuItem(item));
            }

            return toolStripMenuItem;
        }

        private static Image GetIcon(String iconPath)
        {
            if (String.IsNullOrEmpty(iconPath) || !File.Exists(iconPath))
            {
                return null;
            }

            var ext = Path.GetExtension(iconPath);
            if (ext == ".exe") // Also works with .dll
            {
                /* IconExtractor does a better job (produces better resolution) extracting
                 * icons from .exe/.dll files then Icon.ExtractAssociatedIcon()
                 * Example was taken from here: 
                 * http://www.codeproject.com/Articles/26824/Extract-icons-from-EXE-or-DLL-files
                 */
                var iconExtractor = new IconExtractor(iconPath);
                var icon = iconExtractor.GetIcon(0);
                var splitIcons = IconUtil.Split(icon);
                return IconUtil.ToBitmap(splitIcons[0]);
            }

            if (ext == ".ico")
            {
                var icon = Icon.ExtractAssociatedIcon(iconPath);
                return icon == null ? null : icon.ToBitmap();
            }

            if (ext == ".png" || ext == ".bmp" || ext == ".gif" ||
                ext == ".jpg" || ext == ".jpeg" || ext == ".tiff")
            {
                return Image.FromFile(iconPath);
            }

            return null;
        }

        private static void UpdateEnvironmentVariables()
        {
            /* A regular user environment variable overrides completely a system one with
             * the same name if both exist, but only for the specific user it is specified 
             * for. However, the user path variables is treated differently. It is appended 
             * to the system path variable when evaluating, rather than completely 
             * replacing it. Source:
             * http://stackoverflow.com/questions/5126512/how-environment-variables-are-evaluated
             */ 

            var sysEnvVars = Environment.GetEnvironmentVariables(EnvironmentVariableTarget.Machine);
            foreach (DictionaryEntry envVar in sysEnvVars)
            {
                Environment.SetEnvironmentVariable((String)envVar.Key, (String)envVar.Value);
            }

            var usrEnvVars = Environment.GetEnvironmentVariables(EnvironmentVariableTarget.User);
            foreach (DictionaryEntry envVar in usrEnvVars)
            {
                // The PATH variable is treated differently
                if ((String)envVar.Key == "PATH")
                {
                    String sysPath = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Machine);
                    String combinedPath = sysPath + ";" + (String) envVar.Value; // Combine system and user paths
                    Environment.SetEnvironmentVariable("PATH", combinedPath);
                    continue;
                }
                Environment.SetEnvironmentVariable((String)envVar.Key, (String)envVar.Value);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (_components != null))
            {
                HotKeyManager.UnRegisterHotKey(_hotkeyId);
                _components.Dispose();
            }

            base.Dispose(disposing);
        }
    }

    public class MenuItem
    {
        public String         Arguments { get; set; }
        public List<MenuItem> SubMenus  { get; set; }
        public String         Name      { get; set; }
        public String         Program   { get; set; }
        public String         Icon      { get; set;}
        public Boolean        Hidden    { get; set; }

        public MenuItem()
        {
            Arguments = string.Empty;
            SubMenus  = new List<MenuItem>();
            Name      = string.Empty;
            Program   = string.Empty;
            Icon      = string.Empty;
            Hidden   = false;
        }
    }

}

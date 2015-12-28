#define NEWOUTPUT

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;

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

            BuildIconContextMenu();
            BuildPopUpContextMenu();
        }

        private static List<MenuItem> GetMenuItems(XElement element)
        {
            var menuItems = new List<MenuItem>();
            foreach (var el in element.Elements())
            {
                /*if (el.Name == "keyboardShortcut")
                {

                    continue;
                }
                // Besides keyBoardShortcut in the root menu tag all other elements are menuItem
*/
                var name    = el.Attribute("name");
                var program = el.Attribute("program");
                var args    = el.Attribute("args");
                var hidden  = el.Attribute("hidden");

                if (name == null) continue;

                var nameVal    = name.Value;
                var programVal = (program == null) ? string.Empty : program.Value;
                var argsVal    = (args    == null) ? string.Empty : args.Value;
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

            HotKeyManager.RegisterHotKey(Keys.B, KeyModifiers.Alt | KeyModifiers.Control);
            HotKeyManager.HotKeyPressed += HotKeyManager_HotKeyPressed;
        }

        private static void BuildPopUpContextMenu()
        {
            // Unregister Hot key if it is registered and destroy the _contextMenu
            if (_contextMenu != null)
            {
                HotKeyManager.UnRegisterHotKey(_hotkeyId);
                HotKeyManager.HotKeyPressed -= HotKeyManager_HotKeyPressed;
                _contextMenu.Dispose();
            }

             _contextMenu = new ContextMenuStrip(_components);

            // Get MenuItems from menu configuration file
            var menuRoot = XDocument.Load(_menuFileInfo.FullName).Element("menu");
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

            _hotkeyId = HotKeyManager.RegisterHotKey(Keys.B, KeyModifiers.Alt | KeyModifiers.Control);
            HotKeyManager.HotKeyPressed += HotKeyManager_HotKeyPressed;
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
                    
                    var process = new Process {StartInfo = startInfo};

                    try
                    {
                        process.Start();
                    }
                    catch
                    {
                        var tipText = String.Format("Command \"{0}\" wasn't executed correctly.", menuItem.Name);
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
                    else
                    {
                        _notifyIcon.ShowBalloonTip(5000, "Completed", "Application started.", ToolTipIcon.Info);
                    }
                };
            }

            foreach (var item in menuItem.SubMenus)
            {
                toolStripMenuItem.DropDown.Items.Add(CreateToolStripMenuItem(item));
            }

            return toolStripMenuItem;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (_components != null))
            {
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
        public Boolean        Hidden   { get; set; }

        public MenuItem()
        {
            Arguments = string.Empty;
            SubMenus  = new List<MenuItem>();
            Name      = string.Empty;
            Program   = string.Empty;
            Hidden   = false;
        }
    }

}

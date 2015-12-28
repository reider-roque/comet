using System;
using System.Windows.Forms;

namespace Comet
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            var hotKeyMessageLoop = new HotKeyMessageLoop();
            Application.AddMessageFilter(hotKeyMessageLoop);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new AppContext());

            Application.RemoveMessageFilter(hotKeyMessageLoop);
        }
    }
}

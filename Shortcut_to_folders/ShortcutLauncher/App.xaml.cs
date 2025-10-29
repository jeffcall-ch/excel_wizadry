using System;
using System.Threading;
using System.Windows;

namespace ShortcutLauncher
{
    public partial class App : Application
    {
        private static Mutex? _mutex = null;
        private const string MutexName = "ShortcutLauncher_SingleInstance_Mutex";

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            
            // Try to create a named mutex for single instance
            bool createdNew;
            _mutex = new Mutex(true, MutexName, out createdNew);
            
            if (!createdNew)
            {
                // Another instance is already running
                MessageBox.Show("Shortcut Launcher is already running.\n\nCheck the taskbar for the minimized window.", 
                    "Already Running", MessageBoxButton.OK, MessageBoxImage.Information);
                Shutdown();
                return;
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            // Release the mutex when application exits
            _mutex?.ReleaseMutex();
            _mutex?.Dispose();
            base.OnExit(e);
        }
    }
}

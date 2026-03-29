using System;
using System.IO;

namespace FileExplorer.App;

public static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        App.StartupPath = args.Length > 0 ? args[0] : null;

        global::WinRT.ComWrappersSupport.InitializeComWrappers();
        global::Microsoft.UI.Xaml.Application.Start((p) =>
        {
            var context = new global::Microsoft.UI.Dispatching.DispatcherQueueSynchronizationContext(
                global::Microsoft.UI.Dispatching.DispatcherQueue.GetForCurrentThread());
            global::System.Threading.SynchronizationContext.SetSynchronizationContext(context);
            new App();
        });
    }
}

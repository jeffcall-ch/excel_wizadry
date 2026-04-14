using System;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FileExplorer.App;

public static class Program
{
    internal const string PipeName = "FileExplorer_SingleInstance";

    [STAThread]
    static void Main(string[] args)
    {
        var processDirectory = Path.GetDirectoryName(Environment.ProcessPath ?? string.Empty);
        if (!string.IsNullOrWhiteSpace(processDirectory) && Directory.Exists(processDirectory))
        {
            Environment.CurrentDirectory = processDirectory;
        }

        var startupPath = args.Length > 0 ? args[0] : null;

        // ── Single-instance: try to hand off to an already-running instance ──
        if (TrySendToExistingInstance(startupPath))
            return; // existing instance received the path — we're done

        // ── First instance: own the pipe server and start normally ────────────
        App.StartupPath = startupPath;
        _ = Task.Run(() => RunPipeServer());

        global::WinRT.ComWrappersSupport.InitializeComWrappers();
        global::Microsoft.UI.Xaml.Application.Start((p) =>
        {
            var context = new global::Microsoft.UI.Dispatching.DispatcherQueueSynchronizationContext(
                global::Microsoft.UI.Dispatching.DispatcherQueue.GetForCurrentThread());
            global::System.Threading.SynchronizationContext.SetSynchronizationContext(context);
            new App();
        });
    }

    /// <summary>
    /// Tries to connect to an existing instance and send the path.
    /// Returns true if a handoff was made (caller should exit).
    /// </summary>
    private static bool TrySendToExistingInstance(string? path)
    {
        try
        {
            using var client = new NamedPipeClientStream(".", PipeName, PipeDirection.Out);
            client.Connect(timeout: 400); // ms — fail fast if no server
            var payload = path ?? string.Empty;
            var bytes = Encoding.Unicode.GetBytes(payload);
            client.Write(bytes, 0, bytes.Length);
            return true;
        }
        catch
        {
            return false; // no server running → we are the first instance
        }
    }

    /// <summary>
    /// Runs on a background thread in the first (server) instance.
    /// Accepts connections from subsequent launches and routes the path
    /// to the main window via <see cref="App.Instance"/>.
    /// </summary>
    private static void RunPipeServer()
    {
        while (true)
        {
            try
            {
                using var server = new NamedPipeServerStream(
                    PipeName,
                    PipeDirection.In,
                    maxNumberOfServerInstances: 1,
                    transmissionMode: PipeTransmissionMode.Byte,
                    options: PipeOptions.Asynchronous);

                server.WaitForConnection();

                using var ms = new MemoryStream();
                var buf = new byte[4096];
                int read;
                while ((read = server.Read(buf, 0, buf.Length)) > 0)
                    ms.Write(buf, 0, read);

                var path = Encoding.Unicode.GetString(ms.ToArray());
                App.Instance?.DispatchOpenPath(path);
            }
            catch
            {
                // Pipe error (e.g. app shutting down) — stop the loop.
                break;
            }
        }
    }
}

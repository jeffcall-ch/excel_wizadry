namespace RqGui;

static class Program
{
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        // Set Application User Model ID for proper taskbar icon persistence
        SetAppUserModelId("RqGui.FileSearch.1.0");
        
        // To customize application configuration such as set high DPI settings or default font,
        // see https://aka.ms/applicationconfiguration.
        ApplicationConfiguration.Initialize();
        Application.Run(new Form1());
    }
    
    // Win32 API to set Application User Model ID
    [System.Runtime.InteropServices.DllImport("shell32.dll", SetLastError = true)]
    static extern void SetCurrentProcessExplicitAppUserModelID([System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string AppID);
    
    static void SetAppUserModelId(string appId)
    {
        try
        {
            SetCurrentProcessExplicitAppUserModelID(appId);
        }
        catch
        {
            // Ignore errors - this is just for taskbar icon persistence
        }
    }
}
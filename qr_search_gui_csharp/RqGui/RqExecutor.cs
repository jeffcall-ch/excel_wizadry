using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RqGui
{
    /// <summary>
    /// Search options for rq.exe execution.
    /// </summary>
    public class SearchOptions
    {
        public bool UseGlob { get; set; }
        public bool UseRegex { get; set; }
        public bool CaseSensitive { get; set; }
        public bool IncludeHidden { get; set; }
        public string Extensions { get; set; } = "";
        public string FileType { get; set; } = "";
        public string SizeFilter { get; set; } = "";
        public string DateAfter { get; set; } = "";
        public string DateBefore { get; set; } = "";
        public int MaxDepth { get; set; } = 0;
        public int Threads { get; set; } = 0;
        public int Timeout { get; set; } = 0;
    }
    
    /// <summary>
    /// Result of a search execution.
    /// </summary>
    public class SearchResult
    {
        public List<string> FilePaths { get; set; } = new List<string>();
        public string StandardError { get; set; } = "";
        public int ExitCode { get; set; }
        public bool WasCancelled { get; set; }
    }

    /// <summary>
    /// Executes rq.exe file search utility and captures results.
    /// Uses secure argument passing to prevent command injection.
    /// </summary>
    public class RqExecutor
    {
        private readonly string _rqPath;

        public RqExecutor(string rqPath)
        {
            _rqPath = rqPath;
        }

        /// <summary>
        /// Execute rq.exe search with proper argument escaping to prevent command injection.
        /// </summary>
        /// <param name="directory">The root directory to search in.</param>
        /// <param name="searchTerm">The search pattern or term to find.</param>
        /// <param name="options">Search options and filters.</param>
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        /// <param name="progress">Progress reporter for real-time updates.</param>
        /// <returns>A SearchResult containing found file paths and execution details.</returns>
        /// <exception cref="DirectoryNotFoundException">Thrown when the directory does not exist.</exception>
        public async Task<SearchResult> ExecuteSearchAsync(
            string directory, 
            string searchTerm, 
            SearchOptions options,
            CancellationToken cancellationToken = default,
            IProgress<string>? progress = null)
        {
            var result = new SearchResult();
            
            // Thread-safe collection for results
            var results = new ConcurrentBag<string>();
            var errorOutput = new StringBuilder();
            
            try
            {
                // Validate directory exists
                if (!System.IO.Directory.Exists(directory))
                {
                    throw new DirectoryNotFoundException($"Directory not found: {directory}");
                }
                
                // Build arguments list (SAFE - no command injection possible)
                var argumentsList = new List<string>
                {
                    directory,          // <where to look>
                    searchTerm          // <what to find>
                };
                
                // Add search options
                if (options.UseGlob)
                    argumentsList.Add("--glob");
                    
                if (options.UseRegex)
                    argumentsList.Add("--regex");
                    
                if (options.CaseSensitive)
                    argumentsList.Add("--case");
                    
                if (options.IncludeHidden)
                    argumentsList.Add("--include-hidden");
                
                // Always follow symlinks (default behavior)
                argumentsList.Add("--follow-symlinks");
                
                // Extension filter (comma-separated list)
                if (!string.IsNullOrWhiteSpace(options.Extensions))
                {
                    argumentsList.Add("--ext");
                    argumentsList.Add(options.Extensions.Trim());
                }
                
                // File type filter
                if (!string.IsNullOrWhiteSpace(options.FileType) && options.FileType.ToLower() != "any")
                {
                    argumentsList.Add("--type");
                    argumentsList.Add(options.FileType.ToLower());
                }
                
                // Size filter
                if (!string.IsNullOrWhiteSpace(options.SizeFilter))
                {
                    argumentsList.Add("--size");
                    argumentsList.Add(options.SizeFilter);
                }
                
                // Date filters
                if (!string.IsNullOrWhiteSpace(options.DateAfter))
                {
                    argumentsList.Add("--after");
                    argumentsList.Add(options.DateAfter);
                }
                
                if (!string.IsNullOrWhiteSpace(options.DateBefore))
                {
                    argumentsList.Add("--before");
                    argumentsList.Add(options.DateBefore);
                }
                
                // Max depth
                if (options.MaxDepth > 0)
                {
                    argumentsList.Add("--max-depth");
                    argumentsList.Add(options.MaxDepth.ToString());
                }
                
                // Threads
                if (options.Threads > 0)
                {
                    argumentsList.Add("--threads");
                    argumentsList.Add(options.Threads.ToString());
                }
                
                // Timeout
                if (options.Timeout > 0)
                {
                    argumentsList.Add("--timeout");
                    argumentsList.Add(options.Timeout.ToString());
                }
                
                Logger.LogInfo($"Executing rq with args: {string.Join(" ", argumentsList.Select(a => $"\"{a}\""))}");
                
                var processInfo = new ProcessStartInfo
                {
                    FileName = _rqPath,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,              // CRITICAL: Hides console window
                    WindowStyle = ProcessWindowStyle.Hidden,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };
                
                // Add arguments safely (prevents command injection)
                foreach (var arg in argumentsList)
                {
                    processInfo.ArgumentList.Add(arg);
                }

                using (var process = new Process { StartInfo = processInfo })
                {
                    // Capture stdout
                    process.OutputDataReceived += (sender, e) =>
                    {
                        if (!string.IsNullOrWhiteSpace(e.Data))
                        {
                            results.Add(e.Data);
                            progress?.Report(e.Data);
                        }
                    };
                    
                    // Capture stderr
                    process.ErrorDataReceived += (sender, e) =>
                    {
                        if (!string.IsNullOrWhiteSpace(e.Data))
                        {
                            errorOutput.AppendLine(e.Data);
                            Logger.LogWarning($"RQ stderr: {e.Data}");
                        }
                    };

                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();
                    
                    // Wait for completion or cancellation
                    await process.WaitForExitAsync(cancellationToken);
                    
                    result.ExitCode = process.ExitCode;
                    result.WasCancelled = cancellationToken.IsCancellationRequested;
                }
                
                result.FilePaths = results.ToList();
                result.StandardError = errorOutput.ToString();
                
                Logger.LogInfo($"Search completed. Found {result.FilePaths.Count} files. Exit code: {result.ExitCode}");
                
                if (!string.IsNullOrWhiteSpace(result.StandardError))
                {
                    Logger.LogWarning($"Search stderr output: {result.StandardError}");
                }
            }
            catch (OperationCanceledException)
            {
                result.WasCancelled = true;
                Logger.LogInfo("Search was cancelled by user");
            }
            catch (Exception ex)
            {
                Logger.LogException("Search execution failed", ex);
                throw;
            }
            
            return result;
        }

        /// <summary>
        /// Validate that the rq.exe path is valid.
        /// </summary>
        public static bool ValidateRqPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return false;
            if (!File.Exists(path))
                return false;
            if (!path.EndsWith("rq.exe", StringComparison.OrdinalIgnoreCase))
                return false;
            return true;
        }
    }
}

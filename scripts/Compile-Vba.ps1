# Compile the entire project, including legacy routines not reached by tests.
# Excel reports compile errors in modal dialogs even when DisplayAlerts is false.
# Inspect/dismiss only dialogs owned by this build's isolated Excel process.
function Assert-VbaCompilation($Excel, $Project) {
    if (-not ('ExceleratorCompilerDialog' -as [type])) {
        Add-Type @'
using System;
using System.Text;
using System.Threading;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
public static class ExceleratorCompilerDialog {
    delegate bool WindowCallback(IntPtr window, IntPtr parameter);
    [DllImport("user32.dll")] static extern bool EnumWindows(WindowCallback callback, IntPtr parameter);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr window, WindowCallback callback, IntPtr parameter);
    [DllImport("user32.dll")] static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
    [DllImport("user32.dll")] static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr window, uint message, IntPtr w, IntPtr l);
    public static readonly ConcurrentQueue<string> Messages = new ConcurrentQueue<string>();
    static Timer timer;
    public static void Start(uint process) {
        timer = new Timer(_ => EnumWindows((window, unused) => {
            uint owner;
            GetWindowThreadProcessId(window, out owner);
            if (owner != process) return true;
            var name = new StringBuilder(200);
            GetClassName(window, name, name.Capacity);
            if (name.ToString() != "#32770") return true;
            var title = new StringBuilder(200);
            GetWindowText(window, title, title.Capacity);
            string message = title.ToString();
            EnumChildWindows(window, (child, ignored) => {
                var text = new StringBuilder(2000);
                GetWindowText(child, text, text.Capacity);
                message += " | " + text;
                return true;
            }, IntPtr.Zero);
            Messages.Enqueue(message);
            PostMessage(window, 0x111, new IntPtr(1), IntPtr.Zero);
            return true;
        }, IntPtr.Zero), null, 100, 100);
    }
    public static void Stop() {
        using (var completed = new ManualResetEvent(false)) {
            if (timer.Dispose(completed)) completed.WaitOne();
        }
    }
}
'@
    }
    $Project.VBComponents.Item('Utils').Activate()
    $compile = $Excel.VBE.CommandBars.FindControl(1, 578)
    if ($null -eq $compile) { throw 'The VBA Compile command is unavailable.' }
    if (-not $compile.Enabled) { return }
    [uint32]$processId = 0
    [void][ExceleratorCompilerDialog]::GetWindowThreadProcessId([IntPtr]$Excel.Hwnd, [ref]$processId)
    $messages = [Collections.Generic.List[string]]::new()
    $message = ''
    while ([ExceleratorCompilerDialog]::Messages.TryDequeue([ref]$message)) {}
    [ExceleratorCompilerDialog]::Start($processId)
    try { $compile.Execute() } finally { [ExceleratorCompilerDialog]::Stop() }
    while ([ExceleratorCompilerDialog]::Messages.TryDequeue([ref]$message)) { $messages.Add($message) }
    if ($messages.Count -gt 0 -or $compile.Enabled) {
        $pane = $Excel.VBE.ActiveCodePane
        [int]$firstLine = 0; [int]$firstColumn = 0; [int]$lastLine = 0; [int]$lastColumn = 0
        $pane.GetSelection([ref]$firstLine, [ref]$firstColumn, [ref]$lastLine, [ref]$lastColumn)
        $location = $pane.CodeModule.Parent.Name + ':' + $firstLine + ' ' + $pane.CodeModule.Lines($firstLine, 1)
        throw ('VBA compilation failed at ' + $location + '. ' + ($messages -join ' '))
    }
    Write-Output 'Full VBA project compilation passed.'
}

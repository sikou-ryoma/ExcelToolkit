using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelToolkit;

public sealed class ExcelApplication : IDisposable
{
    private Excel.Application? _application;
    private bool _disposed;

    public ExcelApplication()
    {
        _application = new Excel.Application
        {
            Visible = false,
            DisplayAlerts = false
        };
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        if (_application is not null)
        {
            _application.Quit();
            Marshal.FinalReleaseComObject(_application);
            _application = null;
        }

        _disposed = true;
    }
}

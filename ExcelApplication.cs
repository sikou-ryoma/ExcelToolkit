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

    internal Excel.Application Instance
    {
        get
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            return _application
                ?? throw new InvalidOperationException(
                    "Excelが起動していません。");
        }
    }

    public ExcelWorkbook CreateWorkbook()
    {
        Excel.Workbook workbook = Instance.Workbooks.Add();

        return new ExcelWorkbook(this, workbook);
    }

    public ExcelWorkbook OpenWorkbook(string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        string fullPath = Path.GetFullPath(filePath);

        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException(
                "Excelファイルが見つかりません。",
                fullPath);
        }

        Excel.Workbook workbook =
            Instance.Workbooks.Open(fullPath);

        return new ExcelWorkbook(this, workbook);
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        try
        {
            _application?.Quit();
        }
        finally
        {
            if (_application is not null)
            {
                Marshal.FinalReleaseComObject(_application);
                _application = null;
            }

            _disposed = true;
        }
    }
    
    public void RunMacro(string macroName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(macroName);

        Instance.Run(macroName);
    }
}

using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelToolkit;

public sealed class ExcelWorkbook : IDisposable
{
    private readonly ExcelApplication _application;
    private Excel.Workbook? _workbook;
    private bool _disposed;

    internal ExcelWorkbook(
        ExcelApplication application,
        Excel.Workbook workbook)
    {
        _application = application;
        _workbook = workbook;
    }

    internal Excel.Workbook Instance
    {
        get
        {
            ObjectDisposedException.ThrowIf(_disposed, this);

            return _workbook
                ?? throw new InvalidOperationException(
                    "ブックが開かれていません。");
        }
    }

    public void Save()
    {
        Instance.Save();
    }

    public void SaveAs(string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        string fullPath = Path.GetFullPath(filePath);

        Instance.SaveAs(fullPath);
    }

    public void Close(bool saveChanges = false)
    {
        if (_disposed)
        {
            return;
        }

        Excel.Workbook? workbook = _workbook;

        if (workbook is null)
        {
            return;
        }

        try
        {
            workbook.Close(SaveChanges: saveChanges);
        }
        finally
        {
            _workbook = null;
            _disposed = true;
            Marshal.FinalReleaseComObject(workbook);
        }
    }

    public void Dispose()
    {
        Close();
    }

    public ExcelWorksheet GetWorksheet(string name)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

        Excel.Sheets? worksheets = null;
        Excel.Worksheet? worksheet = null;

        try
        {
            worksheets = Instance.Worksheets;
            worksheet = (Excel.Worksheet)worksheets[name];

            return new ExcelWorksheet(worksheet);
        }
        catch
        {
            if (worksheet is not null)
            {
                Marshal.FinalReleaseComObject(worksheet);
            }

            throw;
        }
        finally
        {
            if (worksheets is not null)
            {
                Marshal.FinalReleaseComObject(worksheets);
            }
        }
    }

    public void RunMacro(string macroName)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        _application.RunMacro(macroName);
    }
}

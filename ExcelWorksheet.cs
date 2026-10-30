using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelToolkit;

public sealed class ExcelWorksheet : IDisposable
{
    private Excel.Worksheet? _worksheet;
    private bool _disposed;

    internal ExcelWorksheet(Excel.Worksheet worksheet)
    {
        _worksheet = worksheet;
    }

    private Excel.Worksheet Instance
    {
        get
        {
            ObjectDisposedException.ThrowIf(_disposed, this);

            return _worksheet
                ?? throw new InvalidOperationException(
                    "ワークシートが取得されていません。");
        }
    }

    public void SetValue(string address, object? value)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(address);

        Excel.Range? range = null;

        try
        {
            range = Instance.Range[address];
            range.Value2 = value;
        }
        finally
        {
            if (range is not null)
            {
                Marshal.FinalReleaseComObject(range);
            }
        }
    }

    public object? GetValue(string address)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(address);

        Excel.Range? range = null;

        try
        {
            range = Instance.Range[address];
            return range.Value2;
        }
        finally
        {
            if (range is not null)
            {
                Marshal.FinalReleaseComObject(range);
            }
        }
    }

    public void SetValues(string address, object?[,] values)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(address);
        ArgumentNullException.ThrowIfNull(values);

        if (values.GetLength(0) == 0 || values.GetLength(1) == 0)
        {
            throw new ArgumentException(
                "配列の行数と列数は1以上である必要があります。",
                nameof(values));
        }

        Excel.Range? range = null;

        try
        {
            range = Instance.Range[address];
            range.Value2 = values;
        }
        finally
        {
            if (range is not null)
            {
                Marshal.FinalReleaseComObject(range);
            }
        }
    }

    public object?[,] GetValues(string address)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(address);

        Excel.Range? range = null;

        try
        {
            range = Instance.Range[address];

            object? result = range.Value2;

            // 1セルの場合は配列ではなく単一の値が返る
            if (result is not Array source || source.Rank != 2)
            {
                return new object?[,]
                {
                { result }
                };
            }

            int rowCount = source.GetLength(0);
            int columnCount = source.GetLength(1);

            object?[,] values = new object?[rowCount, columnCount];

            int rowLower = source.GetLowerBound(0);
            int columnLower = source.GetLowerBound(1);

            for (int row = 0; row < rowCount; row++)
            {
                for (int column = 0; column < columnCount; column++)
                {
                    values[row, column] = source.GetValue(
                        row + rowLower,
                        column + columnLower);
                }
            }

            return values;
        }
        finally
        {
            if (range is not null)
            {
                Marshal.FinalReleaseComObject(range);
            }
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        if (_worksheet is not null)
        {
            Marshal.FinalReleaseComObject(_worksheet);
            _worksheet = null;
        }

        _disposed = true;
    }
}
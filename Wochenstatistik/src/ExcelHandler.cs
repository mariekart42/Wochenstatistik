namespace Wochenstatistik;
using Aspose.Cells;

public sealed class ExcelHandler
{
    private static ExcelHandler _instance;
    private Workbook _workbook;
    private readonly Worksheet _worksheet;
    private ExcelHandler(string path)
    {
        _workbook = new Workbook(path);
        _worksheet = _workbook.Worksheets[0];
    }

    public static Worksheet GetWorksheet(string path)
    {
        if (_instance == null)
            _instance = new ExcelHandler(path);
        return _instance._worksheet;
    }
}
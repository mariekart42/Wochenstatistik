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

    public static Worksheet GetWorksheet(string excel_file_path)
    {
        if (string.IsNullOrEmpty(excel_file_path) || File.Exists(excel_file_path) == false)
            throw new Exception("Please provide the Daten Wochenstatistik.xlsx file!");
        if (_instance == null)
            _instance = new ExcelHandler(excel_file_path);
        return _instance._worksheet;
    }
}
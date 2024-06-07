namespace Wochenstatistik;
using Aspose.Cells;

public sealed class ExcelHandler
{
    private static ExcelHandler _instance;
    private Workbook _workbook;
    private readonly Worksheet _worksheet;

    private static string EXCEL_FILE_PATH = "document/Daten Wochenstatistik.xlsx";
    private ExcelHandler(string path)
    {
        _workbook = new Workbook(path);
        _worksheet = _workbook.Worksheets[0];
    }

    public static Worksheet GetWorksheet()
    {
        if (string.IsNullOrEmpty(EXCEL_FILE_PATH) || File.Exists(EXCEL_FILE_PATH) == false)
            throw new Exception("Please provide the Daten Wochenstatistik.xlsx file!");
        if (_instance == null)
            _instance = new ExcelHandler(EXCEL_FILE_PATH);
        return _instance._worksheet;
    }
}
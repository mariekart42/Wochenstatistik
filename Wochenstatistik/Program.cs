using Aspose.Cells;
using Wochenstatistik;

try
{
    Worksheet worksheet = ExcelHandler.GetWorksheet();
    Dictionary<string, string> userDic = DataManager.GetUserDic(worksheet);

    foreach (var user in userDic)
    {
        Console.WriteLine($"SEND EMAIL TO: {user.Key}, {user.Value}");
        DataManager.InitData(worksheet, user);
        DataManager.sendMail();
    }
}
catch (Exception e)
{
    Console.WriteLine($"ERROR: {e}");
}

using Aspose.Cells;
using Wochenstatistik;

try
{
    var config = ConfigHandler.GetConfigFile();
    Worksheet worksheet = ExcelHandler.GetWorksheet(config["EXCEL_FILE_PATH"]);
    DataManager.InitConfigVariables(config);
    Dictionary<string, string> userDic = DataManager.GetUserDic(worksheet, config);
    Console.ForegroundColor = ConsoleColor.Green;
    foreach (var user in userDic)
    {
        DataManager.InitData(worksheet, user);
        DataManager.sendMail();
        Console.WriteLine($"SEND EMAIL TO: {user.Key}, {user.Value}");
    }
}
catch (Exception e)
{
    ConsoleColor originalColor = ConsoleColor.Black;
    Console.ForegroundColor = originalColor;
    Console.WriteLine("\nError in program:");
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("   "+e.Message);
    Console.ForegroundColor = originalColor;
    Console.WriteLine("\nStack Trace:");
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine(e.StackTrace+"\n");
}

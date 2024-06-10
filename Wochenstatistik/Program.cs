using Aspose.Cells;
using Wochenstatistik;

try
{
    Worksheet worksheet = ExcelHandler.GetWorksheet();
    Dictionary<string, string> userDic = DataManager.GetUserDic(worksheet);
    Console.ForegroundColor = ConsoleColor.Green;
    foreach (var user in userDic)
    {
        Console.WriteLine($"SEND EMAIL TO: {user.Key}, {user.Value}");
        DataManager.InitData(worksheet, user);
        DataManager.sendMail();
    }
}
catch (Exception e)
{
    ConsoleColor originalColor = Console.ForegroundColor;
    Console.WriteLine("\nError in program:");
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("   "+e.Message);
    Console.ForegroundColor = originalColor;
    Console.WriteLine("\nStack Trace:");
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine(e.StackTrace+"\n");
}
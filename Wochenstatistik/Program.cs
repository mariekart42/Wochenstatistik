using Aspose.Cells;
using Wochenstatistik;

try
{
    string excel_file_path = Environment.GetEnvironmentVariable("EXCEL_FILE_PATH");
    string user_file_path = Environment.GetEnvironmentVariable("USER_FILE_PATH");

    if (string.IsNullOrEmpty(excel_file_path) || File.Exists(excel_file_path) == false)
        throw new Exception("Please provide the Daten_Wochenstatistik.xlsx file!");
    if (string.IsNullOrEmpty(user_file_path) || File.Exists(user_file_path) == false)
        throw new Exception("Please provide the User_Wochenstatistik.txt file!");

    Worksheet worksheet = ExcelHandler.GetWorksheet(excel_file_path);

    Dictionary<string, string> userDic = DataManager.GetUserDic(user_file_path, worksheet);

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

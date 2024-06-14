namespace Wochenstatistik;

public static class ConfigHandler
{
    public static Dictionary<string, string> GetConfigFile()
    {
        Dictionary<string, string> config = new Dictionary<string, string>();

        string binLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
        string directoryPath = Path.GetDirectoryName(binLocation);
        string configPath = Path.Combine(directoryPath, "config.txt");
        if (!File.Exists(configPath))
        {
            Console.WriteLine($"The config.txt does not exist. Creating it...");
            using (FileStream fs = new FileStream(configPath, FileMode.Create))
            {
                using (StreamWriter writer = new StreamWriter(fs))
                {
                    writer.WriteLine("# lines that start with '#' will be ignored.");
                    writer.WriteLine("# program will not start, if not all variables are initialized.");
                    writer.WriteLine("# Don't use ANY quotation marks [\", \'].\n\n");
                    writer.WriteLine("EXCEL_FILE_PATH=document/Daten Wochenstatistik.xlsx");
                    writer.WriteLine("USER_FILE_PATH=document/Nutzer Liste.txt");
                    writer.WriteLine("EMAIL_HOST=umsatzstatistik@eisenfuhr.com");
                    writer.WriteLine("EMAIL_PASSWORD=");
                    writer.WriteLine("EMAIL_SERVER_HOST=rpx.eisenfuhr.com");
                    writer.WriteLine("EMAIL_SERVER_PORT=587");
                    writer.WriteLine("EMAIL_SERVER_SSL=false");
                }
            }
            throw new Exception(
                $"File 'config.txt' got created inside the root folder. Please override the default values and run the program again.\n   Path to your \'config.txt\': {configPath}");
        }
        Console.WriteLine($"Path to the \'config.txt\' file: {configPath}");

        config = GetConfigDic(configPath);
        config["PATH_TO_CONFIG"] = configPath;
        return config;
    }

    private static Dictionary<string, string> GetConfigDic(string configPath)
    {
        Dictionary<string, string> config = new Dictionary<string, string>();
        using (StreamReader reader = new StreamReader(configPath))
        {
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                line = line.Trim();
                if (!string.IsNullOrWhiteSpace(line) && !line.StartsWith('#'))
                {
                    int equalIndex = line.IndexOf('=');
                    if (equalIndex < 0)
                        continue;

                    string variable = line.Substring(0, equalIndex);
                    string value = line.Substring(equalIndex + 1);
                    if (string.IsNullOrWhiteSpace(value))
                        throw new Exception($"Invalid config.txt file. Value on this line not initialized: {line}.");;
                    switch (variable)
                    {
                        case "EXCEL_FILE_PATH":
                            config["EXCEL_FILE_PATH"] = value;
                            break;
                        case "USER_FILE_PATH":
                            config["USER_FILE_PATH"] = value;
                            break;
                        case "EMAIL_HOST":
                            config["EMAIL_HOST"] = value;
                            break;
                        case "EMAIL_PASSWORD":
                            config["EMAIL_PASSWORD"] = value;
                            break;
                        case "EMAIL_SERVER_HOST":
                            config["EMAIL_SERVER_HOST"] = value;
                            break;
                        case "EMAIL_SERVER_PORT":
                            config["EMAIL_SERVER_PORT"] = value;
                            break;
                        case "EMAIL_SERVER_SSL":
                            config["EMAIL_SERVER_SSL"] = value;
                            break;
                        default:
                            throw new Exception($"Invalid config.txt file. Error on this line: {line}.");
                    }
                }
            }
        }
        return config;
    }
}
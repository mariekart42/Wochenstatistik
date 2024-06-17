using System.Net.Mail;
using System.Reflection;
using Aspose.Cells;
using MimeKit;
using SmtpClient = MailKit.Net.Smtp.SmtpClient;

namespace Wochenstatistik;

public static class DataManager
{
    private static string _emailTo;
    private static string _firm;
    private static string? _currentMonth;
    private static string? _monthSpan;
    private static string? _currentMonthBeaTEUR;
    private static string? _currentMonthGFTEUR;
    private static string? _currentMonthTaxTEUR;
    private static string? _currentMonthGesamtTEUR;
    private static string? _monthSpanBeaTEUR;
    private static string? _monthSpanBeaVorjahrMonat;
    private static string? _monthSpanGFTEUR;
    private static string? _monthSpanGFVorjahrMonat;
    private static string? _monthSpanTaxTEUR;
    private static string? _monthSpanTaxVorjahrMonat;
    private static string? _monthSpanGesamtTEUR;
    private static string? _monthSpanGesamtVorjahrMonat;
    private static string? _offeneLeistungskostenEigen;
    private static string? _offeneLeistungskostenFremd;
    private static string? _offeneLeistungskostenGesamt;

    private static string USER_FILE_PATH;
    private static string EMAIL_HOST;
    private static string EMAIL_PASSWORD;
    private static string EMAIL_SERVER_HOST;
    private static int EMAIL_SERVER_PORT;
    private static bool EMAIL_SERVER_SSL;

    private static string PATH_TO_CONFIG;
    private static string PATH_TO_USER_FILE;

    private static string GetFormattedValue(Cell cell, bool isPercent)
    {
        string value = cell.Value.ToString();

        if (value == "XXX")
            return "/";
        if (isPercent)
            return CutAfterOneNumber(ToPercent(value)) + '%';
        return CutAfterOneNumber(value);
    }

    public static void InitData(Worksheet worksheet, KeyValuePair<string, string> user)
    {
        _emailTo = user.Value;
        _firm = user.Key;
        int rowIndex = GetRowIndex(worksheet, user.Key);
        Dictionary<char, Cell> rowData = GetDataFromRowAsArray(worksheet, rowIndex);

        _currentMonth = worksheet.Cells[0, 2].Value.ToString();
        _monthSpan = worksheet.Cells[0, 11].Value.ToString();
        _currentMonthBeaTEUR = GetFormattedValue(rowData['C'], false);
        _currentMonthGFTEUR = GetFormattedValue(rowData['E'], false);
        _currentMonthTaxTEUR = GetFormattedValue(rowData['G'], false);
        _currentMonthGesamtTEUR = GetFormattedValue(rowData['I'], false);
        _monthSpanBeaTEUR = GetFormattedValue(rowData['L'], false);
        _monthSpanBeaVorjahrMonat = GetFormattedValue(rowData['M'], true);
        _monthSpanGFTEUR = GetFormattedValue(rowData['N'], false);
        _monthSpanGFVorjahrMonat = GetFormattedValue(rowData['O'], true);
        _monthSpanTaxTEUR = GetFormattedValue(rowData['P'], false);
        _monthSpanTaxVorjahrMonat = GetFormattedValue(rowData['Q'], true);
        _monthSpanGesamtTEUR = GetFormattedValue(rowData['R'], false);
        _monthSpanGesamtVorjahrMonat = GetFormattedValue(rowData['S'], true);
        _offeneLeistungskostenEigen = GetFormattedValue(rowData['U'], false);
        _offeneLeistungskostenFremd = GetFormattedValue(rowData['V'], false);
        _offeneLeistungskostenGesamt = GetFormattedValue(rowData['W'], false);
    }

    private static char ToAsciiLetter(int num)
    {
        return (char)(num + 64);
    }

    private static int GetRowIndex(Worksheet worksheet, string? user)
    {
        if (string.IsNullOrEmpty(user))
            throw new Exception("The username can't be empty! Exit.");

        for (int i = 0; i <= worksheet.Cells.MaxDataRow; i++)
            if ((string)worksheet.Cells[i, 0].Value == user)
                return i;
        throw new Exception($"The user [{user}] does not exist. Exit");
    }

    private static string ToPercent(string number)
    {
        if (string.IsNullOrEmpty(number) || number == "XXX")
            return number;

        bool isNegative = number.StartsWith("-");
        if (isNegative)
            number = number.Substring(1);
        int commaIndex = number.IndexOf('.');
        if (commaIndex == -1)
            number += "00";
        else
        {
            string integerPart = number.Substring(0, commaIndex);
            string fractionalPart = number.Substring(commaIndex + 1);

            while (fractionalPart.Length < 2)
                fractionalPart += "0";
            number = integerPart + fractionalPart.Substring(0, 2) + "." + fractionalPart.Substring(2);
        }

        number = number.TrimStart('0');
        if (string.IsNullOrEmpty(number) || number.StartsWith('.'))
            number = "0" + number;
        if (isNegative)
            number = "-" + number;
        return number;
    }

    private static Dictionary<char, Cell> GetDataFromRowAsArray(Worksheet worksheet, int rowIndex)
    {
        Row row = worksheet.Cells.Rows[rowIndex];
        Dictionary<char, Cell> dictionary = new Dictionary<char, Cell>();

        for (int i = 2; i <= worksheet.Cells.MaxDataColumn; i++)
        {
            object cellValue = worksheet.Cells[rowIndex, i].Value;
            if (cellValue == null|| string.IsNullOrWhiteSpace(cellValue.ToString()))
                continue;
            char letter = ToAsciiLetter(i + 1);
            dictionary[ToAsciiLetter(i+1)] = row[i];
            if (letter == 'D' || letter == 'F' || letter == 'H' || letter == 'J' || letter == 'M' || letter == 'O' ||
                letter == 'Q' || letter == 'S')
                dictionary[ToAsciiLetter(i+1)].Value = row[i].Value;
        }
        return dictionary;
    }

    private static string CutAfterOneNumber(string input)
    {
        if (decimal.TryParse(input, out decimal number))
        {
            decimal roundedNumber = Math.Round(number, 1, MidpointRounding.AwayFromZero);
            string result = roundedNumber.ToString("0.0");
            if (result == "-100.0") return "-100";
            if (result == "100.0") return "100";
            return result;
        }
        return input;
    }

    private static string GetStyleDiv(string value, string highlight_colour)
    {
        if (value.StartsWith("-"))
            return "<td class=\"highlight-" + highlight_colour + "\" style=\"color: red;\">" + value + "</td>";
        else
            return "<td class=\"highlight-" + highlight_colour + "\">" + value + "</td>";
    }


    static string ReadEmbeddedResource(string resourceName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        using (Stream stream = assembly.GetManifestResourceStream(resourceName))
        {
            if (stream == null)
            {
                throw new Exception($"Resource not found: {resourceName}");
            }
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }
    }

    private static string GetHtmlContent()
    {
        string result = ReadEmbeddedResource("Wochenstatistik.html.styling.html");
        result += "<title>Hallo " + _firm + ",</title></head><body><div class=\"table-container\"><table><thead>\n<tr><th colspan=\"4\" class=\"main-header\">";
        result += _currentMonth + "</th><th colspan=\"8\" class=\"main-header\">";
        result += _monthSpan;
        result += ReadEmbeddedResource("Wochenstatistik.html.headingTitles.html");
        result += GetStyleDiv(_currentMonthBeaTEUR, "yellow");
        result += GetStyleDiv(_currentMonthGFTEUR, "green");
        result += GetStyleDiv(_currentMonthTaxTEUR, "blue");
        result += GetStyleDiv(_currentMonthGesamtTEUR, "purple");
        result += GetStyleDiv(_monthSpanBeaTEUR, "yellow");
        result += GetStyleDiv(_monthSpanBeaVorjahrMonat, "yellow");
        result += GetStyleDiv(_monthSpanGFTEUR, "green");
        result += GetStyleDiv(_monthSpanGFVorjahrMonat, "green");
        result += GetStyleDiv(_monthSpanTaxTEUR, "blue");
        result += GetStyleDiv(_monthSpanTaxVorjahrMonat, "blue");
        result += GetStyleDiv(_monthSpanGesamtTEUR, "purple");
        result += GetStyleDiv(_monthSpanGesamtVorjahrMonat, "purple");
        result += GetStyleDiv(_offeneLeistungskostenEigen, "");
        result += GetStyleDiv(_offeneLeistungskostenFremd, "");
        result += GetStyleDiv(_offeneLeistungskostenGesamt, "");
        result += "</tr></tbody></table></div></body></html>";
        return result;
    }

    public static void sendMail()
    {
        var message = new MimeMessage ();

        message.From.Add(new MailboxAddress("Wochenstatistik", EMAIL_HOST));
        message.To.Add(new MailboxAddress(_firm, _emailTo));
        message.Subject = "Wochenstatistik für " + _firm;

        message.Body = new TextPart("html")
        { Text = "Hallo " + _firm + ",<br><br>anbei Ihre Wochenstatistik:<br><br>"+ GetHtmlContent() + "<br>Beste Grüße<br>Ihre Buchhaltung" };

        try
        {
            using var client = new SmtpClient ();
            client.Connect (EMAIL_SERVER_HOST, EMAIL_SERVER_PORT, EMAIL_SERVER_SSL);
            client.Authenticate(EMAIL_HOST, EMAIL_PASSWORD);
            client.Send (message);
            client.Disconnect (true);
        }
        catch
        {
            throw new Exception($"SmtpClient failed to send Mail to \'{_emailTo}\'. Please check and correct your \'config.txt\' file. One or more of the following settings caused the error:\n\t- EMAIL_HOST={EMAIL_HOST}\n\t- EMAIL_PASSWORD={EMAIL_PASSWORD}\n\t- EMAIL_SERVER_HOST={EMAIL_SERVER_HOST}\n\t- EMAIL_SERVER_PORT={EMAIL_SERVER_PORT}\n\t- EMAIL_SERVER_SSL={EMAIL_SERVER_SSL}\n   Path to your \'config.txt\': {PATH_TO_CONFIG}");
        }
    }

    private static bool IsValidEmail(string email)
    {
        if (string.IsNullOrWhiteSpace(email))
            return false;
        try
        {
            var addr = new MailAddress(email);
            return addr.Address == email;
        }
        catch (FormatException)
        {
            return false;
        }
    }

    public static void InitConfigVariables(Dictionary<string, string> config)
    {
        string user_file_path = config["USER_FILE_PATH"];
        string email_host = config["EMAIL_HOST"];
        string email_password = config["EMAIL_PASSWORD"];
        string email_server_host = config["EMAIL_SERVER_HOST"];
        string email_server_port = config["EMAIL_SERVER_PORT"];
        string email_server_ssl = config["EMAIL_SERVER_SSL"];
        string path_to_config = config["PATH_TO_CONFIG"];
        string path_to_user_file = config["PATH_TO_USER_FILE"];

        if (string.IsNullOrEmpty(user_file_path) || File.Exists(user_file_path) == false)
            throw new Exception($"\'config.txt\' file is invalid. Please define the path to the \'Nutzer Liste.txt\' file, it should not be empty.\n   Path to your \'config.txt\': {PATH_TO_CONFIG}");
        USER_FILE_PATH = user_file_path;
        if (string.IsNullOrEmpty(email_host) || string.IsNullOrEmpty(email_password))
            throw new Exception($"\'config.txt\' file is invalid. Please assign a value to EMAIL_HOST and EMAIL_PASSWORD. They should not be empty.\n   Path to your \'config.txt\': {PATH_TO_CONFIG}");
        EMAIL_HOST = email_host;
        EMAIL_PASSWORD = email_password;
        if (string.IsNullOrEmpty(email_server_host) || !email_server_port.All(char.IsDigit) || email_server_port == "-1")
            throw new Exception($"\'config.txt\' file is invalid. Please assign a value to EMAIL_SERVER_HOST and EMAIL_SERVER_PORT. They should not be empty.\n   Path to your \'config.txt\': {PATH_TO_CONFIG}");
        EMAIL_SERVER_HOST = email_server_host;
        EMAIL_SERVER_PORT = int.Parse(email_server_port);
        if (email_server_ssl != "false" && email_server_ssl != "true")
            throw new Exception(
                $"\'config.txt\' file is invalid. EMAIL_SERVER_SSL is type bool. Please set it to true (to use ssl) or false (to not use ssl).\n   Path to your \'config.txt\': {PATH_TO_CONFIG}");
        EMAIL_SERVER_SSL = bool.Parse(email_server_ssl);
        PATH_TO_CONFIG = path_to_config;
        PATH_TO_USER_FILE = path_to_user_file;
    }

    public static Dictionary<string, string> GetUserDic(Worksheet worksheet, Dictionary<string, string> config)
    {
        Dictionary<string, string> userDic = new Dictionary<string, string>();
        using StreamReader file = new StreamReader(USER_FILE_PATH);
        string line;
        while ((line = file.ReadLine()) != null)
        {
            if (line.Length == 0 || line.All(c => c == ' ' || c == '\t') || line.StartsWith('#'))
                continue;

            string[] fields = line.Split('|');
            if (fields.Length == 2)
            {
                string email = fields[0].Trim();
                if (!IsValidEmail(email))
                    throw new Exception($"Wrong syntax in \'Nutzer Liste.txt\' file. E-Mail address has wrong syntax. Found E-Mail: {email}.\n   Path to your \'Nutzer Liste.txt\': {PATH_TO_USER_FILE}");
                string firm = fields[1].Trim().ToUpper();
                bool found = false;
                for (int i = 0; i <= worksheet.Cells.MaxDataRow; i++)
                {
                    var cellValue = worksheet.Cells[i, 0].Value?.ToString().Trim().ToUpper();
                    if (cellValue == firm)
                    {
                        found = true;
                        break;
                    }
                }
                if (found)
                    userDic[firm] = email;
                else
                    throw new Exception($"Wrong syntax in \'Nutzer Liste.txt\' file. Firm is invalid or does not exist in the \'Daten Wochenstatistik.xlsx\' file. Found: {firm}.\n   Path to your \'Nutzer Liste.txt\': {PATH_TO_USER_FILE}");
            }
            else
                throw new Exception($"Wrong syntax in \'Nutzer Liste.txt\' file. Expected: <email>|<firm>. Found: {line}.\n   Path to your \'Nutzer Liste.txt\': {PATH_TO_USER_FILE}");
        }
        return userDic;
    }
}
using System.Net.Mail;
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

    private static string USER_FILE_PATH = "document/Nutzer Liste.txt";
    private static string EMAIL_HOST = null; // PROVIDE EMAIL HOST HERE
    private static string EMAIL_PASSWORD = null; // PROVIDE EMAIL PASSWORD HERE;
    private static string EMAIL_SERVER_HOST = null; // PROVIDE EMAIL SERVER HOST HERE eg. smtp-mail.outlook.com;
    private static int EMAIL_SERVER_PORT = -1; // PROVIDE EMAIL SERVER PORT HERE eg. 25 or 587;
    private static bool EMAIL_SERVER_SSL = false; // PROVIDE EMAIL SERVER SSL HERE eg. false or true;

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

    private static string GetHtmlContent()
    {
        string result = File.ReadAllText("html/styling.html");
        result += "<title>Hallo " + _firm + ",</title></head><body><div class=\"table-container\"><table><thead>\n<tr><th colspan=\"4\" class=\"main-header\">";
        result += _currentMonth + "</th><th colspan=\"8\" class=\"main-header\">";
        result += _monthSpan;
        result += File.ReadAllText("html/headingTitles.html");
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
        { Text = "Hallo " + _firm + ",<br><br>Hier ist Ihre Wochenstatistik:<br><br>"+ GetHtmlContent() + "<br>Beste Grüße, Ihre Buchhaltung" };

        using var client = new SmtpClient ();
        client.Connect (EMAIL_SERVER_HOST, EMAIL_SERVER_PORT, EMAIL_SERVER_SSL);
        client.Authenticate(EMAIL_HOST, EMAIL_PASSWORD);
        client.Send (message);
        client.Disconnect (true);
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

    public static Dictionary<string, string> GetUserDic(Worksheet worksheet)
    {
        if (string.IsNullOrEmpty(USER_FILE_PATH) || File.Exists(USER_FILE_PATH) == false)
            throw new Exception("Please provide the Nutzer Liste.txt file!");
        if (string.IsNullOrEmpty(EMAIL_HOST) || string.IsNullOrEmpty(EMAIL_PASSWORD))
            throw new Exception("Please provide your E-Mail address and credentials in DataManager.cs. They should not be empty.");
        if (string.IsNullOrEmpty(EMAIL_SERVER_HOST) || EMAIL_SERVER_PORT == -1)
            throw new Exception("Please provide the E-Mail server host and port in DataManager.cs. They should not be empty.");

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
                    throw new Exception($"Wrong syntax in User_Wochenstatistik file. E-Mail address has wrong syntax. Found E-Mail: {email}. Exit.");
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
                    throw new Exception($"Wrong syntax in User_Wochenstatistik file. Firm is invalid. Found: {firm}. Exit.");
            }
            else
                throw new Exception($"Wrong syntax in User_Wochenstatistik file. Expected: <email>|<firm>. Found: {line}. Exit.");
        }
        return userDic;
    }
}
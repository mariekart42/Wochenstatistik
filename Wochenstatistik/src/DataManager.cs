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
    private static string? _currentMonthBeaVorjahrMonat;
    private static string? _currentMonthGFTEUR;
    private static string? _currentMonthGFVorjahrMonat;
    private static string? _currentMonthTaxTEUR;
    private static string? _currentMonthTaxVorjahrMonat;
    private static string? _currentMonthGesamtTEUR;
    private static string? _currentMonthGesamtVorjahrMonat;
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

    public static void InitData(Worksheet worksheet, KeyValuePair<string, string> user)
    {
        _emailTo = user.Value;
        _firm = user.Key;
        int rowIndex = GetRowIndex(worksheet, user.Key);
        Dictionary<char, Cell> rowData = GetDataFromRowAsArray(worksheet, rowIndex);

        // 0 / C -> Mai
        // 0 / L -> Januar - Mai
        _currentMonth = "Mai";
        _monthSpan = "Januar - Mai";
        _currentMonthBeaTEUR = CutAfterOneNumber(rowData['C'].Value.ToString());
        _currentMonthBeaVorjahrMonat = CutAfterOneNumber(ToPercent(rowData['D'].Value.ToString())) + '%';
        _currentMonthGFTEUR = CutAfterOneNumber(rowData['E'].Value.ToString());
        _currentMonthGFVorjahrMonat = CutAfterOneNumber(ToPercent(rowData['F'].Value.ToString())) + '%';
        _currentMonthTaxTEUR = CutAfterOneNumber(rowData['G'].Value.ToString());
        _currentMonthTaxVorjahrMonat = CutAfterOneNumber(ToPercent(rowData['H'].Value.ToString())) + '%';
        _currentMonthGesamtTEUR = CutAfterOneNumber(rowData['I'].Value.ToString());
        _currentMonthGesamtVorjahrMonat = CutAfterOneNumber(ToPercent(rowData['J'].Value.ToString())) + '%';
        _monthSpanBeaTEUR = CutAfterOneNumber(rowData['L'].Value.ToString());
        _monthSpanBeaVorjahrMonat = CutAfterOneNumber(ToPercent(rowData['M'].Value.ToString())) + '%';
        _monthSpanGFTEUR = CutAfterOneNumber(rowData['N'].Value.ToString());
        _monthSpanGFVorjahrMonat = CutAfterOneNumber(ToPercent(rowData['O'].Value.ToString())) + '%';
        _monthSpanTaxTEUR = CutAfterOneNumber(rowData['P'].Value.ToString());
        _monthSpanTaxVorjahrMonat = CutAfterOneNumber(ToPercent(rowData['Q'].Value.ToString())) + '%';
        _monthSpanGesamtTEUR = CutAfterOneNumber(rowData['R'].Value.ToString());
        _monthSpanGesamtVorjahrMonat = CutAfterOneNumber(ToPercent(rowData['S'].Value.ToString())) + '%';
        _offeneLeistungskostenEigen = CutAfterOneNumber(rowData['U'].Value.ToString());
        _offeneLeistungskostenFremd = CutAfterOneNumber(rowData['V'].Value.ToString());
        _offeneLeistungskostenGesamt = CutAfterOneNumber(rowData['W'].Value.ToString());
    }

    private static char ToAsciiLetter(int num)
    {
        return (char)(num + 64);
    }

    private static int GetRowIndex(Worksheet worksheet, string? user)
    {
        if (string.IsNullOrEmpty(user))
            throw new Exception("The username can't be empty! Exit.");
        Console.WriteLine($"max col: {worksheet.Cells.MaxDataRow}");
        for (int i = 0; i <= worksheet.Cells.MaxDataRow; i++)
        {
            Console.WriteLine($"cell: {(string)worksheet.Cells[i, 0].Value}");
            if ((string)worksheet.Cells[i, 0].Value == user)
                return i;
        }
        throw new Exception($"The user [{user}] does not exist. Exit");
    }

    private static string ToPercent(string number)
    {
        if (string.IsNullOrEmpty(number))
        {
            Console.WriteLine("Input in ToPercent is empty or null. Continue.");
            return number;
        }

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
        if (string.IsNullOrEmpty(number) || number.StartsWith("."))
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
            if (i == 10 || i == 19)
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
        result += "<title>Hallo " + _firm + "!</title></head><body><div class=\"table-container\"><table><thead>\n<tr><th colspan=\"8\" class=\"main-header\">";
        result += _currentMonth + "</th><th colspan=\"8\" class=\"main-header\">";
        result += _monthSpan;
        result += File.ReadAllText("html/headingTitles.html");
        result += GetStyleDiv(_currentMonthBeaTEUR, "yellow");
        result += GetStyleDiv(_currentMonthBeaVorjahrMonat, "yellow");
        result += GetStyleDiv(_currentMonthGFTEUR, "green");
        result += GetStyleDiv(_currentMonthGFVorjahrMonat, "green");
        result += GetStyleDiv(_currentMonthTaxTEUR, "blue");
        result += GetStyleDiv(_currentMonthTaxVorjahrMonat, "blue");
        result += GetStyleDiv(_currentMonthGesamtTEUR, "purple");
        result += GetStyleDiv(_currentMonthGesamtVorjahrMonat, "purple");
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
        string emailHost = Environment.GetEnvironmentVariable("EMAIL_HOST");
        string emailPassword = Environment.GetEnvironmentVariable("EMAIL_PASSWORD");

        message.From.Add(new MailboxAddress("test name", emailHost));
        // Console.WriteLine($"EMAIL TO: {_emailTo}");
        message.To.Add(new MailboxAddress(_firm, _emailTo));
        message.Subject = "Wochenstatistik für " + _firm;

        message.Body = new TextPart("html")
        { Text = "Hallo " + _firm + ",\nHier Ihre Wochenstatistik:\n"+ GetHtmlContent() };

        string host = Environment.GetEnvironmentVariable("EMAIL_SERVER_HOST");
        string port = Environment.GetEnvironmentVariable("EMAIL_SERVER_PORT");
        string ssl = Environment.GetEnvironmentVariable("EMAIL_SERVER_SSL");

        using (var client = new SmtpClient ()) {
            client.Connect (host, int.Parse(port), bool.Parse(ssl));
            client.Authenticate(emailHost, emailPassword);
            client.Send (message);
            client.Disconnect (true);
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

    public static Dictionary<string, string> GetUserDic(string user_file_path, Worksheet worksheet)
    {
        Dictionary<string, string> userDic = new Dictionary<string, string>();
        using (StreamReader file = new StreamReader(user_file_path))
        {
            string line;
            while ((line = file.ReadLine()) != null)
            {
                if (line.Length == 0 || line.All(c => c == ' ' || c == '\t'))
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

                        // Console.WriteLine($"cell: {cellValue}");
                        if (cellValue == firm)
                        {
                            Console.WriteLine($"FOUND FIRM: {cellValue}");
                            found = true;
                            break;
                        }
                    }
                    if (found == false)
                        throw new Exception($"Wrong syntax in User_Wochenstatistik file. Firm is invalid. Found: {firm}. Exit.");
                    else
                        userDic[firm] = email;
                    Console.WriteLine($"{email}|{firm}");
                }
                else
                    throw new Exception($"Wrong syntax in User_Wochenstatistik file. Expected: <email>|<firm>. Found: {line}. Exit.");
            }
        }
        return userDic;
    }
}
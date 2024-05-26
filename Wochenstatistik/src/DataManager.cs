using Aspose.Cells;
using MimeKit;
using MailKit.Net.Smtp;
namespace Wochenstatistik;

public static class DataManager
{

    public static char ToASCIILetter(int num)
    {
        return (char)(num + 64);
    }

    public static int GetRowIndex(Worksheet worksheet, string? user)
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

    private static object ToPercent(object val)
    {
        if (val is double)
        {
            double round = Math.Round( (double)val * 100, 2, MidpointRounding.AwayFromZero );
            return round;
        }
        return val;
    }
    public static Dictionary<char, Cell> GetDataFromRowAsArray(Worksheet worksheet, int rowIndex)
    {
        Row row = worksheet.Cells.Rows[rowIndex];
        Dictionary<char, Cell> dictionary = new Dictionary<char, Cell>();

        for (int i = 2; i <= worksheet.Cells.MaxDataColumn; i++)
        {
            if (i == 10 || i == 19)
                continue;
            char letter = ToASCIILetter(i + 1);
            dictionary[ToASCIILetter(i+1)] = row[i];
            if (letter == 'D' || letter == 'F' || letter == 'H' || letter == 'J' || letter == 'M' || letter == 'O' ||
                letter == 'Q' || letter == 'S')
                dictionary[ToASCIILetter(i+1)].Value = ToPercent(row[i].Value);
            // Console.WriteLine($"Put {dictionary[ToASCIILetter(i+1)].Value} at POS {ToASCIILetter(i+1)}");
        }
        return dictionary;
    }

    private static string CutAfterOneNumber(string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;
        int commaIndex = input.IndexOf('.');
        if (commaIndex == -1 || commaIndex == input.Length - 1)
            return input;
        int length = Math.Min(commaIndex + 2, input.Length);

        char numberAfterComma = input[commaIndex + 1];
        if (numberAfterComma == '0')
            return input.Substring(0, length-2);
        else
            return input.Substring(0, length);
    }

    private static string GetHtmlContent(Dictionary<char, Cell> rowData)
    {
        string result = File.ReadAllText("html/styling.html");
        result += "<title>Excel Table</title></head><body><div class=\"table-container\"><table><thead>\n<tr><th colspan=\"8\" class=\"main-header\">";
        string CurrentMonth = "Mai";
        result += CurrentMonth;
        result += "</th><th colspan=\"8\" class=\"main-header\">";
        string MonthSpan = "Januar - Mai";
        result += MonthSpan;
        result += File.ReadAllText("html/headingTitles.html");

        string CurrentMonthBeaTEUR = CutAfterOneNumber(rowData['C'].Value.ToString());
        if (CurrentMonthBeaTEUR.StartsWith("-"))
            result +="<td class=\"highlight-yellow\" style=\"color: red;\">";
        else
            result += "<td class=\"highlight-yellow\">";
        result += CurrentMonthBeaTEUR;

        string CurrentMonthBeaVorjahrMonat = CutAfterOneNumber(rowData['D'].Value.ToString());
        if (CurrentMonthBeaVorjahrMonat.StartsWith("-"))
            result +="</td><td class=\"highlight-yellow\" style=\"color: red;\">";
        else
            result += "</td><td class=\"highlight-yellow\">";
        result += CurrentMonthBeaVorjahrMonat;

        string CurrentMonthGFTEUR = CutAfterOneNumber(rowData['E'].Value.ToString());
        if (CurrentMonthGFTEUR.StartsWith("-"))
            result += "</td><td class=\"highlight-green\" style=\"color: red;\">";
        else
            result += "</td><td class=\"highlight-green\">";
        result += CurrentMonthGFTEUR;


        string CurrentMonthGFVorjahrMonat = CutAfterOneNumber(rowData['F'].Value.ToString());
        if (CurrentMonthGFVorjahrMonat.StartsWith("-"))
            result += "</td><td class=\"highlight-green\" style=\"color: red;\">";
        else
            result += "</td><td class=\"highlight-green\">";
        result += CurrentMonthGFVorjahrMonat;

        string CurrentMonthTaxTEUR = CutAfterOneNumber(rowData['G'].Value.ToString());
        if (CurrentMonthTaxTEUR.StartsWith("-"))
            result += "</td><td class=\"highlight-blue\" style=\"color: red;\">";
        else
            result += "</td><td class=\"highlight-blue\">";
        result += CurrentMonthTaxTEUR;

        string CurrentMonthTaxVorjahrMonat = CutAfterOneNumber(rowData['H'].Value.ToString());
        if (CurrentMonthTaxTEUR.StartsWith("-"))
            result += "</td><td class=\"highlight-blue\" style=\"color: red;\">";
        else
            result += "</td><td class=\"highlight-blue\">";
        result += CurrentMonthTaxVorjahrMonat;

        string CurrentMonthGesamtTEUR = CutAfterOneNumber(rowData['I'].Value.ToString());
        if (CurrentMonthGesamtTEUR.StartsWith("-"))
            result += "</td><td class=\"highlight-purple\" style=\"color: red;\">";
        else
            result += "</td><td class=\"highlight-purple\">";
        result += CurrentMonthGesamtTEUR;

        string CurrentMonthGesamtVorjahrMonat = CutAfterOneNumber(rowData['J'].Value.ToString());
        if (CurrentMonthGesamtVorjahrMonat.StartsWith("-"))
            result += "</td><td class=\"highlight-purple\" style=\"color: red;\">";
        else
            result += "</td><td class=\"highlight-purple\">";
        result += CurrentMonthGesamtVorjahrMonat;

        string MonthSpanBeaTEUR = CutAfterOneNumber(rowData['L'].Value.ToString());
        if (MonthSpanBeaTEUR.StartsWith("-"))
            result += "</td><td class=\"highlight-yellow\" style=\"color: red;\">";
        else
            result += "</td><td class=\"highlight-yellow\">";
        result += MonthSpanBeaTEUR;

        string MonthSpanBeaVorjahrMonat = CutAfterOneNumber(rowData['M'].Value.ToString());
        if (MonthSpanBeaVorjahrMonat.StartsWith("-"))
            result += "</td><td class=\"highlight-yellow\" style=\"color: red;\">";
        else
            result += "</td><td class=\"highlight-yellow\">";
        result += MonthSpanBeaVorjahrMonat;


        string MonthSpanGFTEUR = CutAfterOneNumber(rowData['N'].Value.ToString());
        if (MonthSpanGFTEUR.StartsWith("-"))
            result += "</td><td class=\"highlight-green\" style=\"color: red;\">";
        else
            result += "</td><td class=\"highlight-green\">";
        result += MonthSpanGFTEUR;


        string MonthSpanGFVorjahrMonat = CutAfterOneNumber(rowData['O'].Value.ToString());
        if (MonthSpanGFVorjahrMonat.StartsWith("-"))
            result += "</td><td class=\"highlight-green\" style=\"color: red;\">";
        else
            result += "</td><td class=\"highlight-green\">";
        result += MonthSpanGFVorjahrMonat;


        string MonthSpanTaxTEUR = CutAfterOneNumber(rowData['P'].Value.ToString());
        if (MonthSpanTaxTEUR.StartsWith("-"))
            result += "</td><td class=\"highlight-blue\" style=\"color: red;\">";
        else
            result += "</td><td class=\"highlight-blue\">";
        result += MonthSpanTaxTEUR;


        string MonthSpanTaxVorjahrMonat = CutAfterOneNumber(rowData['Q'].Value.ToString());
        if (MonthSpanTaxVorjahrMonat.StartsWith("-"))
            result += "</td><td class=\"highlight-blue\" style=\"color: red;\">";
        else
            result += "</td><td class=\"highlight-blue\">";
        result += MonthSpanTaxVorjahrMonat;


        string MonthSpanGesamtTEUR = CutAfterOneNumber(rowData['R'].Value.ToString());
        if (MonthSpanGesamtTEUR.StartsWith("-"))
            result += "</td><td class=\"highlight-purple\" style=\"color: red;\">";
        else
            result += "</td><td class=\"highlight-purple\">";
        result += MonthSpanGesamtTEUR;

        string MonthSpanGesamtVorjahrMonat = CutAfterOneNumber(rowData['S'].Value.ToString());
        if (MonthSpanGesamtVorjahrMonat.StartsWith("-"))
            result += "</td><td class=\"highlight-purple\" style=\"color: red;\">";
        else
            result += "</td><td class=\"highlight-purple\">";
        result += MonthSpanGesamtVorjahrMonat;


        string offeneLSEigen = CutAfterOneNumber(rowData['U'].Value.ToString());
        if (offeneLSEigen.StartsWith("-"))
            result += "</td><td style=\"color: red;\">";
        else
            result += "</td><td>";
        result += offeneLSEigen;

        string offeneLSFremd = CutAfterOneNumber(rowData['V'].Value.ToString());
        if (offeneLSFremd.StartsWith("-"))
            result += "</td><td style=\"color: red;\">";
        else
            result += "</td><td>";
        result += offeneLSFremd;

        string offeneLSGesamt = CutAfterOneNumber(rowData['W'].Value.ToString());
        if (offeneLSGesamt.StartsWith("-"))
            result += "</td><td style=\"color: red;\">";
        else
            result += "</td><td>";
        result += offeneLSGesamt;

        result += "</td></tr></tbody></table></div></body></html>";
        return result;
    }

    public static void sendMail(Dictionary<char, Cell> rowData)
    {
        var message = new MimeMessage ();
        string emailHost = Environment.GetEnvironmentVariable("EMAIL_HOST");
        string emailPassword = Environment.GetEnvironmentVariable("EMAIL_PASSWORD");

        message.From.Add(new MailboxAddress("test name", emailHost));
        message.To.Add(new MailboxAddress("", "mmensing@eisenfuhr.com"));
        message.Subject = "lol";

        message.Body = new TextPart("html")
        {
            Text = GetHtmlContent(rowData)
        };

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
}
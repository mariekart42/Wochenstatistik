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

    private static string GetHtmlContent()
    {
        string result = File.ReadAllText("html/styling.html");
        result += "<title>Excel Table</title></head><body><div class=\"table-container\"><table><thead>\n<tr><th colspan=\"8\" class=\"main-header\">";
        string CurrentMonth = "Mai";
        result += CurrentMonth;
        result += "</th><th colspan=\"8\" class=\"main-header\">";
        string MonthSpan = "Januar - Mai";
        result += MonthSpan;
        result += File.ReadAllText("html/headingTitles.html");
        result += "<td class=\"highlight-yellow\" style=\"\">";
        string CurrentMonthBeaTEUR = "3,2";
        result += CurrentMonthBeaTEUR;
        result += "</td><td class=\"highlight-yellow red-text\" style=\"\">";
        string CurrentMonthBeaVorjahrMonat = "-95,2%";
        result += CurrentMonthBeaVorjahrMonat;
        result += "</td><td class=\"highlight-green\" style=\"\">";
        string CurrentMonthGFTEUR = "0,5";
        result += CurrentMonthGFTEUR;
        result += "</td><td class=\"highlight-green red-text\" style=\"\">";
        string CurrentMonthGFVorjahrMonat = "-60,2%";
        result += CurrentMonthGFVorjahrMonat;
        result += "</td><td class=\"highlight-blue\" style=\"\">";
        string CurrentMonthTaxTEUR = "0,0";
        result += CurrentMonthTaxTEUR;
        result += "</td><td class=\"highlight-blue red-text\" style=\"color: red;\">";
        string CurrentMonthTaxVorjahrMonat = "-100%";
        result += CurrentMonthTaxVorjahrMonat;
        result += "</td><td class=\"highlight-purple\" style=\"\">";
        string CurrentMonthGesamtTEUR = "3,7";
        result += CurrentMonthGesamtTEUR;
        result += "</td><td class=\"highlight-purple red-text\" style=\"\">";
        string CurrentMonthGesamtVorjahrMonat = "-94,6%";
        result += CurrentMonthGesamtVorjahrMonat;
        result += "</td><td class=\"highlight-yellow\" style=\"\">";
        string MonthSpanBeaTEUR = "197,5";
        result += MonthSpanBeaTEUR;
        result += "</td><td class=\"highlight-yellow red-text\" style=\"\">";
        string MonthSpanBeaVorjahrMonat = "-45,5%";
        result += MonthSpanBeaVorjahrMonat;
        result += "</td><td class=\"highlight-green\" style=\"\">";
        string MonthSpanGFTEUR = "26,9";
        result += MonthSpanGFTEUR;
        result += "</td><td class=\"highlight-green red-text\" style=\"\">";
        string MonthSpanGFVorjahrMonat = "-36,3%";
        result += MonthSpanGFVorjahrMonat;
        result += "</td><td class=\"highlight-blue\" style=\"\">";
        string MonthSpanTaxTEUR = "1,3";
        result += MonthSpanTaxTEUR;
        result += "</td><td class=\"highlight-blue red-text\" style=\"\">";
        string MonthSpanTaxVorjahrMonat = "27,1%";
        result += MonthSpanTaxVorjahrMonat;
        result += "</td><td class=\"highlight-purple\" style=\"\">";
        string MonthSpanGesamtTEUR = "225,7";
        result += MonthSpanGesamtTEUR;
        result += "</td><td class=\"highlight-purple red-text\" style=\"\">";
        string MonthSpanGesamtVorjahrMonat = "-44,3%";
        result += MonthSpanGesamtVorjahrMonat;
        result += "</td><td style=\"\">";
        string offeneLSEigen = "22,6";
        result += offeneLSEigen;
        result += "</td><td style=\"\">";
        string offeneLSFremd = "4,8";
        result += offeneLSFremd;
        result += "</td><td style=\"\">";
        string offeneLSGesamt = "27,4";
        result += offeneLSGesamt;
        result += "</td></tr>";
        result += "</tbody></table></div></body></html>";
        return result;
    }

    public static void sendMail()
    {
        var message = new MimeMessage ();
        string emailHost = Environment.GetEnvironmentVariable("EMAIL_HOST");
        string emailPassword = Environment.GetEnvironmentVariable("EMAIL_PASSWORD");

        message.From.Add(new MailboxAddress("test name", emailHost));
        message.To.Add(new MailboxAddress("", "mmensing@eisenfuhr.com"));
        message.Subject = "lol";

        message.Body = new TextPart("html")
        {
            Text = GetHtmlContent()
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
using Aspose.Cells;
using Wochenstatistik;
using System;
using System.Globalization;
using MimeKit;
using MailKit.Net.Smtp;

string path = Environment.GetEnvironmentVariable("DOCUMENT_PATH");
string? user = "PAP";

try
{
    Console.WriteLine($"PATH {path}");
    Worksheet worksheet = ExcelHandler.GetWorksheet(path);
    int rowIndex = DataManager.GetRowIndex(worksheet, user);
    Dictionary<char, Cell> rowData = DataManager.GetDataFromRowAsArray(worksheet, rowIndex);

    for (int i = 2; i <= worksheet.Cells.MaxDataColumn; i++)
    {
        if (i == 10 || i == 19)
            continue;
        Console.WriteLine($"[{DataManager.ToASCIILetter(i+1)}]: {rowData[DataManager.ToASCIILetter(i+1)].Value}");
    }

    var message = new MimeMessage ();
    string emailHost = Environment.GetEnvironmentVariable("EMAIL_HOST");
    string emailPassword = Environment.GetEnvironmentVariable("EMAIL_PASSWORD");

    message.From.Add(new MailboxAddress("test name", emailHost));
    message.To.Add(new MailboxAddress("", "marie.a.mensing@gmail.com"));
    message.Subject = "lol";

    message.Body = new TextPart ("plain") {
        Text = @"This is a test."
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
catch (Exception e)
{
    Console.WriteLine($"ERROR: {e}");
}


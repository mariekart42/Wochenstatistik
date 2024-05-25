using Aspose.Cells;
using Wochenstatistik;
using System;
using System.Globalization;
using Microsoft.Extensions.Configuration;
using MimeKit;
using MailKit.Net.Smtp;

var config = new ConfigurationBuilder()
    .AddUserSecrets<Program>()
    .Build();

string path = config["DOCUMENT_PATH"];
string? user = "PAP";

try
{
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
    string emailHost = config["EMAIL_HOST"];
    string emailPassword = config["EMAIL_PASSWORD"];

    message.From.Add(new MailboxAddress("test name", emailHost));
    message.To.Add(new MailboxAddress("", "marie.a.mensing@gmail.com"));
    message.Subject = "Test";

    message.Body = new TextPart ("plain") {
        Text = @"This is a test."
    };

    string host = config["EMAIL_SERVER_HOST"];
    string port = config["EMAIL_SERVER_PORT"];
    string ssl = config["EMAIL_SERVER_SSL"];

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


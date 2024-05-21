using Aspose.Cells;
using Wochenstatistik;
using System;
using System.Globalization;

string path = "/Users/mariemensing/Documents/Daten_Wochenstatistik.xlsx";
Worksheet worksheet = ExcelHandler.GetWorksheet(path);

Console.WriteLine($"worksheet name: {worksheet.Name}");

string user = "FEG";
int rowIndex = 7;

Console.WriteLine($"All cells from row: {rowIndex}");
Dictionary<char, Cell> rowData = DataManager.GetDataFromRowAsArray(worksheet, rowIndex);

for (int i = 2; i < 22; i++)
{
    if (i == 10 || i == 19)
        continue;
    Console.WriteLine($"[{DataManager.ToASCIILetter(i+1)}]: {rowData[DataManager.ToASCIILetter(i+1)].Value}");
}

var test = rowData['D'];
Console.WriteLine($"Give: {test.Value}");
double percent = DataManager.ToPercent(test);
Console.WriteLine($"Converted: {percent}");
// Console.WriteLine($"Cell D as decimal: {test}");
// var percent_test = test.ToString("P1", CultureInfo.InvariantCulture);
// Console.WriteLine($"Cell D as percent: {percent_test}");

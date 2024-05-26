﻿using Aspose.Cells;
using Wochenstatistik;
using System;
using System.Globalization;


string path = Environment.GetEnvironmentVariable("DOCUMENT_PATH");
string? user = "CMC";

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

    DataManager.sendMail(rowData);

}
catch (Exception e)
{
    Console.WriteLine($"ERROR: {e}");
}


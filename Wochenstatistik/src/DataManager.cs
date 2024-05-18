using Aspose.Cells;

namespace Wochenstatistik;

public static class DataManager
{

    public static char ToASCIILetter(int num)
    {
        return (char)(num + 64);
    }
    public static Dictionary<char, Cell> GetDataFromRowAsArray(Worksheet worksheet, int rowIndex)
    {
        Row row = worksheet.Cells.Rows[rowIndex];
        Dictionary<char, Cell> dictionary = new Dictionary<char, Cell>();

        for (int i = 2; i <= 22; i++)
        {
            if (i == 10 || i == 19)
                continue;
            dictionary[ToASCIILetter(i+1)] = row[i];
            // Console.WriteLine($"Put {dictionary[ToASCIILetter(i+1)].Value} at POS {ToASCIILetter(i+1)}");
        }

        return dictionary;
    }
}
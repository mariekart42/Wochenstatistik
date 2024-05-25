using Aspose.Cells;

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
}
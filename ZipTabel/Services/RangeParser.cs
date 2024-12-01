using System.Text.Json;
using System.Text.RegularExpressions;

namespace ZipTabel.Services
{
    public static class RangeParser
    {
            public static List<string> ParseAddressFormula(string formula)
            {
                string rangePattern = @"([A-Z]+[0-9]+:[A-Z]+[0-9]+)";
                string cellPattern = @"([A-Z]+[0-9]+)";

                List<string> matches = new List<string>();

                foreach (Match match in Regex.Matches(formula, cellPattern))
                {
                    matches.Add(match.Value);
                }

            MatchCollection rangeMatch = Regex.Matches(formula, rangePattern);
            foreach (var match in rangeMatch)
            {
                if (!matches.Contains(match))
                {

                    foreach (var item in GetAddresssInRange(match.ToString()))
                    {
                        matches.Add(item);
                      
                    }


                }
            }

            matches = matches.Distinct().ToList();
            Console.WriteLine("test "+JsonSerializer.Serialize(matches));
            return matches;
            }
        private static string[] GetAddresssInRange(string range)
        {
            var match = Regex.Match(range, @"(?<StartColumn>[A-Z]+)(?<StartRow>\d+):(?<EndColumn>[A-Z]+)(?<EndRow>\d+)");
            if (!match.Success)
            {
                return Array.Empty<string>();
            }

            // Извлекаем значения из группы
            string startColumn = match.Groups["StartColumn"].Value;
            int startRow = int.Parse(match.Groups["StartRow"].Value);
            string endColumn = match.Groups["EndColumn"].Value;
            int endRow = int.Parse(match.Groups["EndRow"].Value);

            // Убедимся, что начальный диапазон меньше конечного
            if (string.Compare(startColumn, endColumn) > 0)
            {
                (startColumn, endColumn) = (endColumn, startColumn);
            }
            if (startRow > endRow)
            {
                (startRow, endRow) = (endRow, startRow);
            }

            // Преобразование столбцов в числовые индексы
            int startColIndex = ColumnToIndex(startColumn);
            int endColIndex = ColumnToIndex(endColumn);

            // Формируем список адресов
            var addresses = new List<string>();
            for (int col = startColIndex; col <= endColIndex; col++)
            {
                for (int row = startRow; row <= endRow; row++)
                {
                    addresses.Add($"{IndexToColumn(col)}{row}");
                }
            }
            return addresses.ToArray();
        }

        // Вспомогательный метод: преобразование столбца в индекс
        private static int ColumnToIndex(string column)
        {
            int index = 0;
            foreach (char c in column)
            {
                index = index * 26 + (c - 'A' + 1);
            }
            return index;
        }

        // Вспомогательный метод: преобразование индекса в столбец
        private static string IndexToColumn(int index)
        {
            string column = "";
            while (index > 0)
            {
                index--;
                column = (char)('A' + index % 26) + column;
                index /= 26;
            }
            return column;
        }

        private static bool IsInRange(string address, string startColumn, int startRow, string endColumn, int endRow)
            {
                var match = Regex.Match(address, "(?<Column>[A-Z]+)(?<Row>\\d+)");
                if (!match.Success) return false;

                string column = match.Groups["Column"].Value;
                int row = int.Parse(match.Groups["Row"].Value);

                return string.Compare(column, startColumn, StringComparison.OrdinalIgnoreCase) >= 0 &&
                       string.Compare(column, endColumn, StringComparison.OrdinalIgnoreCase) <= 0 &&
                       row >= startRow &&
                       row <= endRow;
            }
        }
   }

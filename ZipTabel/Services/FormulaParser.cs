using System.Globalization;
using System.Data;
using System.Text.RegularExpressions;
using ZipTabel.Interfaces;


public enum MatchType
{
    LargestLessThanOrEqual = 1,  // Наибольшее значение <= lookupValue (массив по возрастанию)
    ExactMatch = 0,              // Точное совпадение
    SmallestGreaterThanOrEqual = -1  // Наименьшее значение >= lookupValue (массив по убыванию)
}

public static class ExcelFormulaEvaluator
{
    private static readonly Dictionary<string, Func<string[], List<ICell>, string>> formulaHandlers = new()
    {
        { "СУММ", (arguments, dependencies) => EvaluateSum(arguments, dependencies) },
        { "ЕСЛИ", (arguments, dependencies) => EvaluateIf(arguments, dependencies)},
        { "ПРОСМОТР", (arguments, dependencies) => EvaluateWatch(arguments, dependencies) },
        { "ПОИСКПОЗ", (arguments, dependencies) => EvaluateSearchpoz(arguments, dependencies) },
        { "ВЫБОР", (arguments, dependencies) => EvaluateChoice(arguments, dependencies) },
        { "ДАТА", (arguments, dependencies) => EvaluateDate(arguments, dependencies) },
        { "ДНИ", (arguments, dependencies) => EvaluateDay(arguments, dependencies) },
        { "НАЙТИ", (arguments, dependencies) => EvaluateFound(arguments, dependencies) },
        { "", (arguments, dependencies) => Evaluate(arguments, dependencies) }
    };
    /// <summary>
    /// Получение списка формул
    /// </summary>
    /// <returns></returns>
    public static IEnumerable<string> GetFormulaNames()
    {
        return formulaHandlers.Keys.ToList();
    }
    /// <summary>
    /// Выполняет вычисление на основе заданной формулы и ее зависимостей.
    /// </summary>
    /// <param name="formula">Строка, представляющая математическую формулу для вычисления.</param>
    /// <param name="dependencies">Список зависимостей, необходимых для вычисления формулы. Зависимости могут включать переменные или другие формулы.</param>
    /// <returns>Результат вычисления формулы в виде числа.</returns>
    /// <exception cref="NotSupportedException">Выбрасывается, если формула не поддерживается или содержит недопустимые операции.</exception>
    public static string ParseFormula(string formula, List<ICell> dependencies)
    {
        var mainformula = formula;

        try
        {
            do
            {
                var arguments = ExtractArguments(formula); 

                for (int i = 0; i < arguments.Length; i++)
                {
                    if (Regex.IsMatch(arguments[i], @"[+\-*/^]"))
                    {
                        var result = Evaluate(new string[] { arguments[i] }, dependencies);
                        mainformula = mainformula.Replace(arguments[i], result);  
                        formula = mainformula;  
                    }
                }

                for (int i = 0; i < arguments.Length; i++)
                {
                    var count = arguments[i].Count(c => c == '(' || c == ')');
                    var countformula = formula.Count(c => c == '(' || c == ')');
                    if (count == 0 && countformula == 2)
                    {
                        var formulaname = formula.StartsWith("=") ? formula.Substring(1) : formula;
                        foreach (var entry in formulaHandlers)
                        {
                            if (formulaname.StartsWith(entry.Key, StringComparison.OrdinalIgnoreCase))
                            {
                                var result = entry.Value(arguments, dependencies);
                                mainformula = mainformula.Replace(formula, result); 
                                formula = mainformula;
                                break;
                            }
                        }
                        break;
                    }
                    else if (count > 0 && countformula > 2)
                    {
                        formula = arguments[i];
                        break;
                    }
                }
            }
            while (mainformula.Contains('='));

            if (Regex.IsMatch(mainformula, @"[+\-*/^]"))
            {
                // Оцениваем и заменяем результат в основной формуле
                var result = Evaluate(new string[] { mainformula }, dependencies);
                mainformula = mainformula.Replace(mainformula, result);  // Заменяем в mainformula
                formula = mainformula;  // Обновляем формулу
            }

            return mainformula;
        }
        catch
        {
            throw new NotSupportedException($"Operator invalid: {formula}.");
        }
    }
    private static string[] ExtractArguments(string formula)
    {
        var result = new List<string>();

        int start = formula.IndexOf('(');
        int end = formula.LastIndexOf(')');
        string arg = formula.Substring(start + 1, end - start - 1);
        var arguments = arg.Split(';').Select(arg => arg.Trim()).ToArray();

        for (int i = 0; i < arguments.Length; i++)
        {
            if (arguments[i].Count(c => c == '(') == arguments[i].Count(c => c == ')'))
            {
                result.Add(arguments[i]);
            }
            else
            {
                arguments[i + 1] = arguments[i] + "; " + arguments[i + 1];
            }
        }

        return result.ToArray();
    }

    private static string Evaluate(string[] arguments, List<ICell> dependencies)
    {
        var formula = AddressToValue(arguments[0], dependencies);

        string transformedFormula = TransformPowers(formula);

        var result = new DataTable().Compute(transformedFormula, null);
        return result.ToString();
    }

    private static string EvaluateSum(string[] arguments, List<ICell> dependencies)
    {
        var rangeAddresses = ArgumentsPerAddress(arguments, dependencies);
        for (int i = 0; i < rangeAddresses.Count; i++)
        {
            rangeAddresses[i] = AddressToValue(rangeAddresses[i], dependencies);
        }

        var formula = string.Join("+", rangeAddresses);
        var result = new DataTable().Compute(formula, null);

        return result.ToString();
    }

    private static string EvaluateIf(string[] arguments, List<ICell> dependencies)
    {
        var condition = AddressToValue(arguments[0], dependencies);

        if (bool.TryParse(condition, out var booleanValue))
        {
            if (condition.ToLower() == "true")
            {
                return !string.IsNullOrEmpty(arguments[1]) ? arguments[1] : "True";
            }
            else if (condition.ToLower() == "false")
            {
                return !string.IsNullOrEmpty(arguments[2]) ? arguments[2] : "False";
            }
        }

        var match = Regex.Match(condition, @"(?<Value1>.+?)(?<Operator>>=|<=|>|<|==|!=)(?<Value2>.+)");
        if (!match.Success)
        {
            throw new ArgumentException($"Invalid condition: {arguments[0]}");
        }

        var result = EvaluateCondition(match.Groups[1].Value, match.Groups[2].Value, match.Groups[3].Value);

        if (result)
        {
            return !string.IsNullOrEmpty(arguments[1]) ? arguments[1] : "True";
        }
        else
        {
            return !string.IsNullOrEmpty(arguments[2]) ? arguments[2] : "False";
        }
    }

    private static string EvaluateWatch(string[] arguments, List<ICell> dependencies)
    {
        var MatchValue = Enum.TryParse(arguments[2], out MatchType matchType);
        var conditionrange = ExpandRange(arguments[1], dependencies);
        return FindPosition(arguments[0], conditionrange, matchType).ToString();
    }

    private static string EvaluateSearchpoz(string[] arguments, List<ICell> dependencies)
    {
        var conditionrange = ExpandRange(arguments[1], dependencies);

        foreach (var item in conditionrange)
        {
            if (EvaluateCondition(arguments[0], "==", item.Value))
            {
                var resultrange = ExpandRange(arguments[2], dependencies);
                return (conditionrange.IndexOf(item) + 1).ToString();
            }
        }

        return "#Н|Д";
    }

    private static string EvaluateChoice(string[] arguments, List<ICell> dependencies)
    {
        var rangeAddresses = ArgumentsPerAddress(arguments, dependencies);

        if (int.TryParse(arguments[0], out int argumentValue))
        {
            if (rangeAddresses.Count - 1 < argumentValue)
                return "#Н|Д";
            else
                foreach (var item in dependencies)
                {
                    if (item.Address == rangeAddresses[argumentValue])
                        return item.Value;
                }
        }

        throw new NotSupportedException($"Argument {arguments[0]} is not supported.");
    }

    private static string EvaluateDate(string[] arguments, List<ICell> dependencies)
    {
        int MaxArguments = 3;
        if (arguments.Length == MaxArguments)
        {
            // Формируем строку вида "день.месяц.год"
            string dateString = $"{arguments[0]}.{arguments[1]}.{arguments[2]}";
            return AddressToValue(dateString, dependencies);
        }
        else
        {
            throw new NotSupportedException($"Invalid date format");
        }
    }

    private static string EvaluateDay(string[] arguments, List<ICell> dependencies)
    {
        string date1 = "", date2 = "";
        foreach (var dep in dependencies)
        {
            date1 = arguments[0].Replace(dep.Address, dep.Value);
            date2 = arguments[1].Replace(dep.Address, dep.Value);
        }

        if (arguments.Length == 2)
        {
            DateTime dateTime1 = DateTime.ParseExact(date1, "dd.MM.yyyy", CultureInfo.InvariantCulture);
            DateTime dateTime2 = DateTime.ParseExact(date2, "dd.MM.yyyy", CultureInfo.InvariantCulture);

            TimeSpan difference = dateTime2 - dateTime1;

            return Math.Abs(difference.Days).ToString();
        }
        else
        {
            throw new NotSupportedException($"Invalid date format");
        }
    }

    private static string EvaluateFound(string[] arguments, List<ICell> dependencies)
    {
        string searchText = AddressToValue(arguments[0], dependencies);
        string withinText = AddressToValue(arguments[1], dependencies);
        int startPosition = 1;

        // Обработка необязательного аргумента начальной позиции
        if (arguments.Length == 3)
        {
            if (!int.TryParse(arguments[2], out startPosition) || startPosition < 1)
            {
                throw new ArgumentException("Начальная позиция должна быть числом больше или равным 1.");
            }
        }

        int adjustedStartPosition = startPosition - 1;
        if (adjustedStartPosition >= withinText.Length)
        {
            throw new ArgumentOutOfRangeException($"Начальная позиция ({startPosition}) выходит за пределы текста.");
        }

        int index = withinText.IndexOf(searchText, adjustedStartPosition, StringComparison.Ordinal);

        return index == -1 ? "0" : (index + 1).ToString();
    }

    private static string AddressToValue(string address, List<ICell> dependencies)
    {
        foreach (var dep in dependencies)
        {
            address = address.Replace(dep.Address, dep.Value);
        }

        return address;
    }

    private static List<string> ArgumentsPerAddress(string[] arguments, List<ICell> dependencies)
    {
        List<string> conditionrange = new();
        List<ICell> Buffer = new();

        foreach (var item in arguments)
        {
            Buffer = ExpandRange(item, dependencies);
            if (!Buffer.Any())
            {
                conditionrange.Add(item);
            }
            else
            {
                foreach (var item1 in Buffer)
                {
                    conditionrange.Add(item1.Value);
                }
            }
        }

        return conditionrange;
    }

    private static List<ICell> ExpandRange(string range, List<ICell> dependencies)
    {
        var match = Regex.Match(range, "(?<StartColumn>[A-Z]+)(?<StartRow>\\d+):(?<EndColumn>[A-Z]+)(?<EndRow>\\d+)");
        if (!match.Success)
        {
            return new();
        }

        string startColumn = match.Groups["StartColumn"].Value;
        int startRow = int.Parse(match.Groups["StartRow"].Value);
        string endColumn = match.Groups["EndColumn"].Value;
        int endRow = int.Parse(match.Groups["EndRow"].Value);

        if (string.Compare(startColumn, endColumn) > 0)
        {
            string tempColumn = startColumn;
            startColumn = endColumn;
            endColumn = tempColumn;

            int tempRow = startRow;
            startRow = endRow;
            endRow = tempRow;
        }

        if (startRow > endRow)
        {
            int tempRow = startRow;
            startRow = endRow;
            endRow = tempRow;
        }

        return dependencies
            .Where(cell => IsInRange(cell.Address, startColumn, startRow, endColumn, endRow))
            .Select(cell => cell)
            .ToList();
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

    private static bool EvaluateCondition(string cellValue, string @operator, string value)
    {
        cellValue = ValidateString(cellValue);
        @operator = ValidateString(@operator);
        value = ValidateString(value);

        if (double.TryParse(cellValue, out var cellNumericValue) && double.TryParse(value, out var numericValue))
        {
            return @operator switch
            {
                ">" => cellNumericValue > numericValue,
                "<" => cellNumericValue < numericValue,
                ">=" => cellNumericValue >= numericValue,
                "<=" => cellNumericValue <= numericValue,
                "==" => Math.Abs(cellNumericValue - numericValue) < 0.0001,
                "!=" => Math.Abs(cellNumericValue - numericValue) >= 0.0001,
                _ => throw new NotSupportedException($"Operator {@operator} is not supported."),
            };
        }

        return @operator switch
        {
            "==" => cellValue == value,
            "!=" => cellValue != value,
            _ => throw new NotSupportedException($"Operator {@operator} is not supported for non-numeric values."),
        };
    }

    private static int FindPosition(string lookupValue, List<ICell> range, MatchType matchType)
    {
        var filtered = FilterRange(range, lookupValue, matchType);

        if (filtered.Count == 0)
        {
            throw new ArgumentException($"Не найдено соответствующего значения в {GetMatchTypeName(matchType)} диапозоне.");
        }

        return range.IndexOf(filtered.Last());
    }

    private static List<ICell> FilterRange(List<ICell> range, string lookupValue, MatchType matchType)
    {
        switch (matchType)
        {
            case MatchType.LargestLessThanOrEqual:
                return range.Where(item => Compare(item, lookupValue) <= 0).ToList();
            case MatchType.ExactMatch:
                return range.Where(item => item.Value == lookupValue).ToList();
            case MatchType.SmallestGreaterThanOrEqual:
                return range.Where(item => Compare(item, lookupValue) >= 0).ToList();
            default:
                throw new ArgumentException("Недопустимый тип совпадения. Должно быть По Величине Меньше Или Равно, по Точному Совпадению или По Наименьшему Значению Больше Или Равно.");
        }
    }
    private static string GetMatchTypeName(MatchType matchType)
    {
        return matchType switch
        {
            MatchType.LargestLessThanOrEqual => "ascending",
            MatchType.ExactMatch => "exact",
            MatchType.SmallestGreaterThanOrEqual => "descending",
            _ => "unknown"
        };
    }


    private static int Compare(object a, object b)
    {
        if (a is double numA && b is double numB)
            return numA.CompareTo(numB);

        if (a is string strA && b is string strB)
            return string.Compare(strA, strB, StringComparison.Ordinal);

        if (a is bool boolA && b is bool boolB)
            return boolA.CompareTo(boolB);

        throw new ArgumentException("Unsupported data types for comparison.");
    }

    private static string TransformPowers(string formula)
    {
        return Regex.Replace(formula, @"(\d+)\^(\d+)", match =>
        {
            int baseNumber = int.Parse(match.Groups[1].Value);
            int exponent = int.Parse(match.Groups[2].Value);

            if (exponent <= 0)
                return "1"; 

            // Генерация строки умножений
            return string.Join("*", new string[exponent].Select(_ => baseNumber.ToString()));
        });
    }

    private static string ValidateString(string value)
    {
        int firstQuoteIndex = value.IndexOf('"');
        int lastQuoteIndex = value.LastIndexOf('"');

        if (firstQuoteIndex != -1 && lastQuoteIndex != -1 && firstQuoteIndex < lastQuoteIndex)
        {
            return value.Substring(firstQuoteIndex + 1, lastQuoteIndex - firstQuoteIndex - 1);
        }
        return value.Replace(" ", "");
    }
}
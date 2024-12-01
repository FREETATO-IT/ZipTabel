using System.Globalization;
using System.Data;
using System.Text.RegularExpressions;
using static Program;
using System.Data.SqlTypes;
using System;

public interface ICell
{
    string Address { get; }
    string Value { get; }
}
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

        { "МИН", (arguments, dependencies) => EvaluateMin(arguments, dependencies) },
        { "МАКС", (arguments, dependencies) => EvaluateMax(arguments, dependencies) },
        { "ВПР", (arguments, dependencies) => EvaluateVPR(arguments, dependencies) },
        { "СУММЕСЛИ", (arguments, dependencies) => EvaluateSumIf(arguments, dependencies) },
        { "СЧЁТЕСЛИ", (arguments, dependencies) => EvaluateScoreIf(arguments, dependencies) },
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
                        mainformula = mainformula.Replace(arguments[i], result).Replace("=","");
                        formula = mainformula;
                    }
                }

                for (int i = 0; i < arguments.Length; i++)
                {
                    var count = arguments[i].Count(c => c == '(' || c == ')');
                    var countformula = formula.Count(c => c == '(' || c == ')');
                    if (count == 0 && countformula == 2)
                    {
                        var formulaname = formula.Replace("=","");
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

    // Калькулятор
    private static string Evaluate(string[] arguments, List<ICell> dependencies)
    {
        // Преобразование степени
        string TransformPowers(string formula)
        {
            return Regex.Replace(formula, @"(\d+)\^(\d+)", match =>
            {
                int baseNumber = int.Parse(match.Groups[1].Value);
                int exponent = int.Parse(match.Groups[2].Value);

                if (exponent <= 0)
                    return "1";

                return string.Join("*", new string[exponent].Select(_ => baseNumber.ToString()));
            });
        }

        var formula = AddressToValue(arguments[0], dependencies);

        string transformedFormula = TransformPowers(formula);

        var result = new DataTable().Compute(transformedFormula, null);
        return result.ToString();
    }
    
    // Операции
    private static string EvaluateMin(string[] arguments, List<ICell> dependencies)
    {
        try
        {
            var Args = ArgumentsToAddress(arguments, dependencies);
            List<double> ValueArgs = new();
            foreach (var item in Args)
            {
                var Value = double.Parse(AddressToValue(item, dependencies));
                ValueArgs.Add(Value);
            }
            return ValueArgs.Min().ToString();
        }
        catch
        {
            throw new NotSupportedException($"Проблема в функции МИН");
        }
    }

    private static string EvaluateMax(string[] arguments, List<ICell> dependencies)
    {
        try
        {
            var Args = ArgumentsToAddress(arguments, dependencies);
            List<double> ValueArgs = new();
            foreach (var item in Args)
            {
                var Value = double.Parse(AddressToValue(item, dependencies));
                ValueArgs.Add(Value);
            }
            return ValueArgs.Max().ToString();
        }
        catch
        {
            throw new NotSupportedException($"Проблема в функции МИН");
        }
    }

    private static string EvaluateVPR(string[] arguments, List<ICell> dependencies)
    {
        try
        {
            var Oper = AddressToValue(arguments[0], dependencies);
            var Col = AddressToValue(arguments[0], dependencies);
            var Range = GetCellsInRange(arguments[2], dependencies);
            var Type = arguments[3];


            if (Type == "1" || Type == "True")
            {

            }
            else if (Type == "0" || Type == "False")
                Oper = "*" + Oper + "*";
            else
                throw new NotSupportedException($"Проблема в функции ВПР не верный тип {Type}");

            // Получаем список столбцов
            var Сolumns = new List<string>();
            foreach (var cell in Range)
            {
                string column = new string(cell.Address.TakeWhile(char.IsLetter).ToArray());
                if (!string.IsNullOrEmpty(column) && !Сolumns.Contains(column))
                {
                    Сolumns.Add(column);
                }
            }

            // Нахождение значения
            string NeedRow = null;
            for (int i = 0; i < Range.Count; i++)
            {
                if (EvaluateCondition(Range[i].Value, "LIKE", Oper))
                {
                    string rowPart = new string(Range[i].Address.SkipWhile(char.IsLetter).ToArray());
                    if (int.TryParse(rowPart, out int row))
                    {
                        NeedRow = row.ToString();
                    }
                    break;
                }
            }

            // Получаем список строк
            var Rows = new HashSet<int>();
            foreach (var cell in Range)
            {
                if (cell.Address == Сolumns[int.Parse(Col)] + NeedRow)
                {
                    return cell.Value;
                }
            }

            return "#Н/Д";
        }
        catch
        {
            throw new NotSupportedException($"Проблема в функции ВПР не верные аргументы");
        }
    }

    private static string EvaluateSumIf(string[] arguments, List<ICell> dependencies)
    {
        // Аргументы
        string range = arguments[0];       // Диапазон для условия
        string condition = arguments[1];  // Условие

        // Парсим диапазоны
        var rangeCells = GetCellsInRange(range, dependencies);
        bool rangeCellsIs = false;
        List<ICell> sumRangeCells = new();
        if (arguments.Length == 3)
        {
            sumRangeCells = GetCellsInRange(arguments[2], dependencies);
            rangeCellsIs = true;
            if (rangeCells.Count != sumRangeCells.Count)
                throw new ArgumentException("Диапазоны условия и суммирования должны быть одинакового размера.");
        }

        string conditionOperator;
        string conditionValue;

        ParseCondition(condition, out conditionOperator, out conditionValue);

        double sum = 0;

        for (int i = 0; i < rangeCells.Count; i++)
        {
            string rangeValue = rangeCells[i].Value;

            // Преобразуем значения в числа (если это возможно)
            if (double.TryParse(rangeValue, out double rangeNum))
            {
                if (EvaluateCondition(rangeNum.ToString(), conditionOperator, conditionValue))
                {
                    if (rangeCellsIs)
                    {
                        string sumValue = sumRangeCells[i].Value;
                        if (double.TryParse(sumValue, out double sumNum))
                        {
                            sum += sumNum;
                        }
                    }
                    else
                    {
                        sum += rangeNum;
                    }
                }
            }
        }

        return sum.ToString();
    }

    private static string EvaluateScoreIf(string[] arguments, List<ICell> dependencies)
    {
        // Аргументы
        string range = arguments[0];       // Диапазон для условия
        string condition = arguments[1];  // Условие

        // Парсим диапазоны
        var rangeCells = GetCellsInRange(range, dependencies);

        string conditionOperator;
        string conditionValue;

        ParseCondition(condition, out conditionOperator, out conditionValue);

        double sum = 0;

        for (int i = 0; i < rangeCells.Count; i++)
        {
            string rangeValue = rangeCells[i].Value;

            // Преобразуем значения в числа (если это возможно)
            if (double.TryParse(rangeValue, out double rangeNum) && EvaluateCondition(rangeNum.ToString(), conditionOperator, conditionValue))
            {
                sum += 1;
            }
        }

        return sum.ToString();
    }

    private static string EvaluateSum(string[] arguments, List<ICell> dependencies)
    {
        try
        {
            var rangeAddresses = ArgumentsToAddress(arguments, dependencies);
            double sum = 0;
            for (int i = 0; i < rangeAddresses.Count; i++)
            {
                sum += int.Parse(AddressToValue(rangeAddresses[i], dependencies));
            }

            return sum.ToString();
        }
        catch
        {
            throw new NotSupportedException($"Проблема в СУММ");
        }
    }

    private static string EvaluateIf(string[] arguments, List<ICell> dependencies)
    {
        string condition = AddressToValue(arguments[0], dependencies);

        string operand1;
        string conditionOperator;
        string operand2;

        ParseConditionAll(condition, out operand1, out conditionOperator, out operand2);

        if(EvaluateCondition(operand1, conditionOperator, operand2))
        {
            return arguments[1];
        }
        else
        {
            return arguments[2];
        }
    }

    private static string EvaluateWatch(string[] arguments, List<ICell> dependencies)
    {
        var MatchValue = Enum.TryParse(arguments[2], out MatchType matchType);
        var conditionrange = GetCellsInRange(arguments[1], dependencies);
        return FindPosition(arguments[0], conditionrange, matchType).ToString();
    }

    private static string EvaluateSearchpoz(string[] arguments, List<ICell> dependencies)
    {
        var conditionrange = GetCellsInRange(arguments[1], dependencies);

        foreach (var item in conditionrange)
        {
            if (EvaluateCondition(arguments[0], "==", item.Value))
            {
                var resultrange = GetCellsInRange(arguments[2], dependencies);
                return (conditionrange.IndexOf(item) + 1).ToString();
            }
        }

        return "#Н|Д";
    }

    private static string EvaluateChoice(string[] arguments, List<ICell> dependencies)
    {
        var rangeAddresses = ArgumentsToAddress(arguments, dependencies);

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

    // Адрес в значение
    private static string AddressToValue(string address, List<ICell> dependencies)
    {
        foreach (var dep in dependencies)
        {
            address = address.Replace(dep.Address, dep.Value);
        }

        return address;
    }

    // N аргументов в список адресов
    private static List<string> ArgumentsToAddress(string[] arguments, List<ICell> dependencies)
    {
        List<string> conditionrange = new();
        List<ICell> Buffer = new();

        foreach (var item in arguments)
        {
            Buffer = GetCellsInRange(item, dependencies);
            if (!Buffer.Any())
            {
                conditionrange.Add(item);
            }
            else
            {
                foreach (var item1 in Buffer)
                {
                    conditionrange.Add(item1.Address);
                }
            }
        }

        return conditionrange;
    }

    // Парсинг диапазона
    private static List<ICell> GetCellsInRange(string range, List<ICell> dependencies)
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

    // Парсинг условия oper+right
    private static void ParseCondition(string condition, out string conditionOperator, out string conditionValue)
    {
        // Логические операторы
        string[] operators = { ">=", "<=", ">", "<", "==", "!=", "True", "False" };

        // Инициализируем значения по умолчанию
        conditionOperator = string.Empty;
        conditionValue = string.Empty;

        // Если условие — это "True" или "False"
        if (condition.Equals("True", StringComparison.OrdinalIgnoreCase) ||
            condition.Equals("False", StringComparison.OrdinalIgnoreCase))
        {
            conditionOperator = "==";
            conditionValue = condition; // Сравниваем с самим логическим значением
            return;
        }

        // Ищем совпадение с логическими операторами
        foreach (var op in operators)
        {
            if (condition.StartsWith(op))
            {
                conditionOperator = op;
                conditionValue = condition.Substring(op.Length).Trim();
                return;
            }
        }

        // Проверяем, является ли условие шаблоном поиска
        if (condition.Contains("*") || condition.Contains("?") || condition.Contains("~"))
        {
            conditionOperator = "LIKE"; // Псевдооператор для шаблонов
            conditionValue = condition.Trim();
            return;
        }

        // Если нет операторов, то это может быть просто значение или адрес ячейки
        conditionOperator = "==";
        conditionValue = condition.Trim();
    }
    // Парсинг условия left+oper+right
    private static void ParseConditionAll(string condition, out string operand1, out string conditionOperator, out string operand2)
    {
        // Логические операторы
        string[] operators = { ">=", "<=", ">", "<", "==", "!=", "LIKE" };

        // Инициализация значений по умолчанию
        operand1 = string.Empty;
        conditionOperator = string.Empty;
        operand2 = string.Empty;

        // Если условие — это "True" или "False"
        if (condition.Equals("True", StringComparison.OrdinalIgnoreCase) ||
            condition.Equals("False", StringComparison.OrdinalIgnoreCase))
        {
            operand1 = "True";  // Логический операнд 1
            conditionOperator = "==";  // Оператор равенства
            operand2 = condition;  // Логическое значение
            return;
        }

        // Ищем совпадение с логическими операторами
        foreach (var op in operators)
        {
            int index = condition.IndexOf(op, StringComparison.Ordinal);
            if (index >= 0)
            {
                conditionOperator = op;
                operand1 = condition.Substring(0, index).Trim();
                operand2 = condition.Substring(index + op.Length).Trim();
                return;
            }
        }

        // Если нет операторов, это может быть просто одно значение или адрес ячейки
        operand1 = condition.Trim();
        conditionOperator = "==";  // Оператор равенства
        operand2 = operand1;  // Операнд 2 будет таким же как и операнд 1
    }
    // Исполнения условия left+per+right
    private static bool EvaluateCondition(string cellValue, string @operator, string value)
    {
        // Проверка на соответствие шаблону
        bool MatchesPattern(string text, string pattern)
        {
            // Преобразуем шаблон в регулярное выражение
            var regexPattern = "^" + Regex.Escape(pattern)
                .Replace(@"\*", ".*") // Заменяем '*' на любой набор символов
                .Replace(@"\?", ".")  // Заменяем '?' на любой один символ
                .Replace(@"~", "")    // Удаляем '~', если он используется для экранирования
                + "$";

            return Regex.IsMatch(text, regexPattern);
        }

        // Удаление лишних пробелов
        string ValidateString(string input)
        {
            return input?.Trim() ?? string.Empty;
        }

        cellValue = ValidateString(cellValue);
        @operator = ValidateString(@operator);
        value = ValidateString(value);

        // Шаблонное сравнение
        if (@operator == "LIKE")
        {
            return MatchesPattern(cellValue, value);
        }

        // Лог
        if (@operator == "==")
        {
            if (bool.TryParse(cellValue, out var cellBoolValue) && bool.TryParse(value, out var valueBool))
            {
                return cellBoolValue == valueBool;
            }
        }
        if (@operator == "!=")
        {
            if (bool.TryParse(cellValue, out var cellBoolValue) && bool.TryParse(value, out var valueBool))
            {
                return cellBoolValue != valueBool;
            }
        }

        // Мат
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

        // Текст
        return @operator switch
        {
            "==" => cellValue == value,
            "!=" => cellValue != value,
            _ => throw new NotSupportedException($"Operator {@operator} is not supported for non-numeric values."),
        };
    }
}
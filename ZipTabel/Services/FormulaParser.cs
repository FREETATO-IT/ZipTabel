using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ZipTabel.Interfaces;

namespace ZipTabel.Services
{
    public class FormulaParser
    {
        // Словарь для соответствия ключевых слов и методов
        private static readonly Dictionary<string, Func<string[], List<ICell>, string>> formulaHandlers = new()
{
    { "", (arguments, dependencies) => Evaluate(arguments, dependencies) },
    { "СУММ", (arguments, dependencies) => EvaluateSum(arguments, dependencies) },
    { "ЕСЛИ", (arguments, dependencies) => EvaluateIf(arguments, dependencies) ? "True" : "False" },
    { "ПРОСМОТР", (arguments, dependencies) => EvaluateWatch(arguments, dependencies) },
    { "ПОИСКПОЗ", (arguments, dependencies) => Evaluatesearchpoz(arguments, dependencies) },
    { "ВЫБОР", (arguments, dependencies) => EvaluateChoice(arguments, dependencies) },
    { "ДАТА", (arguments, dependencies) => EvaluateDate(arguments, dependencies) },
    { "ДНИ", (arguments, dependencies) => EvaluateDay(arguments, dependencies) },
    { "НАЙТИ", (arguments, dependencies) => EvaluateFound(arguments, dependencies) }
};

        public static string PASS(string[] arguments, List<ICell> dependencies)
        {
            // Функция ничего не делает
            return string.Empty;  // или можно просто вернуть null, если необходимо
        }

        // Главная функция для обработки формулы
        public static string ParseFormula(string formula, List<ICell> dependencies)
        {
            foreach (var entry in formulaHandlers)
            {
                if (formula.StartsWith(entry.Key, StringComparison.OrdinalIgnoreCase))
                {
                    var arguments = ExtractArguments(formula);
                    return entry.Value(arguments, dependencies);
                }
            }
            throw new NotSupportedException($"Formula {formula} is not supported.");
        }
        private static string[] ExtractArguments(string formula)
        {
            int start = formula.IndexOf('(');
            int end = formula.LastIndexOf(')');
            string args = formula.Substring(start + 1, end - start - 1);
            return args.Split(';').Select(arg => arg.Trim()).ToArray();
        }


        private static string Evaluate(string[] arguments, List<ICell> dependencies)
        {
            var formula = AddressToValue(arguments[0], dependencies);
            var result = new DataTable().Compute(formula, null);
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

        private static bool EvaluateIf(string[] arguments, List<ICell> dependencies)
        {
            var condition = AddressToValue(arguments[0], dependencies);

            if (bool.TryParse(condition, out var booleanValue))
            {
                return booleanValue;
            }

            var match = Regex.Match(condition, @"(?<Value1>.+?)(?<Operator>>=|<=|>|<|==|!=)(?<Value2>.+)");
            if (!match.Success)
            {
                throw new ArgumentException($"Invalid condition: {arguments[0]}");
            }

            var result = EvaluateCondition(match.Groups[1].Value, match.Groups[2].Value, match.Groups[3].Value);

            return result;
        }

        private static string EvaluateWatch(string[] arguments, List<ICell> dependencies)
        {
            var conditionrange = ExpandRange(arguments[1], dependencies);

            return FindPosition(arguments[0], conditionrange, arguments[2]).ToString();
        }

        private static string Evaluatesearchpoz(string[] arguments, List<ICell> dependencies)
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
            if (arguments.Length == 3)
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
            // Выделяем только те ячейки, которые входят в указанный диапазон
            // Предполагаем, что диапазон имеет формат A1:A10
            var match = Regex.Match(range, "(?<StartColumn>[A-Z]+)(?<StartRow>\\d+):(?<EndColumn>[A-Z]+)(?<EndRow>\\d+)");
            if (!match.Success)
            {
                return new();
            }

            string startColumn = match.Groups["StartColumn"].Value;
            int startRow = int.Parse(match.Groups["StartRow"].Value);
            string endColumn = match.Groups["EndColumn"].Value;
            int endRow = int.Parse(match.Groups["EndRow"].Value);

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

        public static int FindPosition(string lookupValue, List<ICell> range, string matchType)
        {
            if (matchType == "1") // Наибольшее значение <= lookupValue (массив по возрастанию)
            {
                var filtered = range
                    .Where(item => Compare(item, lookupValue) <= 0)
                    .ToList();

                if (filtered.Count == 0)
                    throw new ArgumentException("No matching value found in ascending range.");

                return range.IndexOf(filtered.Last());
            }
            else if (matchType == "0") // Точное совпадение
            {
                int index = 0;
                foreach (var item in range)
                {
                    if (item.Value == lookupValue)
                        index = range.IndexOf(item);
                }
                if (index == -1)
                    throw new ArgumentException("Exact match not found.");
                return index;
            }
            else if (matchType == "-1") // Наименьшее значение >= lookupValue (массив по убыванию)
            {
                var filtered = range
                    .Where(item => Compare(item, lookupValue) >= 0)
                    .ToList();

                if (filtered.Count == 0)
                    throw new ArgumentException("No matching value found in descending range.");

                return range.IndexOf(filtered.Last());
            }
            else
            {
                throw new ArgumentException("Invalid matchType. Must be 1, 0, or -1.");
            }
        }

        private static int Compare(object a, object b)
        {
            // Универсальное сравнение (числа, строки, логические значения)
            if (a is double numA && b is double numB)
                return numA.CompareTo(numB);

            if (a is string strA && b is string strB)
                return string.Compare(strA, strB, StringComparison.Ordinal);

            if (a is bool boolA && b is bool boolB)
                return boolA.CompareTo(boolB);

            throw new ArgumentException("Unsupported data types for comparison.");
        }
    }

}

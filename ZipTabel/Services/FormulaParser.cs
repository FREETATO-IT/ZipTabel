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
        public string Evaluate(string formula, List<ICell> dependencies)
        {
            // Пример простого парсера: поддержка "+", "-", "*", "/"
            // TODO: Расширить поддержку функций (например, SUM, AVERAGE)

            // Заменить ссылки на ячейки их значениями
            foreach (var dep in dependencies)
            {
                formula = formula.Replace(dep.Address, dep.Value);
            }

            // Вычислить значение формулы
            var result = new DataTable().Compute(formula, null);
            return result.ToString();
        }

        public List<string> ParseDependencies(string formula)
        {

            var matches = Regex.Matches(formula, @"[A-Z]+[0-9]+");
            return matches.Cast<Match>().Select(m => m.Value).ToList();
        }
    }

}

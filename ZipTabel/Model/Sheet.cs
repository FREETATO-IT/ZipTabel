using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ZipTabel.Services;

namespace ZipTabel.Model
{
    public class Sheet
    {
        public Dictionary<string, Cell> _cells; // Хранение ячеек по их адресу (A1, B2 и т.д.)

        public Sheet()
        {
            _cells = new Dictionary<string, Cell>();
           
        }

        public Cell GetCell(string address)
        {
            if (!_cells.ContainsKey(address))
            {
                _cells[address] = new Cell(address); // Создать новую ячейку, если её ещё нет
            }
            return _cells[address];
        }

       public  void AddCell(string address, Cell cell)
        {
            if (!_cells.ContainsKey(address))
            {
                _cells.Add(address, cell);
            }
        }
        public void SetCellFormula(string address, string formula)
        {
            var cell = GetCell(address);
            cell.Formula = formula;

            // Обновить зависимости
            var parser = new FormulaParser();
            var dependencies = parser.ParseDependencies(formula);

            foreach (var depAddress in dependencies)
            {
                var dependencyCell = GetCell(depAddress);
                cell.Dependencies.Add(dependencyCell);
                dependencyCell.Dependents.Add(cell);
            }

            cell.Recalculate();
        }
    }

}

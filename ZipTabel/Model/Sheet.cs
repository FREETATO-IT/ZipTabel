using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using ZipTabel.Services;

namespace ZipTabel.Model
{
    public record SheetExpoer(string Name,IEnumerable<KeyValuePair<string,Cell>> Cells);
    public class Sheet
    {
        public Dictionary<string, Cell> _cells; // Хранение ячеек по их адресу (A1, B2 и т.д.)

        public string Name { get; set; }
        public Sheet(string name)
        {
            _cells = new Dictionary<string, Cell>();
            Name = name;
           
        }

        public Cell GetCell(string address)
        {
            if (!_cells.ContainsKey(address))
            {
                _cells[address] = new Cell(address); // Создать новую ячейку, если её ещё нет
            }
            return _cells[address];
        }

        
        public string Setalize()
        {
            var NoEmpetyCell = _cells.Where(c => !string.IsNullOrEmpty(c.Value.Value));
            return JsonSerializer.Serialize(new SheetExpoer(Name,NoEmpetyCell));

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

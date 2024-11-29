using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ZipTabel.Components;
using ZipTabel.Interfaces;

namespace ZipTabel.Model
{
    public class Cell : ICell
    {
        public string Address { get; private set; } // Например, "A1"

        private string _value;
        public string Value
        {
            get => _value;
            set
            {
                _value = value;
                NotifyDependents(); // Обновить зависимые ячейки
            }
        }

        public CellSettings Settings { get;  set; }


        public string Formula { get; set; }
        public bool HasError { get; private set; }
        public List<ICell> Dependencies { get; private set; }
        public List<ICell> Dependents { get; private set; }
        public bool IsLocked { get; set; }

        public Cell(string address)
        {
            Address = address;
            Dependencies = new List<ICell>();
            Dependents = new List<ICell>();
            IsLocked = false;
        }

        // Метод пересчёта значения
        public void Recalculate()
        {
            if (string.IsNullOrEmpty(Formula)) return;

            try
            {
             
                var parser = new FormulaParser();
                Value = parser.Evaluate(Formula, Dependencies);
                HasError = false;
            }
            catch
            {
                Value = "ERROR";
                HasError = true;
            }
        }

        // Уведомить зависимые ячейки о пересчёте
        private void NotifyDependents()
        {
            foreach (var dependent in Dependents)
            {
                dependent.Recalculate();
            }
        }

     
    }

}

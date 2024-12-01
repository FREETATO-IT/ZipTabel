using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using ZipTabel.Interfaces;
using ZipTabel.Pages;
using ZipTabel.Services;

namespace ZipTabel.Model
{
    public class Cell : ICell, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _value = string.Empty;
        private string _formula = string.Empty;

        public bool IsFormula { get; set; } = false;
        public string Address { get; private set; }
        public string CellBackground { get; set; } = "#FFFFF";
        public string CellForeground { get; set; } = "#FFFFF";
        public bool HasError { get; private set; }
        public bool IsLocked { get; set; }

        public List<ICell> Dependencies { get; set; }
        public List<ICell> Dependents { get; set; }

        public string Value
        {
            get => _value;
            set
            {
                if (_value != value)
                {
                    _value = value;
                    OnPropertyChanged(nameof(Value));
                    NotifyDependents(); // Уведомляем зависимые ячейки
                }
            }
        }

        public string Formula
        {
            get => _formula;
            set
            {
                if (_formula != value)
                {
                    _formula = value;
                    OnPropertyChanged(nameof(Formula));
                    UpdateDependencies(); // Обновляем зависимости
                    Recalculate(); // Пересчитываем значение
                }
            }
        }

        public Cell(string address)
        {
            Address = address;
            Dependencies = new List<ICell>();
            Dependents = new List<ICell>();
            IsLocked = false;
        }

        public bool IsEmpty()
        {
            return string.IsNullOrEmpty(Formula) && string.IsNullOrEmpty(Value);
        }

        public void Recalculate()
        {
            Console.WriteLine($"Recalculating {Address}...");
            if (string.IsNullOrEmpty(Formula)) return;

            try
            {
                Value = ExcelFormulaEvaluator.ParseFormula(Formula, Dependencies);
                HasError = false;
            }
            catch(Exception ex) {
                Value = ex.Message.ToString();
                HasError = true;
            }
        }

        private void NotifyDependents()
        {
            foreach (var dependent in Dependents)
            {
                dependent.Recalculate();
            }
        }

        private void UpdateDependencies()
        {
            // Очистка текущих зависимостей
            foreach (var dependency in Dependencies)
            {
                dependency.Dependents.Remove(this);
            }
            Dependencies.Clear();

            // Получение новых зависимостей из формулы
            if (!string.IsNullOrEmpty(Formula))
            {
                var parsedDependencies = RangeParser.ParseAddressFormula(Formula).ToList();
                Console.WriteLine("parsedDependencies");
                foreach (var depAddress in parsedDependencies)
                {
                    var dependencyCell = Home.Sheet.GetCell(depAddress); // Метод для получения ячейки
                    if (dependencyCell != null && dependencyCell.Value!="")
                    {
                        Dependencies.Add(dependencyCell);
                        dependencyCell.Dependents.Add(this); 
                    }
                }
            }
        }

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

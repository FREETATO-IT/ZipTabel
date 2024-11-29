using System;
using System.Collections.Generic;
using System.Linq;
using ZipTabel.Interfaces;
using ZipTabel.Services;

namespace ZipTabel.Model
{
    public class Cell : ICell
    {
        public string Address { get; private set; } // Например, "A1"

        private string _value = string.Empty;
        private List<string> _charColors = new List<string>(); // Список цветов для каждого символа

        public string Value
        {
            get => _value;
            set
            {
                _value = value;
                _charColors = new List<string>(new string[value.Length].Select(_ => "#000000")); // Инициализация цветов для каждого символа
                NotifyDependents(); // Обновить зависимые ячейки
            }
        }

        public CellSettings Settings { get; set; }
        public string Formula { get; set; } = string.Empty;
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

        // Метод для установки цвета конкретного символа
        public void SetCharColor(int index, string color)
        {
            if (index >= 0 && index < _charColors.Count)
            {
                _charColors[index] = color;
            }
        }

        // Метод для получения цвета конкретного символа
        public string GetCharColor(int index)
        {
            if (index >= 0 && index < _charColors.Count)
            {
                return _charColors[index];
            }
            return "#000000"; // По умолчанию черный
        }
    }
}

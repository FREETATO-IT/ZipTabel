using System;
using System.Collections.Generic;
using System.Linq;
using ZipTabel.Interfaces;
using ZipTabel.Services;

namespace ZipTabel.Model
{
    public class Cell : ICell
    {
        public string Address { get; private set; }

        public string CellBeckground="#FFFFF";

        private string _value = string.Empty;
        private List<CharSettings> _charSettings = new List<CharSettings>(); // Список настроек для каждого символа

        public string Value
        {
            get => _value;
            set
            {
                _value = value;
                // Обновление настроек для каждого символа
                _charSettings = new List<CharSettings>(new CharSettings[value.Length].Select(_ => new CharSettings()));
                NotifyDependents(); // Обновить зависимые ячейки
            }
        }

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

        public bool IsEmpty()
        {
            return string.IsNullOrEmpty(Formula) &&
                   string.IsNullOrEmpty(Value);
        }

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

        private void NotifyDependents()
        {
            foreach (var dependent in Dependents)
            {
                dependent.Recalculate();
            }
        }

        public string GetCharColor(int index)
        {
            return _charSettings[index].Color;
        } 
        public void SetCharColor(int index,string hex)
        {
             _charSettings[index].Color = hex;
        } 
        public int GetCharSize(int index)
        {
            return _charSettings[index].FontSize;
        } 
        public void SetCharSize(int index,int size)
        {
             _charSettings[index].FontSize = size;
        }

        public void SetCharSettings(int index, string color, int fontSize, string fontFamily)
        {
            if (index >= 0 && index < _charSettings.Count)
            {
                var charSetting = _charSettings[index];
                charSetting.Color = color;
                charSetting.FontSize = fontSize;
                charSetting.FontFamily = fontFamily;
            }
        }

        public CharSettings GetCharSettings(int index)
        {
            if (index >= 0 && index < _charSettings.Count)
            {
                return _charSettings[index];
            }
            return new CharSettings(); 
        }
    }
}

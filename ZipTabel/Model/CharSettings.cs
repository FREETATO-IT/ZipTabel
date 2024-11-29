using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZipTabel.Model
{
    public class CharSettings
    {
        private int _fontSize;
        private string _color;
        private string _fontFamily;

        public int FontSize
        {
            get { return _fontSize; }
            set
            {
                if (value > 72)
                {
                    _fontSize = 72; // Ограничение на максимальный размер шрифта
                }
                else
                {
                    _fontSize = value;
                }
            }
        }

        public string Color
        {
            get { return _color; }
            set { _color = value ?? "#000000"; } 
        }

        public string FontFamily
        {
            get { return _fontFamily; }
            set { _fontFamily = value ?? "Arial"; } // По умолчанию Arial
        }

        public CharSettings()
        {
            Color = "#000000"; 
            FontSize = 12;     
            FontFamily = "Arial"; 
        }
    }

}

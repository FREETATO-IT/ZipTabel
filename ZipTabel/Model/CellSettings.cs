using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZipTabel.Model
{
    public class CellSettings
    {

        public CellSettings()
        {
            ForegroundColor = Colors.Black;
            BackgroundColor= Colors.Transparent;
            FontSize= 18;
          
        }
        public Color ForegroundColor { get; set; }
        public Color BackgroundColor { get; set; }

        
        public FontAttributes FontAttributes { get; set; }
        public TextDecorations TextDecorations { get; set; }


        private int _fontSize;

        public int FontSize
        {
            get { return _fontSize; }
            set
            {
                if (value > 72)
                {
                    _fontSize = 72; 
                }
                else
                {
                    _fontSize = value; 
                }
            }
        }



    }
}

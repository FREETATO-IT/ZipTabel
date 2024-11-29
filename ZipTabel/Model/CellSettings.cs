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

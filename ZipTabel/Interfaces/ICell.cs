using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZipTabel.Interfaces
{
    public interface ICell
    {
        // Координаты ячейки (например, A1, B2)
        string Address { get; }

        // Текущее отображаемое значение (результат вычислений)
        string Value { get; set; }

        // Формула, заданная в ячейке (например, "=A1+B1")
        string Formula { get; set; }

        // Признак ошибки в вычислении
        bool HasError { get; }

        // Список ячеек, от которых зависит эта ячейка
        List<ICell> Dependencies { get; }

        // Список ячеек, которые зависят от текущей ячейки
        List<ICell> Dependents { get; }

        // Флаг блокировки (запрет редактирования)
        bool IsLocked { get; set; }

        // Метод для обновления значения ячейки
        void Recalculate();
    }

}

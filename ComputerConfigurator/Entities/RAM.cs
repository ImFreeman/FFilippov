using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    /// <summary>
    /// Комплектующее типа "Оперативная память/RAM"
    /// </summary>
    class RAM : ComputerComponent
    {        
        /// <summary>
        /// Наличие подстветки
        /// </summary>
        public string Backlight { get; set; }

        /// <summary>
        /// Объём памяти в ГБ
        /// </summary>
        public string Memory { get; set; }

        /// <summary>
        /// Тип памяти RAM
        /// </summary>
        public string MemoryType { get; set; }

        /// <summary>
        /// Для игрового ПК
        /// </summary>
        public string ForGamingPC { get; set; }
    }
}

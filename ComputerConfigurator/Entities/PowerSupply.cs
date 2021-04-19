using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    class PowerSupply : ComputerComponent
    {
        

        /// <summary>
        /// Мощность блока
        /// </summary>
        public int Energy { get; set; }

        /// <summary>
        /// Наличие оплётки проводов
        /// </summary>
        public string WireBraiding { get; set; }

        /// <summary>
        /// Наличие подстветки
        /// </summary>
        public string Backlight { get; set; }

        /// <summary>
        /// Отстегивающиеся кабеля
        /// </summary>
        public string DetachableCables { get; set; }
    }
}

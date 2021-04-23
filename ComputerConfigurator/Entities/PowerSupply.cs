using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    /// <summary>
    /// Комплектующее типа "Блок питания"
    /// </summary>
    class PowerSupply : ComputerComponent
    {       
        /// <summary>
        /// Мощность блока
        /// </summary>
        public double Energy { get; set; }

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

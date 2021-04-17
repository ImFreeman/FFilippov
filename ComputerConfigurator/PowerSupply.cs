using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    class PowerSupply
    {
        public string Site { get; set; }
        /// <summary>
        /// Название
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Цена комплектующего
        /// </summary>
        public double Price { get; set; }

        /// <summary>
        /// Производитель
        /// </summary>
        public string Fabricator { get; set; }

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

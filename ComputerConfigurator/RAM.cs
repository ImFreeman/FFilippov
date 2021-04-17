using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    class RAM
    {
        public string Site;
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
        /// Наличие подстветки
        /// </summary>
        public string Backlight { get; set; }

        /// <summary>
        /// Объём памяти в ГБ
        /// </summary>
        public int Memory { get; set; }

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

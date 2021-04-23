using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    /// <summary>
    /// Комплектующее типа "HDD"
    /// </summary>
    class HDD : ComputerComponent
    {
        /// <summary>
        /// Объем памяти
        /// </summary>
        public string Memory { get; set; }

        /// <summary>
        /// Уровень шума
        /// </summary>
        public double LevelOfNoise { get; set; }

        /// <summary>
        /// Скорость обмена данными(Мбайт/сек)
        /// </summary>
        public double DataExchangeRate { get; set; }

        /// <summary>
        /// Объём буфера(Мб)
        /// </summary>
        public string BufferSize { get; set; }
    }
}

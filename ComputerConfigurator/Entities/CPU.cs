using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    /// <summary>
    /// Комплектующее типа "Процессор"
    /// </summary>
    class CPU : ComputerComponent
    {
        /// <summary>
        /// Количество ядер
        /// </summary>
        public string NumberOfCores { get; set; }

        /// <summary>
        /// Наличие графического ядра
        /// </summary>
        public string GraphicCore { get; set; }

        /// <summary>
        /// Тип памяти
        /// </summary>
        public string MemoryType { get; set; }

        /// <summary>
        /// Базовая частота
        /// </summary>
        public double BaseFrequency { get; set; }

        /// <summary>
        /// Наличие мультипоточности
        /// </summary>
        public string Multithreading { get; set; }

        /// <summary>
        /// Сокет процессора
        /// </summary>
        public string Socket { get; set; }

        /// <summary>
        /// Для игрового ПК
        /// </summary>
        public string ForGamingPC { get; set; }

    }
}

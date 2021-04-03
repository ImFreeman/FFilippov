using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    class CPU
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
        /// Производитель(Intel или AMD)
        /// </summary>
        public string Fabricator { get; set; }

        /// <summary>
        /// Количество ядер
        /// </summary>
        public int NumberOfCores { get; set; }

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
        public int BaseFrequency { get; set; }

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

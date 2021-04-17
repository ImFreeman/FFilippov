using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    class HDD
    {
        public string Site;
        // <summary>
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
        /// Объём памяти в ТБ
        /// </summary>
        public string Memory { get; set; }

        /// <summary>
        /// Уровень шума
        /// </summary>
        public double LevelOfNoise { get; set; }

        /// <summary>
        /// Скорость обмена данными(Мбайт/сек)
        /// </summary>
        public int DataExchangeRate { get; set; }

        /// <summary>
        /// Объём буфера(Мб)
        /// </summary>
        public string BufferSize { get; set; }
    }
}

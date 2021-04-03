using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    class SSD
    {
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
        /// Объём памяти в ГБ
        /// </summary>
        public int Memory { get; set; }

        /// <summary>
        /// Скорость записи(Мбайт/сек)
        /// </summary>
        public int WriteSpeed { get; set; }

        /// <summary>
        /// Скорость чтения(Мбайт/сек)
        /// </summary>
        public int ReadSpeed { get; set; }
    }
}

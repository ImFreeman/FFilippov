using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    class Corps
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
        /// Основной цвет
        /// </summary>
        public string MainColor { get; set; }

        /// <summary>
        /// Наличие окна на боковой стенке
        /// </summary>
        public string Window { get; set; }

        /// <summary>
        /// Подсветка
        /// </summary>
        public string Backlight { get; set; }

        /// <summary>
        /// Типоразмер корпуса
        /// </summary>
        public string FrameSize { get; set; }

        public string ForGamingPC { get; set; }
    }
}

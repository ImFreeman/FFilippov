using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    /// <summary>
    /// Комплектующее типа "SSD"
    /// </summary>
    class SSD : ComputerComponent
    {      
        /// <summary>
        /// Объём памяти в ГБ
        /// </summary>
        public double Memory { get; set; }

        /// <summary>
        /// Скорость записи(Мбайт/сек)
        /// </summary>
        public double WriteSpeed { get; set; }

        /// <summary>
        /// Скорость чтения(Мбайт/сек)
        /// </summary>
        public double ReadSpeed { get; set; }
    }
}

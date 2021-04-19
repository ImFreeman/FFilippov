using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    public abstract class ComputerComponent
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
    }
}

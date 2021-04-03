using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComputerConfigurator
{
    
    class GraphicsCard
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
        /// Рекомендуемая мощность блока питания
        /// </summary>
        public int RecommendedEnergy { get; set; }

        /// <summary>
        /// Объём памяти в ГБ
        /// </summary>
        public int Memory { get; set; }

        /// <summary>
        /// Тип памяти
        /// </summary>
        public string MemoryType { get; set; }

        /// <summary>
        /// Производитель графического процессора(AMD или NVIDIA)
        /// </summary>
        public string FabricatorOfGPU { get; set; }

        /// <summary>
        /// Количество одновременно работающих мониторов
        /// </summary>
        public int NumberOfMonitors { get; set; }

        /// <summary>
        /// Версия PCI Express
        /// </summary>
        public string PCIExpress { get; set; }

        /// <summary>
        /// Разрядность шины памяти(бит)
        /// </summary>
        public int MemoryBusWidth { get; set; }

        /// <summary>
        /// Для игрового ПК
        /// </summary>
        public string ForGamingPC { get; set; }

        /// <summary>
        /// Профессиональная видеокарта
        /// </summary>
        public string ProfessionalGraphicsCard { get; set; }
    }
}
